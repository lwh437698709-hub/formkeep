#!/usr/bin/env python3
"""Inspect, validate, and fill explicitly marked OOXML templates."""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import re
import sys
import tempfile
import unicodedata
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
ET.register_namespace("w", W)
ET.register_namespace("p", P)
ET.register_namespace("a", A)
ET.register_namespace("r", R)
SLOT_ID = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$")
PREFIX = "ot.slot:"


class TemplateError(RuntimeError):
    pass


def read_json(path: Path) -> Any:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def emit(value: Any, path: Path | None = None) -> None:
    content = json.dumps(value, ensure_ascii=False, indent=2) + "\n"
    if path:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding="utf-8")
    else:
        sys.stdout.write(content)


def digest(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            value.update(block)
    return value.hexdigest()


def kind(path: Path) -> str:
    result = {".docx": "docx", ".pptx": "pptx", ".xlsx": "xlsx"}.get(path.suffix.lower())
    if result is None:
        raise TemplateError("MVP supports only .docx, .pptx, and .xlsx")
    return result


def xml(data: bytes, part: str) -> ET.Element:
    try:
        return ET.fromstring(data)
    except ET.ParseError as exc:
        raise TemplateError(f"Invalid XML in {part}: {exc}") from exc


def valid_id(value: str) -> str:
    if not SLOT_ID.fullmatch(value):
        raise TemplateError(f"Invalid slot id: {value!r}")
    return value


def xlsx_slot(name: str) -> str | None:
    if name.startswith("ot.slot."):
        return valid_id(name[8:])
    if name.startswith("ot_slot_"):
        return valid_id(name[8:].replace("__", "."))
    return None


def discover_docx(package: zipfile.ZipFile) -> dict[str, list[dict[str, str]]]:
    found: dict[str, list[dict[str, str]]] = {}
    for part in package.namelist():
        if not (part.startswith("word/") and part.endswith(".xml")):
            continue
        root = xml(package.read(part), part)
        for node in root.iter(f"{{{W}}}sdt"):
            props = node.find(f"{{{W}}}sdtPr")
            tag = None if props is None else props.find(f"{{{W}}}tag")
            value = None if tag is None else tag.get(f"{{{W}}}val")
            if value and value.startswith(PREFIX):
                slot = valid_id(value[len(PREFIX):])
                found.setdefault(slot, []).append({"part": part, "value": value})
    return found


def discover_pptx(package: zipfile.ZipFile) -> dict[str, list[dict[str, str]]]:
    found: dict[str, list[dict[str, str]]] = {}
    for part in package.namelist():
        if not re.fullmatch(r"ppt/slides/slide\d+\.xml", part):
            continue
        root = xml(package.read(part), part)
        for shape in root.iter(f"{{{P}}}sp"):
            props = shape.find(f"./{{{P}}}nvSpPr/{{{P}}}cNvPr")
            value = None if props is None else props.get("name")
            if value and value.startswith(PREFIX):
                slot = valid_id(value[len(PREFIX):])
                found.setdefault(slot, []).append({"part": part, "value": value})
    return found


def discover_xlsx(package: zipfile.ZipFile) -> dict[str, list[dict[str, str]]]:
    part = "xl/workbook.xml"
    root = xml(package.read(part), part)
    found: dict[str, list[dict[str, str]]] = {}
    names = root.find(f"{{{S}}}definedNames")
    if names is None:
        return found
    for item in names.findall(f"{{{S}}}definedName"):
        name = item.get("name", "")
        slot = xlsx_slot(name)
        if slot is not None:
            found.setdefault(slot, []).append({"part": part, "value": name, "target": (item.text or "").strip()})
    return found


def inspect(path: Path) -> dict[str, Any]:
    file_kind = kind(path)
    with zipfile.ZipFile(path) as package:
        if file_kind == "docx":
            found, locator = discover_docx(package), "docx-content-control-tag"
        elif file_kind == "pptx":
            found, locator = discover_pptx(package), "pptx-shape-name"
        else:
            found, locator = discover_xlsx(package), "xlsx-defined-name"
    slots = []
    for slot_id, matches in sorted(found.items()):
        slots.append({
            "id": slot_id,
            "locator": {"kind": locator, "matches": matches},
            "required": True,
            "role": "unclassified",
            "content_type": "plain_text",
            "visual": {
                "status": "required",
                "page_or_slide": None,
                "object_id": None,
                "geometry": None,
                "z_order": None,
                "json_pointer": None,
                "notes": "Complete bidirectional visual mapping during template registration.",
            },
            "budget": {"min_chars": 0, "recommended_max_chars": None, "hard_max_chars": None, "max_width_units_per_line": None, "max_lines": None},
            "must_include": [],
            "fit_policy": {
                "preserve_line_count": None,
                "preferred_lines": None,
                "min_font_pt": None,
                "adaptive_bounds": None,
                "protected_neighbors": [],
                "allowed_actions": ["remove_manual_breaks", "widen", "increase_height", "uniform_font_reduce", "tighten_spacing", "compact_variant", "edit_copy"],
            },
            "capacity_states": [],
            "overflow_policy": ["widen", "increase_height", "uniform_font_reduce", "tighten_spacing", "compact_variant", "remove_redundancy", "shorten_sentences", "request_user_choice"],
            "calibration": {"status": "required", "renderer": None, "font": None, "notes": "Calibrate with the production renderer before filling."},
        })
    return {
        "format": "office-template-spec@1",
        "template_kind": file_kind,
        "template_sha256": digest(path),
        "layout": {"expected_pages_or_slides": None, "allow_count_change": False, "fail_closed": True},
        "slots": slots,
    }


def units(text: str) -> float:
    total = 0.0
    for char in text:
        if char == "\n":
            continue
        if char.isspace():
            total += 0.33
        elif unicodedata.east_asian_width(char) in {"W", "F", "A"}:
            total += 1.0
        else:
            total += 0.55
    return total


def line_estimate(text: str, width: float) -> int:
    return sum(max(1, math.ceil(units(line) / width)) for line in text.split("\n"))


def validate_spec(spec: dict[str, Any], expected_kind: str | None = None) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    if spec.get("format") != "office-template-spec@1":
        issues.append({"severity": "error", "code": "SPEC_FORMAT", "message": "Unsupported specification format"})
    if expected_kind and spec.get("template_kind") != expected_kind:
        issues.append({"severity": "error", "code": "SPEC_KIND", "message": "Template type does not match specification"})
    seen: set[str] = set()
    for index, slot in enumerate(spec.get("slots", [])):
        slot_id = slot.get("id")
        if not isinstance(slot_id, str) or not SLOT_ID.fullmatch(slot_id):
            issues.append({"severity": "error", "code": "SLOT_ID", "slot": slot_id, "message": f"Invalid slot id at index {index}"})
            continue
        if slot_id in seen:
            issues.append({"severity": "error", "code": "SLOT_DUPLICATE", "slot": slot_id, "message": "Slot id is duplicated"})
        seen.add(slot_id)
        budget = slot.get("budget", {})
        for field in ["min_chars", "recommended_max_chars", "hard_max_chars", "max_width_units_per_line", "max_lines"]:
            value = budget.get(field)
            if not isinstance(value, (int, float)) or value < 0 or (field in {"max_width_units_per_line", "max_lines"} and value <= 0):
                issues.append({"severity": "error", "code": "SLOT_UNCALIBRATED", "slot": slot_id, "field": field, "message": f"{field} must be calibrated"})
        values = [budget.get(name) for name in ("min_chars", "recommended_max_chars", "hard_max_chars")]
        if all(isinstance(value, (int, float)) for value in values) and not values[0] <= values[1] <= values[2]:
            issues.append({"severity": "error", "code": "BUDGET_ORDER", "slot": slot_id, "message": "Expected min_chars <= recommended_max_chars <= hard_max_chars"})
        if slot.get("calibration", {}).get("status") != "verified":
            issues.append({"severity": "error", "code": "CALIBRATION_REQUIRED", "slot": slot_id, "message": "Slot calibration is not verified"})
    if not seen:
        issues.append({"severity": "error", "code": "NO_SLOTS", "message": "No explicit slots were registered"})
    return issues


def validate(spec: dict[str, Any], bindings: dict[str, Any]) -> dict[str, Any]:
    issues = validate_spec(spec)
    slots = {slot["id"]: slot for slot in spec.get("slots", []) if isinstance(slot.get("id"), str)}
    for slot_id in sorted(set(bindings) - set(slots)):
        issues.append({"severity": "error", "code": "UNKNOWN_BINDING", "slot": slot_id, "message": "Binding is not declared by the template"})
    measurements: dict[str, Any] = {}
    for slot_id, slot in slots.items():
        if slot_id not in bindings:
            if slot.get("required"):
                issues.append({"severity": "error", "code": "REQUIRED_BINDING", "slot": slot_id, "message": "Required binding is missing"})
            continue
        text = bindings[slot_id]
        if not isinstance(text, str):
            issues.append({"severity": "error", "code": "BINDING_TYPE", "slot": slot_id, "message": "MVP bindings must be plain strings"})
            continue
        budget = slot.get("budget", {})
        count = len(text.replace("\n", ""))
        width = budget.get("max_width_units_per_line")
        lines = line_estimate(text, width) if isinstance(width, (int, float)) and width > 0 else None
        measurements[slot_id] = {"characters": count, "estimated_lines": lines, "visual_units": round(units(text), 2)}
        if isinstance(budget.get("min_chars"), (int, float)) and count < budget["min_chars"]:
            issues.append({"severity": "error", "code": "BELOW_MINIMUM", "slot": slot_id, "actual": count, "allowed": budget["min_chars"], "message": "Content is below its useful minimum"})
        if isinstance(budget.get("hard_max_chars"), (int, float)) and count > budget["hard_max_chars"]:
            issues.append({"severity": "error", "code": "HARD_MAXIMUM", "slot": slot_id, "actual": count, "allowed": budget["hard_max_chars"], "message": "Content exceeds the hard character maximum"})
        elif isinstance(budget.get("recommended_max_chars"), (int, float)) and count > budget["recommended_max_chars"]:
            issues.append({"severity": "warning", "code": "RECOMMENDED_MAXIMUM", "slot": slot_id, "actual": count, "allowed": budget["recommended_max_chars"], "message": "Content exceeds the comfortable maximum"})
        if lines is not None and isinstance(budget.get("max_lines"), (int, float)) and lines > budget["max_lines"]:
            issues.append({"severity": "error", "code": "ESTIMATED_LINE_OVERFLOW", "slot": slot_id, "actual": lines, "allowed": budget["max_lines"], "message": "Content exceeds the estimated line limit"})
        for phrase in slot.get("must_include", []):
            if isinstance(phrase, str) and phrase not in text:
                issues.append({"severity": "error", "code": "MUST_INCLUDE", "slot": slot_id, "value": phrase, "message": "Protected text is missing"})
    failed = any(item["severity"] == "error" for item in issues)
    return {"format": "office-template-validation@1", "status": "fail" if failed else "pass", "delivery_status": "blocked" if failed else "needs-native-render-review", "measurements": measurements, "issues": issues}


def serialise(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def replace_nodes(nodes: list[ET.Element], text: str, slot_id: str) -> None:
    if not nodes:
        raise TemplateError(f"Slot {slot_id} has no text nodes")
    nodes[0].text = text
    for node in nodes[1:]:
        node.text = ""


def patch_docx(part: str, data: bytes, bindings: dict[str, str], matched: dict[str, int]) -> bytes:
    if not (part.startswith("word/") and part.endswith(".xml")):
        return data
    root, changed = xml(data, part), False
    for node in root.iter(f"{{{W}}}sdt"):
        props = node.find(f"{{{W}}}sdtPr")
        tag = None if props is None else props.find(f"{{{W}}}tag")
        value = None if tag is None else tag.get(f"{{{W}}}val")
        if not value or not value.startswith(PREFIX) or value[len(PREFIX):] not in bindings:
            continue
        slot_id = value[len(PREFIX):]
        content = node.find(f"{{{W}}}sdtContent")
        replace_nodes([] if content is None else list(content.iter(f"{{{W}}}t")), bindings[slot_id], slot_id)
        matched[slot_id] = matched.get(slot_id, 0) + 1
        changed = True
    return serialise(root) if changed else data


def patch_pptx(part: str, data: bytes, bindings: dict[str, str], matched: dict[str, int]) -> bytes:
    if not re.fullmatch(r"ppt/slides/slide\d+\.xml", part):
        return data
    root, changed = xml(data, part), False
    for shape in root.iter(f"{{{P}}}sp"):
        props = shape.find(f"./{{{P}}}nvSpPr/{{{P}}}cNvPr")
        value = None if props is None else props.get("name")
        if not value or not value.startswith(PREFIX) or value[len(PREFIX):] not in bindings:
            continue
        slot_id = value[len(PREFIX):]
        replace_nodes(list(shape.iter(f"{{{A}}}t")), bindings[slot_id], slot_id)
        matched[slot_id] = matched.get(slot_id, 0) + 1
        changed = True
    return serialise(root) if changed else data


def sheet_map(package: zipfile.ZipFile) -> dict[str, str]:
    book = xml(package.read("xl/workbook.xml"), "xl/workbook.xml")
    rels = xml(package.read("xl/_rels/workbook.xml.rels"), "xl/_rels/workbook.xml.rels")
    targets = {item.get("Id"): item.get("Target") for item in rels.findall(f"{{{PR}}}Relationship")}
    result: dict[str, str] = {}
    sheets = book.find(f"{{{S}}}sheets")
    if sheets is None:
        return result
    for item in sheets.findall(f"{{{S}}}sheet"):
        target = targets.get(item.get(f"{{{R}}}id"))
        if target and item.get("name"):
            result[item.get("name")] = target if target.startswith("xl/") else "xl/" + target.lstrip("/")
    return result


def xlsx_targets(package: zipfile.ZipFile) -> dict[str, tuple[str, str]]:
    book = xml(package.read("xl/workbook.xml"), "xl/workbook.xml")
    sheets = sheet_map(package)
    names = book.find(f"{{{S}}}definedNames")
    result: dict[str, tuple[str, str]] = {}
    pattern = re.compile(r"^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]{1,3})\$?(\d+)$")
    if names is None:
        return result
    for item in names.findall(f"{{{S}}}definedName"):
        slot_id = xlsx_slot(item.get("name", ""))
        if slot_id is None:
            continue
        target = (item.text or "").strip()
        match = pattern.fullmatch(target)
        if not match:
            raise TemplateError(f"XLSX slot {slot_id} must point to one cell, got {target!r}")
        sheet_name = (match.group(1) or match.group(2)).replace("''", "'")
        if sheet_name not in sheets:
            raise TemplateError(f"XLSX slot {slot_id} references unknown sheet {sheet_name!r}")
        result[slot_id] = (sheets[sheet_name], f"{match.group(3)}{match.group(4)}")
    return result


def patch_xlsx(package: zipfile.ZipFile, bindings: dict[str, str]) -> tuple[dict[str, bytes], dict[str, int]]:
    targets, by_part = xlsx_targets(package), {}
    for slot_id, text in bindings.items():
        if slot_id in targets:
            part, cell = targets[slot_id]
            by_part.setdefault(part, []).append((slot_id, cell, text))
    changed, matched = {}, {}
    for part, edits in by_part.items():
        root = xml(package.read(part), part)
        for slot_id, cell_ref, text_value in edits:
            cell = root.find(f".//{{{S}}}c[@r='{cell_ref}']")
            if cell is None:
                raise TemplateError(f"XLSX slot {slot_id} target cell {cell_ref} does not exist")
            for child in list(cell):
                if child.tag in {f"{{{S}}}f", f"{{{S}}}v", f"{{{S}}}is"}:
                    cell.remove(child)
            cell.set("t", "inlineStr")
            text = ET.SubElement(ET.SubElement(cell, f"{{{S}}}is"), f"{{{S}}}t")
            text.text = text_value
            matched[slot_id] = matched.get(slot_id, 0) + 1
        changed[part] = serialise(root)
    return changed, matched


def fill(template: Path, spec: dict[str, Any], bindings: dict[str, str], output: Path, overwrite: bool, visual_book: Path | None = None) -> dict[str, Any]:
    file_kind = kind(template)
    report = validate(spec, bindings)
    issues = validate_spec(spec, file_kind) + report["issues"]
    if file_kind == "pptx":
        book = read_json(visual_book) if visual_book else {}
        if not isinstance(book, dict) or book.get("format") != "office-template-visual@2" or book.get("source_sha256") != digest(template):
            issues.append({"severity": "error", "code": "VISUAL_HANDBOOK_REQUIRED", "message": "Extract a source-bound PPTX visual handbook and pass --visual-book before filling"})
    if spec.get("template_sha256") != digest(template):
        issues.append({"severity": "error", "code": "TEMPLATE_HASH", "message": "Template bytes do not match the registered specification"})
    if any(item["severity"] == "error" for item in issues):
        return {**report, "status": "fail", "delivery_status": "blocked", "issues": issues}
    if output.exists() and not overwrite:
        raise TemplateError(f"Output already exists: {output}; pass --overwrite to replace it")
    output.parent.mkdir(parents=True, exist_ok=True)
    temp_handle = tempfile.NamedTemporaryFile(prefix=f".{output.name}.", suffix=".tmp", dir=output.parent, delete=False)
    temp_path = Path(temp_handle.name)
    temp_handle.close()
    matched: dict[str, int] = {}
    try:
        with zipfile.ZipFile(template) as source:
            xlsx_parts = {}
            if file_kind == "xlsx":
                xlsx_parts, matched = patch_xlsx(source, bindings)
            with zipfile.ZipFile(temp_path, "w") as target:
                for info in source.infolist():
                    data = source.read(info.filename)
                    if file_kind == "docx":
                        data = patch_docx(info.filename, data, bindings, matched)
                    elif file_kind == "pptx":
                        data = patch_pptx(info.filename, data, bindings, matched)
                    else:
                        data = xlsx_parts.get(info.filename, data)
                    target.writestr(info, data)
        missing = sorted(set(bindings) - set(matched))
        if missing:
            raise TemplateError(f"Registered slots were not found: {', '.join(missing)}")
        with zipfile.ZipFile(temp_path) as check:
            bad = check.testzip()
            if bad:
                raise TemplateError(f"Generated package is corrupt at {bad}")
        if file_kind == "pptx":
            from pptx_style_guard import audit
            style_report = audit(template, temp_path)
            if not style_report["structural_pass"]:
                return {**report, "status": "fail", "delivery_status": "blocked", "style_audit": style_report}
        os.replace(temp_path, output)
    finally:
        if temp_path.exists():
            temp_path.unlink()
    return {**report, "status": "pass", "delivery_status": "needs-native-render-review", "output": str(output), "output_sha256": digest(output), "filled_slots": matched, "issues": issues}


def argument_parser() -> argparse.ArgumentParser:
    root = argparse.ArgumentParser(description=__doc__)
    commands = root.add_subparsers(dest="command", required=True)
    inspect_cmd = commands.add_parser("inspect")
    inspect_cmd.add_argument("template", type=Path)
    inspect_cmd.add_argument("--to", type=Path)
    validate_cmd = commands.add_parser("validate")
    validate_cmd.add_argument("spec", type=Path)
    validate_cmd.add_argument("bindings", type=Path)
    validate_cmd.add_argument("--to", type=Path)
    fill_cmd = commands.add_parser("fill")
    fill_cmd.add_argument("template", type=Path)
    fill_cmd.add_argument("spec", type=Path)
    fill_cmd.add_argument("bindings", type=Path)
    fill_cmd.add_argument("output", type=Path)
    fill_cmd.add_argument("--report", type=Path)
    fill_cmd.add_argument("--overwrite", action="store_true")
    fill_cmd.add_argument("--visual-book", type=Path, help="Required source-bound visual handbook for PPTX")
    return root


def main() -> int:
    args = argument_parser().parse_args()
    try:
        if args.command == "inspect":
            result = inspect(args.template)
            emit(result, args.to)
            return 0 if result["slots"] else 2
        if args.command == "validate":
            spec, bindings = read_json(args.spec), read_json(args.bindings)
            if not isinstance(bindings, dict):
                raise TemplateError("Bindings must be one JSON object")
            result = validate(spec, bindings)
            emit(result, args.to)
            return 0 if result["status"] == "pass" else 2
        spec, bindings = read_json(args.spec), read_json(args.bindings)
        if not isinstance(bindings, dict) or not all(isinstance(value, str) for value in bindings.values()):
            raise TemplateError("Bindings must map slot ids to plain strings")
        result = fill(args.template, spec, bindings, args.output, args.overwrite, args.visual_book)
        emit(result, args.report)
        return 0 if result["status"] == "pass" else 2
    except (TemplateError, OSError, ValueError, KeyError, ET.ParseError, zipfile.BadZipFile) as exc:
        emit({"format": "office-template-error@1", "status": "fail", "delivery_status": "blocked", "error": str(exc)})
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
