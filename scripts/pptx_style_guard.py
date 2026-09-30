#!/usr/bin/env python3
"""Extract a PPTX visual handbook and fail closed on unregistered style changes.

Standard library only. Scores describe structural style fidelity, not perceptual
image similarity. Native-render evidence remains a separate delivery gate.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import math
import posixpath
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as E

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
R = "{" + NS["r"] + "}"
COLORS = {"srgbClr", "schemeClr", "sysClr", "scrgbClr", "hslClr", "prstClr"}
SHAPES = {"sp", "pic", "graphicFrame", "cxnSp", "grpSp"}
WEIGHTS = {"palette": 25, "typography": 20, "layout": 25, "assets": 20, "shape_style": 10}


def local(tag):
    return tag.rsplit("}", 1)[-1]


def sha(data):
    return hashlib.sha256(data).hexdigest()


def file_sha(path):
    return sha(Path(path).read_bytes())


def canon(node, ignore_text=False):
    return [node.tag, sorted(node.attrib.items()),
            "" if ignore_text or not (node.text or "").strip() else node.text,
            [canon(child, ignore_text) for child in node]]


def encoded(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def unique(values):
    return sorted({encoded(v) for v in values})


def write(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


class Package:
    def __init__(self, path):
        self.z = zipfile.ZipFile(path)
        infos = self.z.infolist()
        if len(infos) > 10000 or sum(i.file_size for i in infos) > 512 * 1024 * 1024:
            raise ValueError("Package exceeds inspection limit")
        if len({i.filename for i in infos}) != len(infos):
            raise ValueError("Duplicate package parts")

    def xml(self, part):
        data = self.z.read(part)
        if b"<!DOCTYPE" in data.upper() or b"<!ENTITY" in data.upper():
            raise ValueError("DTD/entity declarations are unsupported")
        return E.fromstring(data)

    def rels(self, part):
        name = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
        if name not in self.z.namelist():
            return {}
        result = {}
        for n in self.xml(name):
            target = n.get("Target", "")
            external = n.get("TargetMode", "Internal").lower() == "external"
            resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(part), target))
            if not external and (resolved.startswith("../") or "\\" in target):
                raise ValueError("Unsafe relationship target")
            result[n.get("Id")] = {"target": resolved, "external": external, "type": n.get("Type", "")}
        return result

    def dep(self, part, trail=()):
        if part in trail:
            return "cycle"
        if part not in self.z.namelist():
            return "missing:" + part
        raw = self.z.read(part)
        if not part.endswith(".xml"):
            return sha(raw)
        node = copy.deepcopy(self.xml(part))
        rels = self.rels(part)
        for n in node.iter():
            for key, value in list(n.attrib.items()):
                if key.startswith(R):
                    rel = rels.get(value)
                    # Master layout catalogs are not rendered content. Avoid cycles.
                    token = "layout-catalog" if rel and rel["type"].endswith("/slideLayout") else self.dep(rel["target"], (*trail, part)) if rel and not rel["external"] else "external-or-missing"
                    n.set(key, token)
        inherited = sorted(self.dep(r["target"], (*trail, part)) for r in rels.values()
                           if not r["external"] and r["type"].rsplit("/", 1)[-1] in {"slideMaster", "theme", "themeOverride"})
        return sha(encoded([canon(node), inherited]).encode())


def describe(package, part, node, key, order, parent, width, height):
    root = copy.deepcopy(node)
    # A group's own style is separate; children are inspected with parent links.
    for child in list(root):
        if local(child.tag) in SHAPES:
            root.remove(child)
    props = next((n for n in root.iter() if local(n.tag) == "cNvPr"), None)
    name = props.get("name", "") if props is not None else ""
    rels = package.rels(part)
    assets, unresolved = [], []
    for n in root.iter():
        for attr, value in list(n.attrib.items()):
            if attr.startswith(R):
                rel = rels.get(value)
                if rel and not rel["external"] and rel["target"] in package.z.namelist():
                    token = package.dep(rel["target"])
                    assets.append({"type": rel["type"], "sha256": token})
                else:
                    token = "unresolved-relationship"
                    unresolved.append(value)
                n.set(attr, token)
    xfrm = next((n for n in root.iter() if local(n.tag) == "xfrm"), None)
    box = None
    if xfrm is not None:
        off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
        if off is not None and ext is not None:
            box = [int(off.get("x", 0)), int(off.get("y", 0)), int(ext.get("cx", 0)), int(ext.get("cy", 0))]
    palette = unique(canon(n) for n in root.iter() if local(n.tag) in COLORS)
    typography = unique(canon(n) for n in root.iter() if local(n.tag) in {"rPr", "defRPr", "endParaRPr", "pPr", "bodyPr", "lstStyle", "fontRef"})
    fonts = sorted({n.get("typeface") for n in root.iter() if n.get("typeface")})
    sizes = sorted({int(n.get("sz")) / 100 for n in root.iter() if n.get("sz", "").isdigit()})
    # Ignore text content and object labels, but not effects, strokes, crops or styles.
    visual = copy.deepcopy(root)
    for n in list(visual.iter()):
        for child in list(n):
            if local(child.tag) in {"txBody", "cNvPr", "xfrm"}:
                n.remove(child)
    text = "".join(n.text or "" for n in root.iter() if local(n.tag) == "t")
    normal = None if box is None or parent else [round(box[0]/width, 6), round(box[1]/height, 6), round(box[2]/width, 6), round(box[3]/height, 6)]
    anchors = None if normal is None else {
        "left": normal[0], "right": round(normal[0]+normal[2], 6),
        "top": normal[1], "bottom": round(normal[1]+normal[3], 6),
        "center_x": round(normal[0]+normal[2]/2, 6), "center_y": round(normal[1]+normal[3]/2, 6),
    }
    return {
        "key": key, "name": name, "kind": local(root.tag), "parent": parent, "z_order": order,
        "editable": name.startswith("ot.slot:"), "text_preview": text[:120],
        "bbox_emu": box, "normalized_bbox": normal, "anchors": anchors,
        "coordinate_space": "group-local" if parent else "slide",
        "palette": palette, "typography": typography, "fonts": fonts, "font_sizes_pt": sizes,
        "shape_style": canon(visual), "assets": sorted(assets, key=encoded),
        "transform": None if xfrm is None else canon(xfrm),
        "locked_fingerprint": sha(encoded(canon(root)).encode()),
        "unresolved": unresolved,
    }


def extract(path):
    pkg = Package(path)
    try:
        pres = pkg.xml("ppt/presentation.xml")
        size = pres.find("p:sldSz", NS)
        if size is None:
            raise ValueError("Missing slide dimensions")
        width, height = int(size.get("cx")), int(size.get("cy"))
        if width <= 0 or height <= 0:
            raise ValueError("Invalid slide dimensions")
        relations = pkg.rels("ppt/presentation.xml")
        presentation_style = copy.deepcopy(pres)
        for child in list(presentation_style):
            if local(child.tag) == "sldIdLst":
                presentation_style.remove(child)
        for n in presentation_style.iter():
            for attr, value in list(n.attrib.items()):
                if attr.startswith(R):
                    rel = relations.get(value)
                    n.set(attr, pkg.dep(rel["target"]) if rel and not rel["external"] else "unresolved")
        slides, themes = [], []
        for n in pres.findall("p:sldIdLst/p:sldId", NS):
            part = relations[n.get(R+"id")]["target"]
            tree = pkg.xml(part)
            shapes = tree.find("p:cSld/p:spTree", NS)
            objects = []

            def visit(container, parent=None):
                for node in container:
                    if local(node.tag) not in SHAPES:
                        continue
                    props = next((c for c in node.iter() if local(c.tag) == "cNvPr"), None)
                    if props is None:
                        raise ValueError("Shape missing stable ID")
                    key = props.get("id")
                    objects.append(describe(pkg, part, node, key, len(objects), parent, width, height))
                    if local(node.tag) == "grpSp":
                        visit(node, key)

            if shapes is not None:
                visit(shapes)
            if len({o["key"] for o in objects}) != len(objects):
                raise ValueError("Duplicate shape ID")
            slide_rels = pkg.rels(part)
            layouts = [v["target"] for v in slide_rels.values() if v["type"].endswith("/slideLayout")]
            bg = tree.find("p:cSld/p:bg", NS)
            # Includes transition/timing and flags; content/notes are evaluated separately.
            shell = copy.deepcopy(tree)
            cs = shell.find("p:cSld", NS)
            if cs is not None:
                for child in list(cs):
                    if local(child.tag) == "spTree":
                        cs.remove(child)
            slides.append({"number": len(slides)+1, "part": part,
                           "unmodeled_tree": [canon(c) for c in shapes if local(c.tag) not in SHAPES] if shapes is not None else [],
                           "layout_dependencies": sorted(pkg.dep(l) for l in layouts),
                           "slide_shell": canon(shell), "background": None if bg is None else canon(bg),
                           "objects": objects})
        shared = []
        palette, fonts = set(), set()
        for part in pkg.z.namelist():
            if re.fullmatch(r"ppt/(theme|slideMasters|slideLayouts)/[^/]+\.xml", part):
                root = pkg.xml(part)
                shared.append(pkg.dep(part))
                for node in root.iter():
                    if local(node.tag) in COLORS:
                        palette.add(encoded(canon(node)))
                    if node.get("typeface"):
                        fonts.add(node.get("typeface"))
                if "/theme/" in part:
                    themes.append({"part": part,
                                   "color_scheme": [canon(x) for x in root.findall("a:themeElements/a:clrScheme", NS)],
                                   "font_scheme": [canon(x) for x in root.findall("a:themeElements/a:fontScheme", NS)]})
        for s in slides:
            for o in s["objects"]:
                palette.update(o["palette"])
                fonts.update(o["fonts"])
        if not slides:
            raise ValueError("No slides to inspect")
        return {"format": "office-template-visual@2", "source_sha256": file_sha(path),
                "presentation_style": canon(presentation_style),
                "slide_size_emu": [width, height], "shared_style_fingerprints": sorted(shared),
                "themes": themes, "palette_tokens": sorted(palette), "font_families": sorted(fonts),
                "direct_slide_fonts": sorted({f for s in slides for o in s["objects"] for f in o["fonts"]}),
                "font_availability": "not-checked", "font_fallbacks": [],
                "slides": slides, "registration_status": "requires-render-and-role-review",
                "limits": ["Theme tokens and transforms are preserved, not flattened to RGB.",
                           "Theme/layout inheritance is captured in dependency fingerprints.",
                           "Group-local boxes need composed transforms for UI hit-testing.",
                           "No font installation or native rendering is inferred."]}
    finally:
        pkg.z.close()


def handbook(book):
    lines = ["# 模板视觉说明书（提取草稿）", "", f"模板 SHA-256：`{book['source_sha256']}`", "",
             "必须补充角色标注、字体可用性、替代字体实测、图标样式分类及原生渲染验收。", "",
             "## 配色与字体", "", "配色来源包含主题、母版、布局与页面直接样式；主题色及明暗/透明度变换按原样保存。",
             "照片中的颜色不可自动成为新增图形配色。跨页复制色彩前需确认本页角色，不能任取主题色。", "",
             "页面直接声明的字体/主题别名：" + "、".join(book["direct_slide_fonts"]), "",
             "主题和样式中另外声明的字体（并非都实际使用）：" + "、".join(sorted(set(book["font_families"])-set(book["direct_slide_fonts"]))), "",
             "替代建议仅为候选：中文无衬线可评估 Noto Sans CJK SC / 思源黑体；中文衬线可评估 Noto Serif CJK SC / 思源宋体。",
             "书法/品牌字体无通用等价替代。必须确认安装、字符覆盖、字号字重及换行后才可登记替代。", "",
             "详细主题色槽、原始样式、素材哈希与对象字段见同名 JSON。", "",
             "|原始色值/色槽|附加变换|", "|---|---|"]
    for value in book["palette_tokens"]:
        node = json.loads(value)
        attrs = dict(node[1])
        color = attrs.get("val", attrs.get("lastClr", ""))
        if local(node[0]) == "sysClr":
            color += " / lastClr=" + attrs.get("lastClr", "")
        transforms = "; ".join(local(c[0])+":"+encoded(dict(c[1])) for c in node[3]) or "无"
        lines.append(f"|{local(node[0])}: `{color}`|{transforms}|")
    for s in book["slides"]:
        lines += ["", f"## 第 {s['number']} 页", "", "|对象 ID / 名称|属性|边界 EMU|左右锚点（0–1）|字体|", "|---|---|---|---|---|"]
        for o in s["objects"]:
            a = o["anchors"]
            anchors = f"{a['left']} / {a['right']}" if a else "需解析组合或继承位置"
            name = o["name"].replace("|", "\\|").replace("\n", " ")
            lines.append(f"|{o['key']} / {name}|{'可填充' if o['editable'] else '锁定'}|{o['bbox_emu']}|{anchors}|{'、'.join(o['fonts']) or '继承，见布局/主题'}|")
        lines += ["", "登记项：标题基线、正文左右边界、栏间距、页脚保护区、图标描边/填充、图表系列颜色角色、允许扩展范围。"]
    return "\n".join(lines) + "\n"


def geometry_ok(before, after, rule):
    if before["parent"] != after["parent"]:
        return False
    if before["transform"] == after["transform"]:
        return True
    a, b = before["bbox_emu"], after["bbox_emu"]
    bounds = rule.get("max_delta_emu")
    if not a or not b or not isinstance(bounds, list) or len(bounds) != 4:
        return False
    if not all(isinstance(v, (int, float)) and math.isfinite(v) and v >= 0 for v in bounds):
        return False
    if not all(abs(x-y) <= m for x, y, m in zip(a, b, bounds)):
        return False
    # Rotation, flip and group transforms cannot be smuggled through box bounds.
    def rest(transform):
        value = copy.deepcopy(transform)
        if value:
            value[3] = [c for c in value[3] if local(c[0]) not in {"off", "ext"}]
        return value
    return rest(before["transform"]) == rest(after["transform"])


def typography_ok(before, after, rule):
    if before == after:
        return True
    scale = rule.get("font_scale", 1)
    floor = rule.get("min_font_pt")
    if not isinstance(scale, (int, float)) or not math.isfinite(scale) or not 0 < scale <= 1:
        return False
    if scale != 1 and (not isinstance(floor, (int, float)) or not math.isfinite(floor) or floor <= 0):
        return False
    substitutions = rule.get("font_fallbacks", {})
    if substitutions and not rule.get("fallback_evidence"):
        return False
    changed = []
    valid = True

    def transform(node):
        nonlocal valid
        for attr in node[1]:
            if attr[0] == "typeface" and attr[1] in substitutions:
                attr[1] = substitutions[attr[1]]
            if attr[0] == "sz" and scale != 1:
                attr[1] = str(round(int(attr[1])*scale))
                if int(attr[1])/100 < floor:
                    valid = False
        for child in node[3]:
            transform(child)
    for value in before:
        node = json.loads(value)
        transform(node)
        changed.append(node)
    return valid and unique(changed) == after


def audit(source, output, plan=None, review=None):
    base, current = extract(source), extract(output)
    plan = plan or {}
    if not isinstance(plan, dict) or not isinstance(plan.get("objects", {}), dict):
        raise ValueError("Plan and objects must be JSON objects")
    if any(not isinstance(v, dict) for v in plan.get("objects", {}).values()):
        raise ValueError("Each object rule must be a JSON object")
    if review is not None and not isinstance(review, dict):
        raise ValueError("Render review must be a JSON object")
    issues, scores = [], []

    def fail(code, slide=None, obj=None, detail=None):
        issues.append({"code": code, "slide": slide, "object_id": obj, "detail": detail})

    if plan and plan.get("source_sha256") != base["source_sha256"]:
        fail("PLAN_SOURCE_MISMATCH")
    if base["slide_size_emu"] != current["slide_size_emu"]:
        fail("SLIDE_SIZE_CHANGED")
    if base["presentation_style"] != current["presentation_style"]:
        fail("PRESENTATION_DEFAULTS_OR_EMBEDDED_FONTS_CHANGED")
    mapping = plan.get("slide_map", list(range(1, len(current["slides"])+1)))
    if not isinstance(mapping, list):
        raise ValueError("slide_map must be a list of source slide numbers")
    if not plan.get("slide_map") and len(base["slides"]) != len(current["slides"]):
        fail("SLIDE_COUNT_CHANGED")
    if len(mapping) != len(current["slides"]) or any(type(i) is not int or not 1 <= i <= len(base["slides"]) for i in mapping):
        fail("INVALID_SLIDE_MAP")
        mapping = []
    for out_slide, source_index in zip(current["slides"], mapping):
        before = base["slides"][source_index-1]
        num = out_slide["number"]
        if before["layout_dependencies"] != out_slide["layout_dependencies"]:
            fail("MASTER_LAYOUT_THEME_CHANGED", num)
        if before["slide_shell"] != out_slide["slide_shell"]:
            fail("BACKGROUND_OR_SLIDE_SHELL_CHANGED", num)
        if before["unmodeled_tree"] != out_slide["unmodeled_tree"]:
            fail("UNMODELED_TREE_CHANGED", num)
        objects = {o["key"]: o for o in before["objects"]}
        afters = {o["key"]: o for o in out_slide["objects"]}
        hits = {k: 0 for k in WEIGHTS}
        total = max(len(set(objects) | set(afters)), 1)
        if not objects and not afters:
            hits = {k: 1 for k in WEIGHTS}
        # Relative order of existing elements is invariant, even when inserting clones.
        old_order = [o["key"] for o in before["objects"] if o["key"] in afters]
        new_order = [o["key"] for o in out_slide["objects"] if o["key"] in objects]
        if old_order != new_order:
            fail("Z_ORDER_CHANGED", num)
        for key in sorted(set(objects) | set(afters)):
            a, b = objects.get(key), afters.get(key)
            rule = plan.get("objects", {}).get(f"{num}:{key}", {})
            addition = False
            if b is None:
                fail("OBJECT_DELETED", num, key)
                continue
            if a is None:
                addition = True
                donor = rule.get("donor")
                # MVP clones only an existing leaf on the same source page.
                a = objects.get(str(donor)) if donor is not None else None
                if not a or a["kind"] == "grpSp" or not rule.get("purpose"):
                    fail("UNREGISTERED_ADDITION", num, key, "New elements require a same-page donor and purpose")
                    continue
            if a["unresolved"] or b["unresolved"]:
                fail("UNRESOLVED_RESOURCE", num, key)
            matches = {
                "palette": a["palette"] == b["palette"],
                "typography": typography_ok(a["typography"], b["typography"], rule),
                "layout": geometry_ok(a, b, rule),
                "assets": a["assets"] == b["assets"],
                "shape_style": a["kind"] == b["kind"] and a["shape_style"] == b["shape_style"],
            }
            for dimension, matches_dimension in matches.items():
                hits[dimension] += int(matches_dimension)
                if not matches_dimension:
                    fail(dimension.upper() + "_DRIFT", num, key)
            if not addition:
                if a["name"] != b["name"]:
                    fail("OBJECT_IDENTITY_CHANGED", num, key)
                if not a["editable"] and a["locked_fingerprint"] != b["locked_fingerprint"]:
                    fail("LOCKED_OBJECT_CHANGED", num, key)
            else:
                if rule.get("bbox_emu") != b["bbox_emu"]:
                    fail("ADDITION_ANCHOR_MISMATCH", num, key)
        dims = {k: round(100*v/total, 2) for k, v in hits.items()}
        score = round(sum(dims[k]*w/100 for k, w in WEIGHTS.items()), 2)
        if score < 90:
            fail("STYLE_SCORE_BELOW_90", num)
        scores.append({"slide": num, "source_slide": source_index, "dimensions": dims, "score": score})
    structural = not issues
    visual_ok = False
    if review:
        if review.get("output_sha256") != current["source_sha256"] or review.get("source_sha256") != base["source_sha256"]:
            fail("RENDER_EVIDENCE_HASH_MISMATCH")
        elif not review.get("renderer") or review.get("native_office_render") is not True:
            fail("NATIVE_RENDER_REQUIRED")
        else:
            reviewed = review.get("slides", [])
            expected = list(range(1, len(current["slides"])+1))
            if sorted(r.get("slide", 0) for r in reviewed) != expected:
                fail("INCOMPLETE_VISUAL_REVIEW")
            checks = ["no_overflow", "no_occlusion", "fonts_verified", "style_roles_match", "anchors_match"]
            for item in reviewed:
                if any(item.get(c) is not True for c in checks):
                    fail("VISUAL_REVIEW_FAILED", item.get("slide"))
                for name in ["baseline_render", "output_render"]:
                    ev = item.get(name, {})
                    p = Path(ev.get("path", ""))
                    if not p.is_file() or ev.get("sha256") != file_sha(p):
                        fail("RENDER_IMAGE_MISSING_OR_CHANGED", item.get("slide"), detail=name)
            visual_ok = not issues
    return {"format": "office-template-style-audit@2", "status": "fail" if issues else "pass",
            "delivery_status": "blocked" if issues else "deliverable" if visual_ok else "needs-native-render-review",
            "source_sha256": base["source_sha256"], "output_sha256": current["source_sha256"],
            "score_kind": "structural-style-fidelity-not-perceptual-similarity",
            "weights": WEIGHTS, "threshold_per_slide": 90,
            "deck_score": round(sum(s["score"] for s in scores)/len(scores), 2) if scores else None,
            "slides": scores, "structural_pass": structural, "visual_review_pass": visual_ok,
            "issues": issues, "claim_boundary": "Native-render review is a recorded attestation with hashed image evidence, not automatic proof of appearance."}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    ex = sub.add_parser("extract")
    ex.add_argument("source", type=Path)
    ex.add_argument("--to", required=True, type=Path)
    au = sub.add_parser("audit")
    au.add_argument("source", type=Path)
    au.add_argument("output", type=Path)
    au.add_argument("--plan", type=Path)
    au.add_argument("--review", type=Path)
    au.add_argument("--to", required=True, type=Path)
    args = parser.parse_args()
    try:
        if args.command == "extract":
            result = extract(args.source)
            write(args.to, result)
            args.to.with_suffix(".md").write_text(handbook(result), encoding="utf-8")
            print(json.dumps({"slides": len(result["slides"]), "json": str(args.to), "status": result["registration_status"]}))
            return 0
        plan = json.loads(args.plan.read_text()) if args.plan else None
        review = json.loads(args.review.read_text()) if args.review else None
        result = audit(args.source, args.output, plan, review)
        write(args.to, result)
        print(json.dumps({k: result[k] for k in ["status", "delivery_status", "deck_score", "issues"]}, ensure_ascii=False))
        return 0 if result["delivery_status"] == "deliverable" else 2
    except (OSError, ValueError, KeyError, TypeError, E.ParseError, zipfile.BadZipFile) as exc:
        write(args.to, {"status": "fail", "delivery_status": "blocked", "error": str(exc)})
        print(str(exc), file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
