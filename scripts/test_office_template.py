#!/usr/bin/env python3
from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

SCRIPT = Path(__file__).with_name("office_template.py")


def write_zip(path: Path, files: dict[str, str]) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        for name, value in files.items():
            package.writestr(name, value)


def run(*args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run([sys.executable, str(SCRIPT), *args], text=True, capture_output=True, check=False)


class OfficeTemplateTest(unittest.TestCase):
    def test_docx_inspect_validate_and_fill(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            template, output = root / "template.docx", root / "output.docx"
            spec_path, bindings_path = root / "spec.json", root / "bindings.json"
            write_zip(template, {
                "[Content_Types].xml": "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>",
                "word/document.xml": """<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body><w:sdt><w:sdtPr><w:tag w:val='ot.slot:resume.summary'/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>placeholder</w:t></w:r></w:p></w:sdtContent></w:sdt></w:body></w:document>""",
            })
            result = run("inspect", str(template), "--to", str(spec_path))
            self.assertEqual(result.returncode, 0, result.stdout)
            spec = json.loads(spec_path.read_text())
            self.assertEqual(spec["slots"][0]["id"], "resume.summary")
            spec["slots"][0]["budget"] = {"min_chars": 2, "recommended_max_chars": 20, "hard_max_chars": 24, "max_width_units_per_line": 12, "max_lines": 2}
            spec["slots"][0]["must_include"] = ["18%"]
            spec["slots"][0]["calibration"] = {"status": "verified", "renderer": "test", "font": "test"}
            spec_path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
            bindings_path.write_text(json.dumps({"resume.summary": "点击率提升18%"}, ensure_ascii=False), encoding="utf-8")
            result = run("fill", str(template), str(spec_path), str(bindings_path), str(output))
            self.assertEqual(result.returncode, 0, result.stdout)
            with zipfile.ZipFile(output) as package:
                self.assertIn("点击率提升18%", package.read("word/document.xml").decode())

    def test_overflow_is_blocked(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            spec_path, bindings_path = root / "spec.json", root / "bindings.json"
            spec = {"format": "office-template-spec@1", "template_kind": "pptx", "template_sha256": "x", "slots": [{"id": "title", "required": True, "budget": {"min_chars": 1, "recommended_max_chars": 3, "hard_max_chars": 4, "max_width_units_per_line": 4, "max_lines": 1}, "must_include": [], "calibration": {"status": "verified"}}]}
            spec_path.write_text(json.dumps(spec), encoding="utf-8")
            bindings_path.write_text(json.dumps({"title": "too long"}), encoding="utf-8")
            result = run("validate", str(spec_path), str(bindings_path))
            self.assertEqual(result.returncode, 2)
            self.assertIn("HARD_MAXIMUM", result.stdout)

    def test_pptx_inspect_and_fill(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            template, output = root / "template.pptx", root / "output.pptx"
            spec_path, bindings_path = root / "spec.json", root / "bindings.json"
            write_zip(template, {
                "[Content_Types].xml": "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>",
                "ppt/slides/slide1.xml": """<p:sld xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main' xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id='2' name='ot.slot:title'/></p:nvSpPr><p:txBody><a:p><a:r><a:t>placeholder</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>""",
            })
            self.assertEqual(run("inspect", str(template), "--to", str(spec_path)).returncode, 0)
            spec = json.loads(spec_path.read_text())
            spec["slots"][0]["budget"] = {"min_chars": 1, "recommended_max_chars": 20, "hard_max_chars": 24, "max_width_units_per_line": 20, "max_lines": 2}
            spec["slots"][0]["calibration"] = {"status": "verified", "renderer": "test", "font": "test"}
            spec_path.write_text(json.dumps(spec), encoding="utf-8")
            bindings_path.write_text(json.dumps({"title": "Quarterly Review"}), encoding="utf-8")
            result = run("fill", str(template), str(spec_path), str(bindings_path), str(output))
            self.assertEqual(result.returncode, 2, result.stdout)
            self.assertIn("VISUAL_HANDBOOK_REQUIRED", result.stdout)
            self.assertFalse(output.exists())

    def test_xlsx_inspect_and_fill(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            template, output = root / "template.xlsx", root / "output.xlsx"
            spec_path, bindings_path = root / "spec.json", root / "bindings.json"
            write_zip(template, {
                "[Content_Types].xml": "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>",
                "xl/workbook.xml": """<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><sheets><sheet name='Sheet1' sheetId='1' r:id='rId1'/></sheets><definedNames><definedName name='ot_slot_resume__name'>Sheet1!$B$2</definedName></definedNames></workbook>""",
                "xl/_rels/workbook.xml.rels": """<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Target='worksheets/sheet1.xml'/></Relationships>""",
                "xl/worksheets/sheet1.xml": """<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><sheetData><row r='2'><c r='B2' t='inlineStr'><is><t>placeholder</t></is></c></row></sheetData></worksheet>""",
            })
            self.assertEqual(run("inspect", str(template), "--to", str(spec_path)).returncode, 0)
            spec = json.loads(spec_path.read_text())
            self.assertEqual(spec["slots"][0]["id"], "resume.name")
            spec["slots"][0]["budget"] = {"min_chars": 1, "recommended_max_chars": 20, "hard_max_chars": 24, "max_width_units_per_line": 20, "max_lines": 2}
            spec["slots"][0]["calibration"] = {"status": "verified", "renderer": "test", "font": "test"}
            spec_path.write_text(json.dumps(spec), encoding="utf-8")
            bindings_path.write_text(json.dumps({"resume.name": "Ada Lovelace"}), encoding="utf-8")
            result = run("fill", str(template), str(spec_path), str(bindings_path), str(output))
            self.assertEqual(result.returncode, 0, result.stdout)
            with zipfile.ZipFile(output) as package:
                self.assertIn("Ada Lovelace", package.read("xl/worksheets/sheet1.xml").decode())


if __name__ == "__main__":
    unittest.main()
