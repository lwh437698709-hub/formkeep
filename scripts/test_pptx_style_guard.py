"""Adversarial regressions: tiny additions cannot hide behind an unchanged deck."""
import copy
import json
import tempfile
import unittest
import zipfile
from pathlib import Path

from pptx_style_guard import extract, audit, file_sha, typography_ok
from office_template import inspect, fill

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"


def shape(i=2, color="A1313D", text="Title", name="ot.slot:title", x=100000, size=2400, font="Source Sans"):
    return f'''<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="{name}"/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="{x}" y="100000"/><a:ext cx="2000000" cy="500000"/></a:xfrm>
    <a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:ln w="12700"/></p:spPr>
    <p:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="{size}"><a:latin typeface="{font}"/></a:rPr><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>'''


def package(path, shapes=None, theme="A1313D", extra=""):
    shapes = shapes or shape()
    files = {
        "ppt/presentation.xml": f'<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>',
        "ppt/_rels/presentation.xml.rels": f'<Relationships xmlns="{PR}"><Relationship Id="rId1" Type="{R}/slide" Target="slides/slide1.xml"/></Relationships>',
        "ppt/slides/slide1.xml": f'<p:sld xmlns:p="{P}" xmlns:a="{A}"><p:cSld><p:spTree>{shapes}{extra}</p:spTree></p:cSld></p:sld>',
        "ppt/slides/_rels/slide1.xml.rels": f'<Relationships xmlns="{PR}"><Relationship Id="rL" Type="{R}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>',
        "ppt/slideLayouts/slideLayout1.xml": f'<p:sldLayout xmlns:p="{P}"><p:cSld/></p:sldLayout>',
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels": f'<Relationships xmlns="{PR}"><Relationship Id="rM" Type="{R}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>',
        "ppt/slideMasters/slideMaster1.xml": f'<p:sldMaster xmlns:p="{P}"><p:cSld/></p:sldMaster>',
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": f'<Relationships xmlns="{PR}"><Relationship Id="rT" Type="{R}/theme" Target="../theme/theme1.xml"/></Relationships>',
        "ppt/theme/theme1.xml": f'<a:theme xmlns:a="{A}"><a:themeElements><a:clrScheme name="Test"><a:accent1><a:srgbClr val="{theme}"/></a:accent1></a:clrScheme></a:themeElements></a:theme>',
    }
    with zipfile.ZipFile(path, "w") as z:
        for name, content in files.items():
            z.writestr(name, content)


class StyleGuardTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.source = Path(self.tmp.name)/"source.pptx"
        self.output = Path(self.tmp.name)/"output.pptx"
        package(self.source)

    def codes(self, report):
        return {i["code"] for i in report["issues"]}

    def test_text_replacement_retains_style_but_requires_render(self):
        package(self.output, shape(text="新的业务标题"))
        r = audit(self.source, self.output)
        self.assertEqual(r["deck_score"], 100)
        self.assertTrue(r["structural_pass"])
        self.assertEqual(r["delivery_status"], "needs-native-render-review")

    def test_new_blue_icon_is_blocked_above_90(self):
        originals = "".join(shape(i=i, name=f"ot.slot:s{i}") for i in range(2, 32))
        package(self.source, originals)
        package(self.output, originals + shape(i=99, color="00AAFF", name="New icon"))
        r = audit(self.source, self.output)
        self.assertGreaterEqual(r["deck_score"], 90)
        self.assertIn("UNREGISTERED_ADDITION", self.codes(r))
        self.assertEqual(r["delivery_status"], "blocked")

    def test_donor_cannot_whitelist_wrong_color(self):
        package(self.output, shape()+shape(i=9, color="00AAFF", name="Clone"))
        plan = {"source_sha256": file_sha(self.source), "objects": {"1:9": {
            "donor": "2", "purpose": "metric", "bbox_emu": [100000, 100000, 2000000, 500000]}}}
        self.assertIn("PALETTE_DRIFT", self.codes(audit(self.source, self.output, plan)))

    def test_registered_clone_and_anchor(self):
        package(self.output, shape()+shape(i=9, x=400000, name="Clone"))
        plan = {"source_sha256": file_sha(self.source), "objects": {"1:9": {
            "donor": "2", "purpose": "metric", "bbox_emu": [400000, 100000, 2000000, 500000], "max_delta_emu": [300000, 0, 0, 0]}}}
        self.assertTrue(audit(self.source, self.output, plan)["structural_pass"])
        plan["objects"]["1:9"]["bbox_emu"][0] = 500000
        self.assertIn("ADDITION_ANCHOR_MISMATCH", self.codes(audit(self.source, self.output, plan)))

    def test_direct_font_drift(self):
        package(self.output, shape(font="Unrelated Font"))
        self.assertIn("TYPOGRAPHY_DRIFT", self.codes(audit(self.source, self.output)))

    def test_font_fallback_requires_evidence_and_exact_mapping(self):
        package(self.output, shape(font="Candidate Sans"))
        plan = {"source_sha256": file_sha(self.source), "objects": {"1:2": {"font_fallbacks": {"Source Sans": "Candidate Sans"}}}}
        self.assertIn("TYPOGRAPHY_DRIFT", self.codes(audit(self.source, self.output, plan)))
        plan["objects"]["1:2"]["fallback_evidence"] = "Fixture for installed-font and rendered-width evidence"
        self.assertTrue(audit(self.source, self.output, plan)["structural_pass"])
        self.assertEqual(audit(self.source, self.output, plan)["delivery_status"], "needs-native-render-review")

    def test_stroke_drift_and_stale_plan(self):
        changed = shape().replace('w="12700"', 'w="25400"')
        package(self.output, changed)
        self.assertIn("SHAPE_STYLE_DRIFT", self.codes(audit(self.source, self.output)))
        self.assertIn("PLAN_SOURCE_MISMATCH", self.codes(audit(self.source, self.output, {"source_sha256": "wrong"})))

    def test_theme_change_is_detected_through_master(self):
        package(self.output, theme="00AAFF")
        self.assertIn("MASTER_LAYOUT_THEME_CHANGED", self.codes(audit(self.source, self.output)))

    def test_protected_logo_and_deletion(self):
        package(self.source, shape()+shape(i=3, name="Logo", text="LOGO"))
        package(self.output, shape()+shape(i=3, name="Logo", text="OTHER"))
        self.assertIn("LOCKED_OBJECT_CHANGED", self.codes(audit(self.source, self.output)))
        package(self.output, shape())
        self.assertIn("OBJECT_DELETED", self.codes(audit(self.source, self.output)))

    def test_reorder_and_unsupported_element(self):
        package(self.source, shape()+shape(i=3, name="Logo"))
        package(self.output, shape(i=3, name="Logo")+shape())
        self.assertIn("Z_ORDER_CHANGED", self.codes(audit(self.source, self.output)))
        package(self.output, shape()+shape(i=3, name="Logo"), extra='<p:unknown/>')
        self.assertIn("UNMODELED_TREE_CHANGED", self.codes(audit(self.source, self.output)))

    def test_bounds_and_uniform_font_scale(self):
        package(self.output, shape(x=150000, size=2280))
        plan = {"source_sha256": file_sha(self.source), "objects": {"1:2": {"max_delta_emu": [50000, 0, 0, 0], "font_scale": 0.95, "min_font_pt": 18}}}
        self.assertTrue(audit(self.source, self.output, plan)["structural_pass"])
        plan["objects"]["1:2"]["min_font_pt"] = 23
        self.assertIn("TYPOGRAPHY_DRIFT", self.codes(audit(self.source, self.output, plan)))

    def test_stale_or_missing_visual_evidence(self):
        package(self.output)
        r = audit(self.source, self.output, review={"source_sha256": "wrong"})
        self.assertIn("RENDER_EVIDENCE_HASH_MISMATCH", self.codes(r))
        review = {"source_sha256": file_sha(self.source), "output_sha256": file_sha(self.output), "renderer": "test", "native_office_render": True, "slides": []}
        self.assertIn("INCOMPLETE_VISUAL_REVIEW", self.codes(audit(self.source, self.output, review=review)))

    def test_extract_anchors_and_palette(self):
        book = extract(self.source)
        o = book["slides"][0]["objects"][0]
        self.assertIn("A1313D", "".join(book["palette_tokens"]))
        self.assertGreater(o["anchors"]["right"], o["anchors"]["left"])
        self.assertEqual(book["font_availability"], "not-checked")

    def test_fill_integrates_pre_and_post_style_gate(self):
        spec = inspect(self.source)
        slot = spec["slots"][0]
        slot["budget"] = {"min_chars": 1, "recommended_max_chars": 20, "hard_max_chars": 24, "max_width_units_per_line": 20, "max_lines": 2}
        slot["calibration"] = {"status": "verified"}
        bookfile = Path(self.tmp.name)/"visual.json"
        bookfile.write_text(json.dumps(extract(self.source)))
        r = fill(self.source, spec, {"title": "业绩分析"}, self.output, False, bookfile)
        self.assertEqual(r["status"], "pass")
        self.assertEqual(audit(self.source, self.output)["deck_score"], 100)
        badbook = extract(self.source)
        badbook["source_sha256"] = "wrong"
        bookfile.write_text(json.dumps(badbook))
        r = fill(self.source, spec, {"title": "业绩分析"}, self.output, True, bookfile)
        self.assertEqual(r["status"], "fail")


if __name__ == "__main__":
    unittest.main()
