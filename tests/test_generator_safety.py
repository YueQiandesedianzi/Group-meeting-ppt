import contextlib
import io
import os
from pathlib import Path
import sys
import tempfile
import unittest

import yaml
from pptx import Presentation

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import generate_ppt as g
import workbench as wb


class InvalidBriefTests(unittest.TestCase):
    def test_invalid_structures_raise_value_error(self):
        invalid = [
            None, [], "text", {"meta": [], "sections": []},
            {"meta": {"title": 4}, "sections": []},
            {"meta": {"title": "Demo"}, "sections": [None]},
            {"meta": {"title": "Demo"}, "sections": [{"type": []}]},
            {"meta": {"title": "Demo"}, "sections": [{"type": "points", "heading": "H", "points": [None]}]},
            {"meta": {"title": "Demo"}, "sections": [{"type": "points", "heading": "H", "points": "text"}]},
            {"meta": {"title": "Demo"}, "sections": [{"type": "metrics", "heading": "H", "metrics": [1]}]},
        ]
        for value in invalid:
            with self.subTest(value=value), self.assertRaises(ValueError):
                g.validate(value, str(ROOT))


class GeneratorSafetyTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.dir = Path(self.temp.name)
        self.brief = self.dir / "brief.yaml"
        self.content = {"meta": {"title": "Demo"}, "sections": [{"type": "cover", "title": "Demo"}]}
        self.write_brief()

    def write_brief(self):
        self.brief.write_text(yaml.safe_dump(self.content), encoding="utf-8")

    def run_main(self, *args):
        with contextlib.redirect_stderr(io.StringIO()), contextlib.redirect_stdout(io.StringIO()):
            return g.main([str(self.brief), *map(str, args)])

    def test_output_cannot_overwrite_brief(self):
        before = self.brief.read_bytes()
        self.assertEqual(self.run_main("-o", self.brief), 2)
        with self.assertRaises(ValueError):
            wb.run_generate(str(self.brief), str(self.brief))
        self.assertEqual(self.brief.read_bytes(), before)

    def test_output_hardlink_cannot_overwrite_brief(self):
        alias = self.dir / "alias.pptx"
        try:
            os.link(self.brief, alias)
        except OSError as exc:
            self.skipTest(f"hard links unavailable: {exc}")
        before = self.brief.read_bytes()
        self.assertEqual(self.run_main("-o", alias), 2)
        with self.assertRaises(ValueError):
            wb.run_generate(str(self.brief), str(alias))
        self.assertEqual(self.brief.read_bytes(), before)

    def test_output_cannot_overwrite_template(self):
        template_dir = self.dir / "templates"
        template_dir.mkdir()
        template = template_dir / "default.yaml"
        template.write_text("name: Test\n", encoding="utf-8")
        before = template.read_bytes()
        self.assertEqual(self.run_main("-o", template, "--template-dir", template_dir), 2)
        with self.assertRaises(ValueError):
            wb.run_generate(str(self.brief), str(template), tdir=str(template_dir))
        self.assertEqual(template.read_bytes(), before)

    def test_output_cannot_overwrite_figure(self):
        image = self.dir / "figure.png"
        image.write_bytes((ROOT / "assets" / "demo_figure.png").read_bytes())
        self.content["sections"].append({"type": "figure", "heading": "H", "image": str(image)})
        self.write_brief()
        before = image.read_bytes()
        self.assertEqual(self.run_main("-o", image), 2)
        with self.assertRaises(ValueError):
            wb.run_generate(str(self.brief), str(image))
        self.assertEqual(image.read_bytes(), before)

    def test_single_slide_mismatched_count_does_not_change_existing_ppt(self):
        output = self.dir / "existing.pptx"
        prs = Presentation()
        prs.slides.add_slide(prs.slide_layouts[6])
        prs.slides.add_slide(prs.slide_layouts[6])
        prs.save(output)
        before = output.read_bytes()
        self.assertEqual(self.run_main("-o", output, "--slide", "1"), 2)
        self.assertEqual(output.read_bytes(), before)

    def test_invalid_yaml_document_fails_cleanly(self):
        self.brief.write_text("- wrong\n- document\n", encoding="utf-8")
        output = self.dir / "output.pptx"
        self.assertEqual(self.run_main("-o", output), 2)
        self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()
