"""Regression tests for dense agenda, plan and timeline layouts."""
import sys
import unittest
from pathlib import Path

from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import generate_ppt as ppt


class LayoutBoundsTests(unittest.TestCase):
    def setUp(self):
        self.prs = Presentation()
        self.theme = ppt.Theme({"slide": {"width_in": 13.333, "height_in": 7.5}})
        self.prs.slide_width = ppt.Inches(self.theme.W)
        self.prs.slide_height = ppt.Inches(self.theme.H)
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])

    def assert_inside_slide(self):
        for shape in self.slide.shapes:
            with self.subTest(shape=shape.name):
                self.assertGreaterEqual(shape.left, 0)
                self.assertGreaterEqual(shape.top, 0)
                self.assertLessEqual(shape.left + shape.width, self.prs.slide_width + 2)
                self.assertLessEqual(shape.top + shape.height, self.prs.slide_height + 2)

    def test_agenda_twelve_items_fit_and_remain_editable(self):
        ppt.draw_agenda(self.slide, {"items": [f"Item {i}" for i in range(12)]}, self.theme)
        self.assert_inside_slide()
        text = [s.text for s in self.slide.shapes if s.has_text_frame]
        for i in range(12):
            self.assertIn(f"Item {i}", text)

    def test_next_steps_twelve_items_fit(self):
        ppt.draw_next_steps(self.slide, {"items": [{"task": f"Task {i}"} for i in range(12)]}, self.theme)
        self.assert_inside_slide()
        text = [s.text for s in self.slide.shapes if s.has_text_frame]
        for i in range(12):
            self.assertIn(f"Task {i}", text)

    def test_timeline_twelve_titles_do_not_overlap(self):
        steps = [{"time": f"Time {i}", "title": f"Step {i}", "detail": "Detail"} for i in range(12)]
        ppt.draw_timeline(self.slide, {"steps": steps}, self.theme)
        self.assert_inside_slide()
        titles = [s for s in self.slide.shapes if s.has_text_frame and s.text.startswith("Step ")]
        self.assertEqual(len(titles), 12)
        for a in titles:
            for b in titles:
                if a is b:
                    continue
                overlap = (a.left < b.left + b.width and b.left < a.left + a.width and
                           a.top < b.top + b.height and b.top < a.top + a.height)
                self.assertFalse(overlap, f"{a.text} overlaps {b.text}")


if __name__ == "__main__":
    unittest.main()
