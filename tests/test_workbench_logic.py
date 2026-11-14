import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
import workbench_logic as wl


def brief(*sections):
    return {"meta": {"title": "t"}, "sections": list(sections)}


class DisplayLabelsTests(unittest.TestCase):
    def test_basename_only_when_unique(self):
        paths = [r"C:\a\x.yaml", r"C:\b\y.yaml"]
        self.assertEqual(wl.display_labels(paths), ["x.yaml", "y.yaml"])

    def test_duplicate_names_get_parent_folder(self):
        paths = [r"C:\a\same.yaml", r"D:\b\same.yaml"]
        self.assertEqual(wl.display_labels(paths), ["same.yaml（a）", "same.yaml（b）"])

    def test_empty_input(self):
        self.assertEqual(wl.display_labels([]), [])


class EstimatePagesTests(unittest.TestCase):
    def test_counts_sections(self):
        self.assertEqual(wl.estimate_pages(brief({"type": "cover"}, {"type": "points"})), 2)

    def test_bad_input_is_zero(self):
        self.assertEqual(wl.estimate_pages(None), 0)
        self.assertEqual(wl.estimate_pages({"sections": "nope"}), 0)


class WarningTests(unittest.TestCase):
    def test_clean_brief_has_no_warnings(self):
        b = brief({"type": "points", "heading": "方法", "points": ["a", "b"]})
        self.assertEqual(wl.compute_warnings(b), [])

    def test_too_many_points(self):
        b = brief({"type": "points", "heading": "方法", "points": [str(i) for i in range(8)]})
        self.assertEqual(len(wl.compute_warnings(b)), 1)
        self.assertIn("points 有 8 条", wl.compute_warnings(b)[0])

    def test_long_point_and_title(self):
        b = brief({"type": "points", "heading": "这是一个明显超过二十个字的很长的标题文本内容", "points": ["x" * 41]})
        warnings = wl.compute_warnings(b)
        self.assertEqual(len(warnings), 2)
        self.assertTrue(any("标题" in w for w in warnings))
        self.assertTrue(any("超过 40 字" in w for w in warnings))

    def test_too_many_takeaways(self):
        b = brief({"type": "figure", "heading": "结果", "image": "a.png",
                   "takeaways": ["1", "2", "3", "4"]})
        self.assertIn("takeaways", wl.compute_warnings(b)[0])

    def test_many_sections(self):
        b = brief(*[{"type": "section_header", "heading": "x"} for _ in range(25)])
        self.assertTrue(any("篇幅较长" in w for w in wl.compute_warnings(b)))

    def test_missing_sections(self):
        self.assertEqual(wl.compute_warnings({"meta": {}}), [])


class PreflightTextTests(unittest.TestCase):
    def test_with_warnings_lists_items(self):
        text = wl.format_preflight(12, ["第 3 段 points 有 8 条"])
        self.assertIn("预计 12 页", text)
        self.assertIn("1 条提醒", text)
        self.assertIn("• 第 3 段 points 有 8 条", text)

    def test_without_warnings(self):
        text = wl.format_preflight(5, [])
        self.assertIn("没有发现", text)


if __name__ == "__main__":
    unittest.main()
