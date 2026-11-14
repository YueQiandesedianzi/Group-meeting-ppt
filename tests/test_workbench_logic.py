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




SAMPLE = """meta:
  title: 测试
sections:
  - type: cover
    title: 封面

  - type: points
    heading: 方法
    points:
      - A

  - type: conclusion
    heading: 小结
    points:
      - B

notes: keep me
"""


def parse(text):
    import yaml
    return yaml.safe_load(text)


class SectionOutlineTests(unittest.TestCase):
    def test_rows_have_index_type_label_and_title(self):
        rows = wl.section_outline(parse(SAMPLE))
        self.assertEqual([r["index"] for r in rows], [1, 2, 3])
        self.assertEqual([r["type"] for r in rows], ["cover", "points", "conclusion"])
        self.assertEqual([r["label"] for r in rows], ["封面", "要点", "结论"])
        self.assertEqual([r["title"] for r in rows], ["封面", "方法", "小结"])

    def test_unknown_and_broken_sections(self):
        rows = wl.section_outline({"sections": [{"type": "weird"}, "nope"]})
        self.assertEqual(rows[0]["label"], "weird")
        self.assertEqual(rows[1]["label"], "未知")

    def test_bad_input(self):
        self.assertEqual(wl.section_outline(None), [])
        self.assertEqual(wl.section_outline({"sections": 3}), [])

    def test_type_label_fallback(self):
        self.assertEqual(wl.type_label("table"), "表格")
        self.assertEqual(wl.type_label("unknown"), "unknown")
        self.assertEqual(wl.type_label(None), "未知")


class SectionSpansTests(unittest.TestCase):
    def test_three_spans_match_parsed_sections(self):
        text = SAMPLE.rstrip("\n") + "\n"
        spans = wl.section_spans(text)
        self.assertEqual(len(spans), 3)
        lines = text.split("\n")
        for (start, end), typ in zip(spans, ["cover", "points", "conclusion"]):
            block = "\n".join(lines[start:end])
            self.assertIn(f"type: {typ}", block)

    def test_no_sections_key(self):
        self.assertEqual(wl.section_spans("meta:\n  title: x\n"), [])


class MoveRemoveTests(unittest.TestCase):
    def test_move_down_swaps_neighbours_and_keeps_tail(self):
        moved = wl.move_section_text(SAMPLE, 0, 1)
        types = [s["type"] for s in parse(moved)["sections"]]
        self.assertEqual(types, ["points", "cover", "conclusion"])
        self.assertIn("notes: keep me", moved)

    def test_move_up_and_boundaries(self):
        moved = wl.move_section_text(SAMPLE, 2, -1)
        self.assertEqual([s["type"] for s in parse(moved)["sections"]],
                         ["cover", "conclusion", "points"])
        self.assertEqual(wl.move_section_text(SAMPLE, 0, -1), SAMPLE)
        self.assertEqual(wl.move_section_text(SAMPLE, 2, 1), SAMPLE)
        self.assertEqual(wl.move_section_text(SAMPLE, 9, 1), SAMPLE)

    def test_remove_section(self):
        fewer = wl.remove_section_text(SAMPLE, 1)
        self.assertEqual([s["type"] for s in parse(fewer)["sections"]], ["cover", "conclusion"])
        self.assertEqual(wl.remove_section_text(SAMPLE, 5), SAMPLE)

    def test_unparseable_text_is_untouched(self):
        self.assertEqual(wl.move_section_text("meta:\n  title: x\n", 0, 1), "meta:\n  title: x\n")
        self.assertEqual(wl.remove_section_text("meta:\n  title: x\n", 0), "meta:\n  title: x\n")


class SnippetTests(unittest.TestCase):
    def test_append_adds_one_section_of_each_type(self):
        text = SAMPLE
        for i, typ in enumerate(wl.SECTION_TYPES, start=4):
            text = wl.append_snippet(text, typ)
            sections = parse(text)["sections"]
            self.assertEqual(len(sections), i)
            self.assertEqual(sections[-1]["type"], typ)
        self.assertIn("notes: keep me", text)

    def test_append_keeps_yaml_valid_for_every_snippet(self):
        import base64
        import os
        import sys
        import tempfile
        sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
        import generate_ppt as g
        png = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8AARAAB/2wD5QAAAABJRU5ErkJggg==")
        with tempfile.TemporaryDirectory() as tmp:
            with open(os.path.join(tmp, "图片文件名.png"), "wb") as fh:
                fh.write(png)
            for typ in wl.SECTION_TYPES:
                with self.subTest(typ=typ):
                    text = wl.append_snippet(SAMPLE, typ)
                    g.validate(parse(text), tmp)

    def test_unknown_type_raises(self):
        with self.assertRaises(ValueError):
            wl.append_snippet(SAMPLE, "poster")

    def test_missing_sections_raises(self):
        with self.assertRaises(ValueError):
            wl.append_snippet("meta:\n  title: x\n", "cover")

    def test_append_into_empty_sections(self):
        text = "meta:\n  title: x\nsections:\n"
        out = wl.append_snippet(text, "cover")
        self.assertEqual(len(parse(out)["sections"]), 1)


if __name__ == "__main__":
    unittest.main()