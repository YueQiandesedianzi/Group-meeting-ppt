#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""M0 回归测试：把技术验证固化成可复现断言。

默认只跑快速用例（不调用 LibreOffice）。需要客户端往返的用例设置环境变量
GM_PPT_SLOW=1 后运行：

    python -m unittest discover -s v2/tests -v
    $env:GM_PPT_SLOW=1; python -m unittest discover -s v2/tests -v
"""
import json
import os
import shutil
import sys
import tempfile
import unittest

V2 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(V2, "tools"))
sys.path.insert(0, os.path.join(V2, "engine"))

import object_ops as ops  # noqa: E402
import pptx_inspect as pi  # noqa: E402
import retention_check as retention  # noqa: E402

SAMPLES = os.path.join(V2, "samples")
REFERENCE = os.path.join(SAMPLES, "reference-template.pptx")
HAND_EDITED = os.path.join(SAMPLES, "hand-edited.pptx")
TAGS_SAMPLE = os.path.join(SAMPLES, "tags-sample.pptx")
OBJECT_OPS = os.path.join(SAMPLES, "object-ops.pptx")
SLOW = os.environ.get("GM_PPT_SLOW") == "1"

from pptx import Presentation  # noqa: E402


def markers(report):
    out = []
    for slide in report["slides"]:
        for shape in slide["shapes"]:
            if shape.get("marker"):
                out.append(shape["marker"])
    return out


def find_marker(report, prefix):
    for slide in report["slides"]:
        for shape in slide["shapes"]:
            if (shape.get("marker") or "").startswith(prefix):
                return shape["marker"], slide["slide_id"]
    raise AssertionError("找不到前缀 %s" % prefix)


class InspectTests(unittest.TestCase):
    def setUp(self):
        self.report = pi.inspect_pptx(REFERENCE)

    def test_template_shape_and_size(self):
        self.assertEqual(self.report["slide_count"], 5)
        self.assertEqual(self.report["slide_size_in"], [13.333, 8.333])

    def test_master_carries_logo_and_footer(self):
        master_markers = [s.get("marker") for m in self.report["masters"] for s in m["shapes"]]
        self.assertTrue(any((m or "").startswith("gmppt:master-logo:") for m in master_markers))
        self.assertTrue(any((m or "").startswith("gmppt:master-footer:") for m in master_markers))

    def test_every_page_has_stable_identity_marker(self):
        for slide in self.report["slides"]:
            self.assertTrue([s for s in slide["shapes"] if s.get("marker")],
                            "第 %d 页没有身份标记" % (slide["index"] + 1))


class IdentityConflictTests(unittest.TestCase):
    """复制页面会导致身份重复；引擎必须拒绝写入而不是猜。"""

    def setUp(self):
        self.report = pi.inspect_pptx(HAND_EDITED)
        self.marker, _ = find_marker(self.report, "gmppt:page-title:cover-title=")

    def test_duplicate_identity_is_detected(self):
        hits = [m for m in markers(self.report) if m == self.marker]
        self.assertEqual(len(hits), 2, "复制页应产生恰好两个相同身份")

    def test_find_objects_reports_ambiguity(self):
        prs = Presentation(HAND_EDITED)
        found = ops.find_objects(prs, self.marker)
        self.assertEqual(len(found), 2)

    def test_write_is_refused_on_duplicate_identity(self):
        prs = Presentation(HAND_EDITED)
        with self.assertRaises(ops.CapabilityError):
            ops.set_text(prs, self.marker, "不应写入")

    def test_unique_identity_still_writable(self):
        prs = Presentation(HAND_EDITED)
        marker, _ = find_marker(self.report, "gmppt:page-title:summary-title=")
        result = ops.set_text(prs, marker, "小结（已改）")
        self.assertEqual(result["op"], "set_text")


class ObjectOpsTests(unittest.TestCase):
    def test_targeted_edit_keeps_untouched_objects_and_parts(self):
        tmp = tempfile.mkdtemp(prefix="m0-ops-")
        self.addCleanup(shutil.rmtree, tmp, True)
        target = os.path.join(tmp, "edited.pptx")
        shutil.copyfile(REFERENCE, target)

        before = pi.inspect_pptx(target)
        cover_marker, _ = find_marker(before, "gmppt:page-title:cover-title=")
        chart_marker, _ = find_marker(before, "gmppt:picture:result-chart=")

        prs = Presentation(target)
        ops.apply_plan(prs, [
            {"op": "set_text", "identity": cover_marker, "text": "只改这一段"},
            {"op": "set_geometry", "identity": chart_marker, "width": 6.4, "keep_aspect": True},
        ])
        prs.save(target)

        after = pi.inspect_pptx(target)
        touched = {cover_marker, chart_marker}

        def fingerprint(report):
            out = {}
            for slide in report["slides"]:
                for shape in slide["shapes"]:
                    out[(slide["slide_id"], shape.get("name"))] = {
                        "type": shape.get("type"), "text": shape.get("text"),
                        "frame": shape.get("frame"), "marker": shape.get("marker")}
            return out

        fp_before, fp_after = fingerprint(before), fingerprint(after)
        changed = [k for k, v in fp_before.items()
                   if v.get("marker") not in touched and fp_after.get(k) != v]
        self.assertEqual(changed, [], "未指向对象被改动：%s" % changed)
        self.assertEqual(set(fp_after) - set(fp_before), set(), "出现了意外新增对象")

        result = retention.compare(REFERENCE, target)
        self.assertGreaterEqual(result["totals"].get("identical", 0), 40,
                                "定向修改应保留绝大多数部件：%s" % result["totals"])

    def test_reorder_rejects_mismatched_parameters(self):
        prs = Presentation(REFERENCE)
        with self.assertRaises(ops.CapabilityError):
            ops.reorder_slides(prs, [1, 2, 3])

    def test_picture_replacement_uses_new_media_part(self):
        tmp = tempfile.mkdtemp(prefix="m0-pic-")
        self.addCleanup(shutil.rmtree, tmp, True)
        target = os.path.join(tmp, "pic.pptx")
        shutil.copyfile(REFERENCE, target)
        report = pi.inspect_pptx(target)
        chart_marker, _ = find_marker(report, "gmppt:picture:result-chart=")
        image = os.path.join(SAMPLES, "assets", "chart-accuracy-v2.png")

        import zipfile
        media_before = [n for n in zipfile.ZipFile(target).namelist() if n.startswith("ppt/media/")]

        prs = Presentation(target)
        result = ops.replace_picture(prs, chart_marker, image)
        prs.save(target)
        self.assertTrue(result["shared_media_untouched"], "共享媒体不应被覆盖")
        self.assertTrue(result["new_media"].startswith("/ppt/media/"))
        media_after = [n for n in zipfile.ZipFile(target).namelist() if n.startswith("ppt/media/")]
        self.assertEqual(len(media_after), len(media_before) + 1, "应新增一个媒体部件")


class TagsTests(unittest.TestCase):
    def test_tags_survive_our_own_save(self):
        self.assertTrue(os.path.isfile(TAGS_SAMPLE), "先运行 v2/tools/tags_probe.py")
        tmp = tempfile.mkdtemp(prefix="m0-tags-")
        self.addCleanup(shutil.rmtree, tmp, True)
        resaved = os.path.join(tmp, "resaved.pptx")
        prs = Presentation(TAGS_SAMPLE)
        prs.save(resaved)

        before = pi.inspect_pptx(TAGS_SAMPLE)
        after = pi.inspect_pptx(resaved)

        def tag_map(report):
            return {s["marker"]: s["tags"] for sl in report["slides"] for s in sl["shapes"]
                    if s.get("marker") and s.get("tags")}
        bt, at = tag_map(before), tag_map(after)
        self.assertEqual(len(bt), 9)
        self.assertEqual(bt, at, "我方保存后 Tags 应完全一致")


@unittest.skipUnless(SLOW, "需要 LibreOffice 往返：设置 GM_PPT_SLOW=1")
class RoundTripTests(unittest.TestCase):
    def test_client_save_rewrites_most_parts(self):
        converted = os.path.join(SAMPLES, "reference-template.roundtrip.pptx")
        if not os.path.isfile(converted):
            import roundtrip_check as rc
            converted = os.path.join(SAMPLES, "reference-template.roundtrip.pptx")
            rc.convert(rc.DEFAULT_NODE, rc.DEFAULT_CLI, REFERENCE, converted)
        result = retention.compare(REFERENCE, converted)
        self.assertLessEqual(result["totals"].get("identical", 0), 5)
        self.assertGreaterEqual(result["totals"].get("rewritten", 0), 40)

    def test_client_save_drops_powerpoint_tags(self):
        converted = os.path.join(SAMPLES, "tags-sample.roundtrip.pptx")
        if not os.path.isfile(converted):
            import roundtrip_check as rc
            rc.convert(rc.DEFAULT_NODE, rc.DEFAULT_CLI, TAGS_SAMPLE, converted)
        report = pi.inspect_pptx(converted)
        tagged = [s for sl in report["slides"] for s in sl["shapes"] if s.get("tags")]
        self.assertEqual(tagged, [], "LibreOffice 保存后不应残留 Tags（代理行为）")


if __name__ == "__main__":
    unittest.main()
