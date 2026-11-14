#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""对象级修改原型验证：定向修改 + 未触碰对象保留 + 部件级保留。

产物：samples/object-ops.pptx 与 reports/object-ops.json
"""
import json
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "engine"))

import pptx_inspect as pi  # noqa: E402
import retention_check as retention  # noqa: E402
import object_ops as ops  # noqa: E402

from PIL import Image, ImageDraw  # noqa: E402
from pptx import Presentation  # noqa: E402

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SAMPLES = os.path.join(HERE, "samples")
REPORTS = os.path.join(HERE, "reports")


def find_marker(report, prefix):
    for slide in report["slides"]:
        for shape in slide["shapes"]:
            marker = shape.get("marker") or ""
            if marker.startswith(prefix):
                return marker, slide["slide_id"]
    raise SystemExit("找不到标记前缀 %s" % prefix)


def fingerprint(report):
    out = {}
    for slide in report["slides"]:
        for shape in slide["shapes"]:
            key = (slide["slide_id"], shape.get("name"))
            out[key] = {
                "type": shape.get("type"),
                "text": shape.get("text"),
                "frame": shape.get("frame"),
                "marker": shape.get("marker"),
            }
    return out


def make_new_chart(path):
    img = Image.new("RGB", (800, 450), "white")
    draw = ImageDraw.Draw(img)
    draw.rectangle([0, 0, 799, 449], outline="#C9C9C9")
    for index, value in enumerate([0.91, 0.93, 0.95]):
        x = 120 + index * 220
        height = int(value * 380)
        draw.rectangle([x, 415 - height, x + 120, 415], fill="#17406D")
        draw.text((x + 8, 420), "v2-%d %.2f" % (index + 1, value), fill="#333333")
    draw.text((20, 12), "v2 accuracy (replaced)", fill="#17406D")
    img.save(path)
    return path


def main():
    source = os.path.join(SAMPLES, "reference-template.pptx")
    target = os.path.join(SAMPLES, "object-ops.pptx")
    shutil.copyfile(source, target)

    before = pi.inspect_pptx(target)
    cover_marker, cover_slide = find_marker(before, "gmppt:page-title:cover-title=")
    chart_marker, chart_slide = find_marker(before, "gmppt:picture:result-chart=")
    summary_marker, summary_slide = find_marker(before, "gmppt:page-title:summary-title=")
    new_chart = make_new_chart(os.path.join(SAMPLES, "assets", "chart-accuracy-v2.png"))

    prs = Presentation(target)
    order = [s["slide_id"] for s in before["slides"]]
    order.remove(summary_slide)
    order.insert(1, summary_slide)

    plan = [
        {"op": "set_text", "identity": cover_marker, "text": "对象级修改：只改这一段"},
        {"op": "replace_picture", "identity": chart_marker, "image": new_chart},
        {"op": "set_geometry", "identity": chart_marker, "left": 1.2, "top": 2.2,
         "width": 6.4, "keep_aspect": True},
        {"op": "reorder_slides", "order": order},
    ]
    results = ops.apply_plan(prs, plan)
    prs.save(target)

    after = pi.inspect_pptx(target)
    fp_before, fp_after = fingerprint(before), fingerprint(after)
    touched = {cover_marker, chart_marker}
    changed, missing, added = [], [], []
    for key, value in fp_before.items():
        if value.get("marker") in touched:
            continue
        if key not in fp_after:
            missing.append(key)
        elif fp_after[key] != value:
            changed.append({"key": list(key), "before": value, "after": fp_after[key]})
    for key in fp_after:
        if key not in fp_before:
            added.append(list(key))

    retention_report = retention.compare(source, target)
    report = {
        "source": source,
        "target": target,
        "plan": plan,
        "results": results,
        "touched_identities": sorted(touched),
        "untouched_changed": changed,
        "untouched_missing": missing,
        "unexpected_added": added,
        "retention_totals": retention_report["totals"],
        "retention_by_category": retention_report["by_category"],
        "verdict": "pass" if not (changed or missing or added) else "review",
    }
    os.makedirs(REPORTS, exist_ok=True)
    with open(os.path.join(REPORTS, "object-ops.json"), "w", encoding="utf-8") as fh:
        json.dump(report, fh, ensure_ascii=False, indent=2)
    print(json.dumps({k: report[k] for k in ("verdict", "untouched_changed", "untouched_missing",
                                             "unexpected_added", "retention_totals")},
                     ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
