#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""只读检查 PPTX 结构：页面、对象、身份候选与保留信息。

M0 用途：为 v2 的项目/页面/对象身份提供事实来源，并提供两文件结构差异。
不修改输入文件，不依赖 PowerPoint。
"""
import argparse
import json
import os
import sys

from pptx import Presentation
from pptx.util import Emu

DESCR_PREFIX = "gmppt:"
TAGS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags"
NS_TAGS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def shape_text(shape):
    if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
        text = shape.text_frame.text
        return text.strip()[:80]
    if shape.has_table if hasattr(shape, "has_table") else False:
        rows = [" | ".join(c.text for c in row.cells) for row in shape.table.rows]
        return " / ".join(rows)[:80]
    return ""


def shape_marker(shape):
    """读取写在备用文本(alt text)里的身份标记。"""
    try:
        descr = shape._element._nvXxPr.cNvPr.get("descr") or ""
    except Exception:
        return None
    return descr if descr.startswith(DESCR_PREFIX) else None


def shape_tags(shape):
    """读取挂在对象上的 PowerPoint Tags（p:nvPr/p:custDataLst/p:tags → tags 部件）。"""
    try:
        nv_pr = shape._element._nvXxPr.nvPr
        tags_ref = nv_pr.find(".//p:custDataLst/p:tags", {"p": NS_TAGS})
        if tags_ref is None:
            return {}
        rid = tags_ref.get("{%s}id" % "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        part = shape.part.related_part(rid)
    except Exception:
        return {}
    try:
        from lxml import etree
        root = etree.fromstring(part.blob)
        return {tag.get("name"): tag.get("val")
                for tag in root.findall("{%s}tag" % NS_TAGS)}
    except Exception:
        return {}


def describe_shape(shape):
    item = {
        "shape_id": getattr(shape, "shape_id", None),
        "name": getattr(shape, "name", None),
        "type": str(getattr(shape, "shape_type", None)),
        "marker": shape_marker(shape),
        "tags": shape_tags(shape),
    }
    text = shape_text(shape)
    if text:
        item["text"] = text
    try:
        item["frame"] = [int(shape.left), int(shape.top), int(shape.width), int(shape.height)]
    except Exception:
        pass
    return item


def shape_tree(shapes):
    out = []
    for shape in shapes:
        if shape.shape_type == 6 and hasattr(shape, "shapes"):  # GROUP
            out.append({"group": shape.name, "children": shape_tree(shape.shapes)})
        else:
            out.append(describe_shape(shape))
    return out


def native_slide_ids(prs):
    """p:sldIdLst 里的原生 SlideID（PowerPoint 明确保证重排后不变）。"""
    ns = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
    sld_ids = prs.slides._sldIdLst.findall("p:sldId", ns)
    return [int(el.get("id")) for el in sld_ids]


def inspect_pptx(path):
    prs = Presentation(path)
    ids = native_slide_ids(prs)
    slides = []
    for index, slide in enumerate(prs.slides):
        slides.append({
            "index": index,
            "slide_id": slide.slide_id,
            "part": str(slide.part.partname),
            "layout": slide.slide_layout.name,
            "shapes": shape_tree(slide.shapes),
        })
    masters = []
    for master in prs.slide_masters:
        masters.append({
            "part": str(master.part.partname),
            "shapes": shape_tree(master.shapes),
            "layouts": [l.name for l in master.slide_layouts],
        })
    return {
        "file": os.path.abspath(path),
        "bytes": os.path.getsize(path),
        "slide_size": [int(prs.slide_width), int(prs.slide_height)],
        "slide_size_in": [round(prs.slide_width / 914400, 3), round(prs.slide_height / 914400, 3)],
        "slide_count": len(prs.slides),
        "sld_id_order": ids,
        "masters": masters,
        "slides": slides,
    }


def _flat_shapes(slide):
    flat = {}

    def walk(nodes):
        for node in nodes:
            if "group" in node:
                walk(node["children"])
            else:
                flat[(node.get("name"), node.get("shape_id"))] = node
    walk(slide.get("shapes", []))
    return flat


def diff_reports(before, after):
    """比较两分报告，输出页面/对象层面的保留与丢失情况。"""
    result = {
        "slide_count": [before["slide_count"], after["slide_count"]],
        "slide_size_changed": before["slide_size"] != after["slide_size"],
        "sld_id_lst_before": before["sld_id_order"],
        "sld_id_lst_after": after["sld_id_order"],
        "slides": [],
    }
    by_id_before = {s["slide_id"]: s for s in before["slides"]}
    for position, slide_after in enumerate(after["slides"]):
        entry = {
            "index": position,
            "slide_id": slide_after["slide_id"],
            "part": slide_after["part"],
            "matched_by_slide_id": slide_after["slide_id"] in by_id_before,
            "shapes_before": 0,
            "shapes_after": len(_flat_shapes(slide_after)),
            "markers_before": [],
            "markers_after": [],
        }
        source = by_id_before.get(slide_after["slide_id"])
        if source is not None:
            src_shapes = _flat_shapes(source)
            entry["shapes_before"] = len(src_shapes)
            entry["markers_before"] = sorted(m["marker"] for m in src_shapes.values() if m.get("marker"))
        entry["markers_after"] = sorted(m["marker"] for m in _flat_shapes(slide_after).values() if m.get("marker"))
        entry["markers_kept"] = sorted(set(entry["markers_after"]) & set(entry["markers_before"]))
        entry["markers_lost"] = sorted(set(entry["markers_before"]) - set(entry["markers_after"]))
        result["slides"].append(entry)
    return result


def main(argv=None):
    parser = argparse.ArgumentParser(description="只读检查 PPTX 结构")
    parser.add_argument("input")
    parser.add_argument("--diff", metavar="OTHER", help="与另一个文件比较")
    parser.add_argument("-o", "--output", help="把报告写到文件")
    args = parser.parse_args(argv)

    report = inspect_pptx(args.input)
    payload = report
    if args.diff:
        payload = {"before": report, "after": inspect_pptx(args.diff),
                   "diff": diff_reports(report, inspect_pptx(args.diff))}
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.output:
        os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)
        with open(args.output, "w", encoding="utf-8") as fh:
            fh.write(text + "\n")
    print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
