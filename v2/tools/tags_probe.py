#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PowerPoint Tags 往返试验：验证规格 §6.1 建议的正式身份载体能否存活。

Tags 存放在独立部件（ppt/tags/tagN.xml），由对象 p:nvPr/p:custDataLst/p:tags 以 r:id 引用。
本工具创建 Tags、读回校验、经客户端保存后再比较；结果写入 JSON 报告。
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
import uuid

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pptx_inspect as pi  # noqa: E402
import roundtrip_check as rc  # noqa: E402

from lxml import etree  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.opc.package import Part  # noqa: E402
from pptx.opc.packuri import PackURI  # noqa: E402

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
TAGS_CT = "application/vnd.openxmlformats-officedocument.presentationml.tags+xml"


def _tags_part(slide_part, package, partname="/ppt/tags/tag1.xml"):
    """取得或创建该幻灯片的 tags 部件（同一页的多个对象共享一个 tags 部件）。"""
    for rid, rel in slide_part.rels.items():
        if rel.reltype.endswith("/tags"):
            return rel.target_part, rid
    blob = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<p:tagLst xmlns:p="%s"/>' % P).encode("utf-8")
    part = Part(PackURI(partname), TAGS_CT, package, blob)
    rid = slide_part.relate_to(part, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags")
    return part, rid


def add_shape_tags(shape, mapping, package, partname):
    """把 {name: value} 写进该对象的 tags 部件，并在 nvPr 中挂引用。"""
    slide_part = shape.part
    part, rid = _tags_part(slide_part, package, partname)
    root = etree.fromstring(part.blob)
    existing = {tag.get("name") for tag in root.findall("{%s}tag" % P)}
    for name, value in mapping.items():
        if name in existing:
            for tag in root.findall("{%s}tag" % P):
                if tag.get("name") == name:
                    tag.set("val", value)
            continue
        tag = etree.SubElement(root, "{%s}tag" % P)
        tag.set("name", name)
        tag.set("val", value)
    part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    nv_pr = shape._element._nvXxPr.nvPr
    cust = nv_pr.find("{%s}custDataLst" % P)
    if cust is None:
        cust = etree.SubElement(nv_pr, "{%s}custDataLst" % P)
    if cust.find("{%s}tags" % P) is None:
        tags = etree.SubElement(cust, "{%s}tags" % P)
        tags.set("{%s}id" % R, rid)
    return rid


def tag_map(report):
    """marker → tags，便于前后比较。"""
    out = {}
    for slide in report["slides"]:
        for shape in slide["shapes"]:
            marker = shape.get("marker")
            if marker and shape.get("tags"):
                out[marker] = shape["tags"]
    return out


def main(argv=None):
    parser = argparse.ArgumentParser(description="PowerPoint Tags 往返试验")
    parser.add_argument("--input", default=os.path.join(os.path.dirname(os.path.dirname(
        os.path.abspath(__file__))), "samples", "reference-template.pptx"))
    parser.add_argument("--outdir", default=os.path.join(os.path.dirname(os.path.dirname(
        os.path.abspath(__file__))), "samples"))
    parser.add_argument("--report", default=os.path.join(os.path.dirname(os.path.dirname(
        os.path.abspath(__file__))), "reports", "tags-roundtrip.json"))
    args = parser.parse_args(argv)

    tagged = os.path.join(args.outdir, "tags-sample.pptx")
    shutil.copyfile(args.input, tagged)

    project_id = str(uuid.uuid4())
    prs = Presentation(tagged)
    tagged_shapes = 0
    for slide_index, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            try:
                descr = shape._element._nvXxPr.cNvPr.get("descr") or ""
            except Exception:
                continue
            if not descr.startswith("gmppt:"):
                continue
            add_shape_tags(shape, {
                "gmppt.project": project_id,
                "gmppt.slide": "slide-%02d" % (slide_index + 1),
                "gmppt.object": descr.split("=")[-1],
            }, prs.part.package, "/ppt/tags/tag%02d.xml" % (slide_index + 1))
            tagged_shapes += 1
    prs.save(tagged)

    before = pi.inspect_pptx(tagged)
    before_tags = tag_map(before)

    converted = os.path.join(args.outdir, "tags-sample.roundtrip.pptx")
    meta = rc.convert(rc.DEFAULT_NODE, rc.DEFAULT_CLI, tagged, converted)
    after = pi.inspect_pptx(converted)
    after_tags = tag_map(after)

    tags_part_before = [n for n in _zip_names(tagged) if n.startswith("ppt/tags/")]
    tags_part_after = [n for n in _zip_names(converted) if n.startswith("ppt/tags/")]

    kept = sorted(set(before_tags) & set(after_tags))
    equal = [m for m in kept if before_tags[m] == after_tags[m]]
    report = {
        "client": "libreoffice-kit",
        "client_is_powerpoint": False,
        "tagged_shapes": tagged_shapes,
        "tags_parts_before": tags_part_before,
        "tags_parts_after": tags_part_after,
        "tags_readable_before": len(before_tags),
        "tags_readable_after": len(after_tags),
        "markers_kept": len(kept),
        "tags_identical": len(equal),
        "tags_changed": [m for m in kept if before_tags[m] != after_tags[m]],
        "tags_lost": sorted(set(before_tags) - set(after_tags)),
        "convert_meta": meta,
    }
    text = json.dumps(report, ensure_ascii=False, indent=2)
    os.makedirs(os.path.dirname(os.path.abspath(args.report)), exist_ok=True)
    with open(args.report, "w", encoding="utf-8") as fh:
        fh.write(text + "\n")
    print(text)
    return 0


def _zip_names(path):
    import zipfile
    with zipfile.ZipFile(path) as zf:
        return zf.namelist()


if __name__ == "__main__":
    sys.exit(main())
