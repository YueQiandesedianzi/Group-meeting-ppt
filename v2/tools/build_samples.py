#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""M0 样例生成：带母版/Logo/页脚的参考模板，以及模拟“手工修改”的对照版本。

说明：样例是合成资产，用于验证机制，不等同于真实模板验收（见规格 §14）。
"""
import argparse
import copy
import json
import os
import sys
import uuid

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Inches, Pt
from lxml import etree

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
R_EMBED = "{%s}embed" % NS["r"]
R_ID = "{%s}id" % NS["r"]

PRIMARY = RGBColor(0x17, 0x40, 0x6D)
ACCENT = RGBColor(0x0F, 0x6F, 0xC6)
MUTED = RGBColor(0x80, 0x80, 0x80)
LIGHT = RGBColor(0xDB, 0xEF, 0xF9)
FONT_CJK = "微软雅黑"


def marker(prefix, name):
    return "%s%s=%s" % (prefix, name, uuid.uuid4())


def set_marker(shape, value):
    shape._element._nvXxPr.cNvPr.set("descr", value)


def make_images(assets_dir):
    os.makedirs(assets_dir, exist_ok=True)
    logo = os.path.join(assets_dir, "logo.png")
    img = Image.new("RGBA", (360, 96), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle([0, 8, 359, 88], radius=16, fill="#17406D")
    try:
        font = ImageFont.truetype(r"C:\Windows\Fonts\msyhbd.ttc", 34)
    except Exception:
        font = ImageFont.load_default()
    draw.text((24, 28), "ACME LAB", font=font, fill="#FFFFFF")
    img.save(logo)

    chart1 = os.path.join(assets_dir, "chart-accuracy.png")
    img = Image.new("RGB", (800, 450), "white")
    draw = ImageDraw.Draw(img)
    draw.title = None
    draw.rectangle([0, 0, 799, 449], outline="#C9C9C9")
    for i, (label, value) in enumerate([("baseline", 0.882), ("method A", 0.914), ("method B", 0.907)]):
        x = 90 + i * 240
        h = int(value * 400)
        draw.rectangle([x, 420 - h, x + 130, 420], fill="#0F6FC6")
        draw.text((x + 10, 428), label, fill="#333333")
        draw.text((x + 10, 420 - h - 20), "%.3f" % value, fill="#17406D")
    draw.text((20, 12), "test accuracy", fill="#17406D")
    img.save(chart1)

    chart2 = os.path.join(assets_dir, "chart-loss.png")
    img = Image.new("RGB", (640, 480), "white")
    draw = ImageDraw.Draw(img)
    draw.rectangle([0, 0, 639, 479], outline="#C9C9C9")
    points = [(30 + i * 60, 420 - int(380 * (1 - 0.9 ** i))) for i in range(10)]
    draw.line(points, fill="#0F6FC6", width=4)
    draw.text((20, 12), "training loss", fill="#17406D")
    img.save(chart2)
    return logo, chart1, chart2


def add_textbox(slide, text, left, top, width, height, size=18, bold=False, color=None, align=None):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    frame = box.text_frame
    frame.word_wrap = True
    paragraph = frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = FONT_CJK
    if color is not None:
        run.font.color.rgb = color
    if align is not None:
        from pptx.enum.text import PP_ALIGN
        paragraph.alignment = align
    return box


def add_bullets(slide, title, items, left=0.9, top=1.6):
    add_textbox(slide, title, left, 0.9, 11.5, 0.8, size=26, bold=True, color=PRIMARY)
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(11.5), Inches(5.0))
    frame = box.text_frame
    frame.word_wrap = True
    for index, item in enumerate(items):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        run = paragraph.add_run()
        run.text = "• " + item
        run.font.size = Pt(18)
        run.font.name = FONT_CJK
    return box


def _append_to_master(master, temp_slide, element):
    new_element = copy.deepcopy(element)
    if etree.QName(new_element).localname == "pic":
        blip = new_element.find(".//a:blip", NS)
        image_part = temp_slide.part.related_part(blip.get(R_EMBED))
        new_rid = master.part.relate_to(image_part, RT.IMAGE)
        blip.set(R_EMBED, new_rid)
    nv = new_element.find(".//p:cNvPr", NS)
    if nv is not None:
        existing = [int(e.get("id")) for e in master.shapes._spTree.iter(
            "{%s}cNvPr" % NS["p"]) if e.get("id") and e.get("id").isdigit()]
        nv.set("id", str(max(existing + [1]) + 1))
    master.shapes._spTree.append(new_element)
    return new_element


def decorate_master(prs, logo_png):
    """在母版上放 Logo、页脚与装饰线，验证继承关系。"""
    temp = prs.slides.add_slide(prs.slide_layouts[6])
    logo = temp.shapes.add_picture(logo_png, Inches(0.45), Inches(0.28), height=Inches(0.42))
    set_marker(logo, marker("gmppt:master-logo:", "logo"))
    line = temp.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.45), Inches(0.86), Inches(12.4), Pt(1.6))
    line.fill.solid()
    line.fill.fore_color.rgb = ACCENT
    line.line.fill.background()
    footer = add_textbox(temp, "ACME LAB · 内部资料 · 组会汇报", 0.45, 7.55, 8.0, 0.4,
                         size=10, color=MUTED)
    set_marker(footer, marker("gmppt:master-footer:", "footer"))

    for element in (logo._element, line._element, footer._element):
        _append_to_master(prs.slide_master, temp, element)

    # 删除临时页（只用它来承载关系与形状）
    rel_id = None
    for rid, rel in prs.part.rels.items():
        if rel.reltype.endswith("/slide") and rel.target_part is temp.part:
            rel_id = rid
            break
    if rel_id:
        prs.part.drop_rel(rel_id)
    sld_ids = prs.slides._sldIdLst
    for sld_id in list(sld_ids):
        if sld_id.get(R_ID) == rel_id:
            sld_ids.remove(sld_id)


def build_reference(out_path, assets_dir, mapping_path):
    logo, chart1, chart2 = make_images(assets_dir)
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(8.333)
    decorate_master(prs, logo)
    blank = prs.slide_layouts[6]
    report = {"pages": []}

    # 1 封面
    slide = prs.slides.add_slide(blank)
    title = add_textbox(slide, "课题进展汇报", 1.2, 2.6, 11.0, 1.2, size=40, bold=True, color=PRIMARY)
    set_marker(title, marker("gmppt:page-title:", "cover-title"))
    add_textbox(slide, "ACME LAB 组会 · 2026-11", 1.2, 3.9, 8.0, 0.6, size=18, color=MUTED)
    add_textbox(slide, "汇报人：张三", 1.2, 4.5, 8.0, 0.6, size=16, color=MUTED)
    report["pages"].append({"index": 1, "use": "cover"})

    # 2 正文 + 页级固定元素
    slide = prs.slides.add_slide(blank)
    heading = add_bullets(slide, "研究动机", [
        "现有方法在受限数据下泛化不足",
        "拟引入 Z 结构提升稳定性",
    ])
    set_marker(heading, marker("gmppt:page-title:", "body-title"))
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.0), Inches(8.03), Inches(13.333), Inches(0.30))
    band.fill.solid()
    band.fill.fore_color.rgb = LIGHT
    band.line.fill.background()
    fixed = add_textbox(slide, "固定说明：本页页脚为示例页固定元素，不应被内容更新清除", 0.45, 8.06, 9.5, 0.26,
                        size=9, color=PRIMARY)
    set_marker(fixed, marker("gmppt:fixed:", "page-band"))
    report["pages"].append({"index": 2, "use": "body"})

    # 3 图文
    slide = prs.slides.add_slide(blank)
    heading = add_textbox(slide, "精度对比", 0.9, 0.9, 11.5, 0.8, size=26, bold=True, color=PRIMARY)
    set_marker(heading, marker("gmppt:page-title:", "figure-title"))
    picture = slide.shapes.add_picture(chart1, Inches(1.0), Inches(2.0), width=Inches(7.2))
    set_marker(picture, marker("gmppt:picture:", "result-chart"))
    add_textbox(slide, "图注：方法 A 较基线提升 3.2 个百分点", 1.0, 6.55, 7.2, 0.5, size=12, color=MUTED)
    add_textbox(slide, "要点：Z 结构在验证集稳定提升", 8.6, 2.2, 4.0, 1.6, size=14)
    report["pages"].append({"index": 3, "use": "figure"})

    # 4 双图
    slide = prs.slides.add_slide(blank)
    heading = add_textbox(slide, "训练过程与损失", 0.9, 0.9, 11.5, 0.8, size=26, bold=True, color=PRIMARY)
    set_marker(heading, marker("gmppt:page-title:", "two-figure-title"))
    left = slide.shapes.add_picture(chart1, Inches(0.9), Inches(2.0), width=Inches(5.6))
    set_marker(left, marker("gmppt:picture:", "two-figure-left"))
    right = slide.shapes.add_picture(chart2, Inches(6.9), Inches(2.0), width=Inches(5.4))
    set_marker(right, marker("gmppt:picture:", "two-figure-right"))
    report["pages"].append({"index": 4, "use": "two-figure"})

    # 5 总结（简单表格 + 要点）
    slide = prs.slides.add_slide(blank)
    heading = add_textbox(slide, "小结", 0.9, 0.9, 11.5, 0.8, size=26, bold=True, color=PRIMARY)
    set_marker(heading, marker("gmppt:page-title:", "summary-title"))
    table = slide.shapes.add_table(3, 3, Inches(0.9), Inches(2.0), Inches(7.0), Inches(1.8)).table
    for column, text in enumerate(["方法", "精度", "参数量"]):
        table.cell(0, column).text = text
    for row, values in enumerate([["基线", "88.2%", "18M"], ["方法 A", "91.4%", "12M"]], start=1):
        for column, text in enumerate(values):
            table.cell(row, column).text = text
    add_textbox(slide, "结论：Z 结构有效，参数量下降 33%", 8.4, 2.1, 4.2, 1.6, size=14)
    report["pages"].append({"index": 5, "use": "summary"})

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    prs.save(out_path)

    mapping = {
        "template": os.path.basename(out_path),
        "slide_size_in": [13.333, 8.333],
        "mode": "sample-pages",
        "fixed": {
            "master": ["logo", "footer", "accent-line"],
            "page": {"2": ["page-band"]},
        },
        "pages": [
            {"index": 1, "use": "cover", "regions": {"title": "page-title", "meta": "static-text"}},
            {"index": 2, "use": "body", "regions": {"title": "page-title", "content": "bullets"}},
            {"index": 3, "use": "figure", "regions": {"title": "page-title", "image": "picture", "caption": "static-text"}},
            {"index": 4, "use": "two-figure", "regions": {"title": "page-title", "images": "picture"}},
            {"index": 5, "use": "summary", "regions": {"title": "page-title", "table": "table"}},
        ],
        "free_regions": [],
    }
    with open(mapping_path, "w", encoding="utf-8") as fh:
        json.dump(mapping, fh, ensure_ascii=False, indent=2)
    return report


def _move_slide(prs, old_index, new_index):
    sld_ids = prs.slides._sldIdLst
    element = sld_ids[old_index]
    sld_ids.remove(element)
    sld_ids.insert(new_index, element)


def _duplicate_text_slide(prs, source_index, blank):
    """复制一页（仅含文本/表格，无图片关系），模拟用户在 PowerPoint 里复制页面。"""
    source = prs.slides[source_index]
    new_slide = prs.slides.add_slide(blank)
    for shape in list(new_slide.shapes):
        shape._element.getparent().remove(shape._element)
    for element in source.shapes._spTree:
        if etree.QName(element).localname in ("nvGrpSpPr", "grpSpPr"):
            continue
        new_slide.shapes._spTree.append(copy.deepcopy(element))
    return new_slide


def build_handedited(reference, out_path):
    prs = Presentation(reference)
    blank = prs.slide_layouts[6]

    # 1) 改封面标题
    cover = prs.slides[0]
    for shape in cover.shapes:
        if shape.has_text_frame and "课题进展汇报" in shape.text_frame.text:
            shape.text_frame.paragraphs[0].runs[0].text = "【手改】课题进展汇报"
            break

    # 2) 移动 + 缩放第 3 页图片，并改图注
    figure_page = prs.slides[2]
    for shape in figure_page.shapes:
        if shape.shape_type == 13:  # PICTURE
            shape.left = shape.left + Inches(0.4)
            shape.top = shape.top - Inches(0.3)
            shape.width = int(shape.width * 1.1)
            break
    for shape in figure_page.shapes:
        if shape.has_text_frame and "图注" in shape.text_frame.text:
            shape.text_frame.paragraphs[0].runs[0].text = "图注（手改）：方法 A 提升 3.2pp"
            break

    # 3) 第 4 页加箭头与说明框（自由元素）
    two_figure = prs.slides[3]
    arrow = two_figure.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(6.35), Inches(3.4), Inches(0.5), Inches(0.4))
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = RGBColor(0xE0, 0x6C, 0x00)
    arrow.line.fill.background()
    set_marker(arrow, marker("gmppt:free:", "user-arrow"))
    note = add_textbox(two_figure, "手改说明：左侧为消融结果", 5.2, 5.9, 4.0, 0.5, size=12, color=MUTED)
    set_marker(note, marker("gmppt:free:", "user-note"))

    # 4) 重排：把第 5 页移到第 2 位
    _move_slide(prs, 4, 1)

    # 5) 复制一页（保留原有身份标记，模拟 PowerPoint 复制自定义数据）
    _duplicate_text_slide(prs, 0, blank)

    # 6) 新增自由页
    free = prs.slides.add_slide(blank)
    box = free.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.6), Inches(5.0), Inches(2.0))
    box.fill.solid()
    box.fill.fore_color.rgb = LIGHT
    box.line.color.rgb = ACCENT
    box.text_frame.text = "用户自由新增页"
    set_marker(box, marker("gmppt:free:", "free-page-box"))

    prs.save(out_path)
    return out_path


def main(argv=None):
    parser = argparse.ArgumentParser(description="生成 M0 样板")
    parser.add_argument("--outdir", default=os.path.join(os.path.dirname(os.path.dirname(
        os.path.abspath(__file__))), "samples"))
    args = parser.parse_args(argv)

    samples = args.outdir
    assets = os.path.join(samples, "assets")
    reference = os.path.join(samples, "reference-template.pptx")
    mapping = os.path.join(samples, "template-map.json")
    hand = os.path.join(samples, "hand-edited.pptx")

    report = build_reference(reference, assets, mapping)
    build_handedited(reference, hand)
    print(json.dumps({"reference": reference, "hand_edited": hand,
                      "mapping": mapping, "pages": report["pages"]},
                     ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
