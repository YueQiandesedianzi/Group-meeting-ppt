#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""组会 PPT 生成脚本：读 brief.yaml → 校验 → 按模板渲染 .pptx。

用法：
    python generate_ppt.py <input.yaml> -o <out.pptx> [--template <name>]
    python generate_ppt.py <input.yaml> -o <out.pptx> --slide N      # 只重建第 N 页并原地替换
    python generate_ppt.py <input.yaml> -o <out.pptx> --history-dir <dir>  # 同时写快照 + 轻量预览
"""
import argparse
import copy
import hashlib
import json
import math
import os
import shutil
import sys
import tempfile
from datetime import datetime

import yaml
from PIL import Image as PILImage

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
from lxml import etree

REQUIRED = {
    "cover": ["title"],
    "agenda": ["items"],
    "section_header": ["heading"],
    "points": ["heading", "points"],
    "figure": ["heading", "image"],
    "metrics": ["heading", "metrics"],
    "table": ["heading", "columns", "rows"],
    "timeline": ["heading", "steps"],
    "conclusion": ["heading", "points"],
    "next_steps": ["heading", "items"],
}
ALLOWED_TYPES = set(REQUIRED)

SNAPSHOT_FORMAT_VERSION = 1
SNAPSHOT_SKILL = "group-meeting-ppt"
MX = 0.6  # 左右边距（英寸）


class Theme:
    def __init__(self, cfg):
        self.name = cfg.get("name", "default")
        self.colors = cfg.get("colors", {})
        self.fonts = cfg.get("fonts", {})
        slide = cfg.get("slide", {})
        self.W = float(slide.get("width_in", 13.333))
        self.H = float(slide.get("height_in", 7.5))

    def c(self, key):
        hx = (self.colors.get(key) or "#333333").lstrip("#")
        return RGBColor.from_string(hx)

    def ea(self):
        return self.fonts.get("cjk", "微软雅黑")

    def latin(self):
        return self.fonts.get("latin", "Calibri")


# ---------------- 基础工具 ----------------

def load_yaml(path):
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def file_hash(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def load_yaml_with_hash(path):
    """从同一字节快照取得 YAML 和哈希，避免读取与哈希错开。"""
    with open(path, "rb") as f:
        raw = f.read()
    return yaml.safe_load(raw.decode("utf-8")), hashlib.sha256(raw).hexdigest()


def default_template_dir():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.normpath(os.path.join(script_dir, "..", "templates"))


def image_path(reference, base_dir):
    if os.path.isabs(reference):
        return os.path.abspath(reference)
    return os.path.abspath(os.path.join(base_dir, reference))


def collect_image_dependencies(brief, base_dir):
    """记录原始图片路径和生成前内容哈希；路径不会写入快照。"""
    dependencies = []
    for section_index, sec in enumerate(brief["sections"], 1):
        if sec.get("type") != "figure":
            continue
        source = image_path(sec["image"], base_dir)
        dependencies.append({
            "section": section_index,
            "source": source,
            "sha256": file_hash(source),
        })
    return dependencies


def describe_inputs(brief_path, template_path, image_dependencies,
                    brief_hash=None, template_hash=None):
    """生成可复现输入的摘要，不含本机源路径。"""
    return {
        "brief_sha256": brief_hash if brief_hash is not None else file_hash(brief_path),
        "template_sha256": template_hash if template_hash is not None else file_hash(template_path),
        "generator_sha256": file_hash(__file__),
        "images": [
            {"section": dep["section"], "sha256": dep["sha256"]}
            for dep in image_dependencies
        ],
    }


def normalized_output_path(output):
    return os.path.normcase(os.path.realpath(os.path.abspath(output)))


def generation_fingerprint(source_info, output):
    payload = {
        "format_version": SNAPSHOT_FORMAT_VERSION,
        "brief": source_info["brief_sha256"],
        "template": source_info["template_sha256"],
        "generator": source_info["generator_sha256"],
        "images": source_info["images"],
        "output": normalized_output_path(output),
    }
    data = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(data.encode("utf-8")).hexdigest()


def generation_marker(source_info, output):
    marker = dict(source_info)
    marker.update({
        "format_version": SNAPSHOT_FORMAT_VERSION,
        "status": "full",
        "output": normalized_output_path(output),
        "output_sha256": file_hash(output),
        "fingerprint": generation_fingerprint(source_info, output),
    })
    return marker


def mark_partial_generation(history_dir, output, source_info, slide):
    """输出已经局部变化，撤销其“全量最新”资格，直到下次全量成功。"""
    marker = dict(source_info)
    marker.update({
        "format_version": SNAPSHOT_FORMAT_VERSION,
        "status": "partial",
        "requires_full_rebuild": True,
        "slide": slide,
        "output": normalized_output_path(output),
        "output_sha256": file_hash(output),
        "fingerprint": generation_fingerprint(source_info, output),
    })
    write_json_atomic(os.path.join(history_dir, ".last_generated.json"), marker)


def mark_history_pending(history_dir, output, source_info):
    """全量 PPT 已保存但历史尚未提交，旧 full 不再代表当前输出。"""
    marker = dict(source_info)
    marker.update({
        "format_version": SNAPSHOT_FORMAT_VERSION,
        "status": "history_pending",
        "requires_full_rebuild": True,
        "output": normalized_output_path(output),
        "output_sha256": file_hash(output),
        "fingerprint": generation_fingerprint(source_info, output),
    })
    try:
        with open(os.path.join(history_dir, ".last_generated.json"),
                  "r", encoding="utf-8") as f:
            previous = json.load(f)
    except (OSError, ValueError, TypeError):
        previous = None
    if isinstance(previous, dict):
        if previous.get("status") == "full":
            marker["last_full"] = previous
        elif isinstance(previous.get("last_full"), dict):
            marker["last_full"] = previous["last_full"]
    write_json_atomic(os.path.join(history_dir, ".last_generated.json"), marker)


def should_skip_generation(history_dir, output, source_info):
    """仅完整、同输入、同输出的上次全量生成允许跳过。"""
    path = os.path.join(history_dir, ".last_generated.json")
    try:
        with open(path, "r", encoding="utf-8") as f:
            previous = json.load(f)
    except (OSError, ValueError, TypeError):
        return False
    if not isinstance(previous, dict) or previous.get("status") != "full":
        return False
    if not previous.get("output_sha256"):
        return False
    return (os.path.isfile(output)
            and previous.get("output") == normalized_output_path(output)
            and previous.get("fingerprint") == generation_fingerprint(source_info, output)
            and file_hash(output) == previous["output_sha256"])


def write_json_atomic(path, data):
    """在同一目录写完临时文件后替换，避免留下半个 JSON。"""
    parent = os.path.dirname(os.path.abspath(path))
    os.makedirs(parent, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix=".tmp-", suffix=".json", dir=parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.write("\n")
        os.replace(temporary, path)
    except Exception:
        try:
            os.close(fd)
        except OSError:
            pass
        if os.path.exists(temporary):
            os.remove(temporary)
        raise


def read_snapshot_manifest(version_dir):
    path = os.path.join(version_dir, "manifest.json")
    if not os.path.isfile(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            manifest = json.load(f)
    except (OSError, ValueError) as e:
        raise ValueError(f"无法读取历史清单 {path}：{e}") from e
    if not isinstance(manifest, dict) or manifest.get("skill") != SNAPSHOT_SKILL:
        raise ValueError(f"不是 {SNAPSHOT_SKILL} 的历史清单：{path}")
    if manifest.get("format_version") != SNAPSHOT_FORMAT_VERSION:
        raise ValueError(f"不支持的历史格式版本：{manifest.get('format_version')!r}")
    return manifest


def resolve_template_path(brief_path, template_name=None, template_dir=None):
    """显式模板优先；新快照在未指定模板时使用归档模板。"""
    if template_name:
        return os.path.join(template_dir or default_template_dir(), template_name + ".yaml")
    # 是否覆盖历史模板只由显式模板名决定，外部模板目录不能改变默认选择。
    version_dir = os.path.dirname(os.path.abspath(brief_path))
    if read_snapshot_manifest(version_dir) is not None:
        template_path = os.path.join(version_dir, "template.yaml")
        if not os.path.isfile(template_path):
            raise ValueError(f"历史版本缺少归档模板：{template_path}")
        return template_path
    return os.path.join(template_dir or default_template_dir(), "default.yaml")


def validate_snapshot_inputs(brief_path, brief, template_path):
    """重建前核对新快照的实际字节；损坏的资源不得静默用于生成。"""
    version_dir = os.path.dirname(os.path.abspath(brief_path))
    manifest = read_snapshot_manifest(version_dir)
    if manifest is None:
        return
    if file_hash(brief_path) != manifest.get("snapshot_brief_sha256"):
        raise ValueError("历史 brief 已改变，文件哈希与 manifest.json 不一致")
    archived_template = os.path.join(version_dir, "template.yaml")
    if (os.path.normcase(os.path.realpath(template_path))
            == os.path.normcase(os.path.realpath(archived_template))):
        if not os.path.isfile(archived_template):
            raise ValueError("历史版本缺少归档模板")
        if file_hash(archived_template) != manifest.get("template", {}).get("sha256"):
            raise ValueError("历史模板已改变，文件哈希与 manifest.json 不一致")
    for entry in manifest.get("images", []):
        try:
            section = brief["sections"][entry["section"] - 1]
            path = image_path(section["image"], version_dir)
            expected = entry["sha256"]
        except (KeyError, IndexError, TypeError) as e:
            raise ValueError("历史 manifest.json 中的图片索引无效") from e
        if not os.path.isfile(path):
            raise ValueError(f"历史图片资产不存在：第 {entry['section']} 段")
        if file_hash(path) != expected:
            raise ValueError(f"历史图片已改变，第 {entry['section']} 段的文件哈希不一致")


def set_run(r, text, th, size=16, bold=False, color=None):
    r.text = text
    f = r.font
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = color if color is not None else th.c("text")
    f.name = th.latin()
    rPr = r._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = etree.SubElement(rPr, qn("a:ea"))
    ea.set("typeface", th.ea())


def add_box(slide, x, y, w, h, fill=None, line=None, line_w=None,
            shape=MSO_SHAPE.ROUNDED_RECTANGLE):
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w or 0.75)
    try:
        sp.shadow.inherit = False
    except Exception:
        pass
    return sp


def add_text(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    return tf


def draw_heading(slide, th, text):
    CW = th.W - 2 * MX
    tf = add_text(slide, MX, 0.45, CW, 0.7)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    set_run(p.add_run(), text or "", th, size=24, bold=True, color=th.c("primary"))
    add_box(slide, MX, 1.12, 1.6, 0.045, fill=th.c("accent"), shape=MSO_SHAPE.RECTANGLE)


# ---------------- 10 类版式 ----------------

def draw_cover(slide, sec, th):
    CW = th.W - 2 * MX
    title = sec.get("title") or ""
    tf = add_text(slide, MX, th.H * 0.30, CW, 1.6, anchor=MSO_ANCHOR.MIDDLE)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    set_run(p.add_run(), title, th, size=36, bold=True, color=th.c("primary"))
    y = th.H * 0.30 + 1.6
    if sec.get("subtitle"):
        tf = add_text(slide, MX, y, CW, 0.7, anchor=MSO_ANCHOR.MIDDLE)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        set_run(p.add_run(), sec["subtitle"], th, size=18, color=th.c("muted"))
        y += 0.7
    meta = " · ".join(str(x) for x in [sec.get("presenter"), sec.get("date")] if x)
    if meta:
        tf = add_text(slide, MX, th.H - 1.3, CW, 0.5, anchor=MSO_ANCHOR.MIDDLE)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        set_run(p.add_run(), meta, th, size=12, color=th.c("muted"))
    add_box(slide, MX, th.H - 0.5, CW, 0.05, fill=th.c("accent"), shape=MSO_SHAPE.RECTANGLE)


def draw_agenda(slide, sec, th):
    CW = th.W - 2 * MX
    draw_heading(slide, th, sec.get("heading", "汇报提纲"))
    y = 1.5
    for i, it in enumerate(sec.get("items", []), 1):
        chip = add_box(slide, MX, y + 0.02, 0.42, 0.42, fill=th.c("accent"))
        ctf = chip.text_frame
        ctf.word_wrap = False
        ctf.margin_left = 0
        ctf.margin_right = 0
        ctf.margin_top = 0
        ctf.margin_bottom = 0
        ctf.vertical_anchor = MSO_ANCHOR.MIDDLE
        cp = ctf.paragraphs[0]
        cp.alignment = PP_ALIGN.CENTER
        set_run(cp.add_run(), str(i), th, size=13, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
        tf = add_text(slide, MX + 0.65, y, CW - 0.65, 0.5, anchor=MSO_ANCHOR.MIDDLE)
        p = tf.paragraphs[0]
        set_run(p.add_run(), str(it), th, size=16)
        y += 0.72


def draw_section_header(slide, sec, th):
    CW = th.W - 2 * MX
    tf = add_text(slide, MX, th.H * 0.38, CW, 1.4, anchor=MSO_ANCHOR.MIDDLE)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    set_run(p.add_run(), sec.get("heading", ""), th, size=32, bold=True, color=th.c("primary"))
    if sec.get("subtitle"):
        tf = add_text(slide, MX, th.H * 0.38 + 1.4, CW, 0.6, anchor=MSO_ANCHOR.MIDDLE)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        set_run(p.add_run(), sec["subtitle"], th, size=16, color=th.c("muted"))
    add_box(slide, (th.W - 2.4) / 2, th.H * 0.38 + 1.5, 2.4, 0.05,
            fill=th.c("accent"), shape=MSO_SHAPE.RECTANGLE)


def draw_points(slide, sec, th):
    CW = th.W - 2 * MX
    draw_heading(slide, th, sec.get("heading", ""))
    h = th.H - 3.1 if sec.get("emphasis") else th.H - 1.7
    tf = add_text(slide, MX, 1.45, CW, h, anchor=MSO_ANCHOR.TOP)
    first = True
    for pt in sec.get("points", []):
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(10)
        set_run(p.add_run(), "• ", th, size=16, bold=True, color=th.c("accent"))
        set_run(p.add_run(), str(pt), th, size=16)
    if sec.get("emphasis"):
        ey = th.H - 1.45
        add_box(slide, MX, ey, 0.08, 0.85, fill=th.c("accent"), shape=MSO_SHAPE.RECTANGLE)
        add_box(slide, MX + 0.08, ey, CW - 0.08, 0.85, fill=th.c("light"), line=None)
        tf = add_text(slide, MX + 0.35, ey, CW - 0.6, 0.85, anchor=MSO_ANCHOR.MIDDLE)
        p = tf.paragraphs[0]
        set_run(p.add_run(), sec["emphasis"], th, size=15, bold=True, color=th.c("primary"))


def draw_figure(slide, sec, th, base_dir):
    CW = th.W - 2 * MX
    draw_heading(slide, th, sec.get("heading", ""))
    img_path = image_path(sec["image"], base_dir)
    left_w = CW * 0.58
    right_x = MX + left_w + 0.4
    right_w = CW - left_w - 0.4
    with PILImage.open(img_path) as im:
        iw, ih = im.size
    max_w, max_h = left_w, th.H - 2.0
    ratio = min(max_w / iw, max_h / ih)
    w, h = iw * ratio, ih * ratio
    ix = MX + (max_w - w) / 2
    iy = 1.4 + (max_h - h) / 2
    slide.shapes.add_picture(img_path, Inches(ix), Inches(iy), width=Inches(w), height=Inches(h))
    if sec.get("caption"):
        tf = add_text(slide, MX, iy + h + 0.08, left_w, 0.4, anchor=MSO_ANCHOR.TOP)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        set_run(p.add_run(), sec["caption"], th, size=12, color=th.c("muted"))
    tf = add_text(slide, right_x, 1.45, right_w, th.H - 1.6, anchor=MSO_ANCHOR.TOP)
    first = True
    for t in sec.get("takeaways", []):
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(8)
        set_run(p.add_run(), "• ", th, size=15, bold=True, color=th.c("accent"))
        set_run(p.add_run(), str(t), th, size=15)
    for m in sec.get("metrics", []):
        p = tf.add_paragraph()
        p.space_before = Pt(10)
        set_run(p.add_run(), f"{m.get('label', '')}：", th, size=14, bold=True, color=th.c("muted"))
        set_run(p.add_run(), str(m.get("value", "")), th, size=16, bold=True, color=th.c("primary"))


def draw_metrics(slide, sec, th):
    CW = th.W - 2 * MX
    draw_heading(slide, th, sec.get("heading", ""))
    ms = sec.get("metrics", [])
    n = len(ms)
    cols = 3 if n > 4 else (2 if n > 1 else 1)
    rows = math.ceil(n / cols)
    top, bottom = 1.6, th.H - 0.6
    gap = 0.25
    cw = (CW - (cols - 1) * gap) / cols
    ch = (bottom - top - (rows - 1) * gap) / rows
    for i, m in enumerate(ms):
        r, c = divmod(i, cols)
        x = MX + c * (cw + gap)
        y = top + r * (ch + gap)
        add_box(slide, x, y, cw, ch, fill=th.c("light"), line=None)
        vtf = add_text(slide, x + 0.2, y + ch * 0.16, cw - 0.4, ch * 0.4, anchor=MSO_ANCHOR.MIDDLE)
        vp = vtf.paragraphs[0]
        vp.alignment = PP_ALIGN.CENTER
        set_run(vp.add_run(), str(m.get("value", "")), th, size=28, bold=True, color=th.c("primary"))
        ltf = add_text(slide, x + 0.2, y + ch * 0.54, cw - 0.4, ch * 0.26, anchor=MSO_ANCHOR.MIDDLE)
        lp = ltf.paragraphs[0]
        lp.alignment = PP_ALIGN.CENTER
        set_run(lp.add_run(), str(m.get("label", "")), th, size=13, bold=True, color=th.c("text"))
        if m.get("note"):
            ntf = add_text(slide, x + 0.2, y + ch * 0.78, cw - 0.4, ch * 0.2, anchor=MSO_ANCHOR.MIDDLE)
            np_ = ntf.paragraphs[0]
            np_.alignment = PP_ALIGN.CENTER
            set_run(np_.add_run(), str(m["note"]), th, size=11, color=th.c("muted"))


def draw_table(slide, sec, th):
    CW = th.W - 2 * MX
    draw_heading(slide, th, sec.get("heading", ""))
    cols = sec.get("columns", [])
    rows = sec.get("rows", [])
    ncols = len(cols)
    nrows = len(rows) + 1
    gf = slide.shapes.add_table(nrows, ncols, Inches(MX), Inches(1.5), Inches(CW), Inches(4.6))
    tbl = gf.table
    tbl.first_row = False
    tbl.horz_banding = False
    for j in range(ncols):
        tbl.columns[j].width = Inches(CW / ncols)
    for j, cname in enumerate(cols):
        cell = tbl.cell(0, j)
        cell.text = ""
        p = cell.text_frame.paragraphs[0]
        set_run(p.add_run(), str(cname), th, size=14, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
        cell.fill.solid()
        cell.fill.fore_color.rgb = th.c("accent")
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.text = ""
            p = cell.text_frame.paragraphs[0]
            set_run(p.add_run(), str(val), th, size=13, color=th.c("text"))
            cell.fill.solid()
            cell.fill.fore_color.rgb = th.c("light") if i % 2 == 0 else RGBColor(0xFF, 0xFF, 0xFF)


def draw_timeline(slide, sec, th):
    CW = th.W - 2 * MX
    draw_heading(slide, th, sec.get("heading", ""))
    steps = sec.get("steps", [])
    if not steps:
        return
    line_y = 3.4
    add_box(slide, MX + 0.6, line_y, CW - 1.2, 0.02, fill=th.c("line"), shape=MSO_SHAPE.RECTANGLE)
    span = CW - 1.2
    n = len(steps)
    for i, st in enumerate(steps):
        cx = MX + 0.6 + span * (i / (n - 1) if n > 1 else 0.5)
        add_box(slide, cx - 0.09, line_y - 0.09, 0.18, 0.18, fill=th.c("accent"), shape=MSO_SHAPE.OVAL)
        tf = add_text(slide, cx - 1.2, line_y - 0.75, 2.4, 0.4, anchor=MSO_ANCHOR.MIDDLE)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        set_run(p.add_run(), str(st.get("time", "")), th, size=14, bold=True, color=th.c("primary"))
        tf = add_text(slide, cx - 1.2, line_y + 0.28, 2.4, 0.5, anchor=MSO_ANCHOR.TOP)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        set_run(p.add_run(), str(st.get("title", "")), th, size=14, bold=True, color=th.c("text"))
        if st.get("detail"):
            tf = add_text(slide, cx - 1.2, line_y + 0.78, 2.4, 0.9, anchor=MSO_ANCHOR.TOP)
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            set_run(p.add_run(), str(st["detail"]), th, size=11, color=th.c("muted"))


def draw_conclusion(slide, sec, th):
    CW = th.W - 2 * MX
    draw_heading(slide, th, sec.get("heading", "小结"))
    tf = add_text(slide, MX, 1.5, CW, th.H - 1.8, anchor=MSO_ANCHOR.TOP)
    first = True
    for i, pt in enumerate(sec.get("points", []), 1):
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(12)
        set_run(p.add_run(), f"{i}. ", th, size=16, bold=True, color=th.c("accent"))
        set_run(p.add_run(), str(pt), th, size=16)


def draw_next_steps(slide, sec, th):
    CW = th.W - 2 * MX
    draw_heading(slide, th, sec.get("heading", "下一步计划"))
    y = 1.5
    ch, gap = 0.85, 0.18
    for it in sec.get("items", []):
        add_box(slide, MX, y, CW, ch, fill=th.c("light"), line=None)
        tf = add_text(slide, MX + 0.3, y, CW - 0.6, ch, anchor=MSO_ANCHOR.MIDDLE)
        p = tf.paragraphs[0]
        set_run(p.add_run(), str(it.get("task", "")), th, size=15, bold=True, color=th.c("text"))
        meta = " · ".join(
            x for x in [
                f"目标：{it['target']}" if it.get("target") else None,
                f"截止：{it['deadline']}" if it.get("deadline") else None,
            ] if x
        )
        if meta:
            set_run(p.add_run(), "    " + meta, th, size=11, color=th.c("muted"))
        y += ch + gap


DISPATCH = {
    "cover": draw_cover,
    "agenda": draw_agenda,
    "section_header": draw_section_header,
    "points": draw_points,
    "figure": draw_figure,
    "metrics": draw_metrics,
    "table": draw_table,
    "timeline": draw_timeline,
    "conclusion": draw_conclusion,
    "next_steps": draw_next_steps,
}


def draw(slide, sec, th, base_dir):
    t = sec.get("type")
    fn = DISPATCH.get(t)
    if fn is None:
        raise ValueError(f"未知版式类型：{t!r}")
    if t == "figure":
        fn(slide, sec, th, base_dir)
    else:
        fn(slide, sec, th)


# ---------------- 校验 ----------------

def validate(brief, base_dir):
    meta = brief.get("meta") or {}
    if not (meta.get("title") or "").strip():
        raise ValueError("meta.title 必填且非空")
    secs = brief.get("sections")
    if not isinstance(secs, list) or not secs:
        raise ValueError("sections 必填且非空列表")
    for i, sec in enumerate(secs, 1):
        t = sec.get("type")
        if t not in ALLOWED_TYPES:
            raise ValueError(f"第 {i} 段：未知类型 {t!r}")
        for k in REQUIRED.get(t, []):
            v = sec.get(k)
            if v is None or (isinstance(v, (str, list)) and len(v) == 0):
                raise ValueError(f"第 {i} 段（{t}）：缺少字段 {k}")
        if t == "figure":
            p = image_path(sec["image"], base_dir)
            if not os.path.isfile(p):
                raise ValueError(f"第 {i} 段（figure）：图片不存在 {sec['image']!r}")
        if t == "table":
            ncols = len(sec["columns"])
            for r_i, row in enumerate(sec["rows"], 1):
                if len(row) != ncols:
                    raise ValueError(
                        f"第 {i} 段（table）：第 {r_i} 行长度 {len(row)} 与列数 {ncols} 不一致"
                    )


# ---------------- 单页替换 ----------------

def replace_slide(prs, idx, layout):
    """原地清空第 idx 页并返回该 slide（保留其在 pptx 中的位置与 part），避免增删 part。"""
    slides = list(prs.slides)
    if 0 <= idx < len(slides):
        slide = slides[idx]
        for shape in list(slide.shapes):
            shape._element.getparent().remove(shape._element)
        return slide
    return prs.slides.add_slide(layout)


# ---------------- 轻量预览 ----------------

def render_preview(brief, th, out_png):
    import numpy as np
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False

    secs = brief["sections"]
    n = len(secs)
    cols = 4
    rows = math.ceil(n / cols)
    fig, axes = plt.subplots(rows, cols, figsize=(cols * 3.2, rows * 1.9), dpi=130)
    axs = np.array(axes).reshape(-1)

    def one(ax, sec):
        ax.set_xlim(0, 16)
        ax.set_ylim(0, 9)
        ax.axis("off")
        ax.add_patch(plt.Rectangle((0, 0), 16, 9, fill=True, facecolor="white", edgecolor="#C9C9C9"))
        t = sec.get("type", "")
        title = sec.get("heading") or sec.get("title") or t
        ax.text(0.8, 8.3, f"[{t}] {title}", fontsize=11, color="#1F3864",
                fontweight="bold", va="top")
        lines = []
        for k, pre in (("points", "• "), ("items", "- ")):
            for x in sec.get(k, [])[:5]:
                if isinstance(x, dict):
                    s = str(x.get("task") or x.get("title") or x.get("label") or x)
                    if x.get("deadline"):
                        s += f"（截止 {x['deadline']}）"
                    lines.append(pre + s)
                else:
                    lines.append(pre + str(x))
        if t == "figure":
            lines.append("[图: " + str(sec.get("image", "")) + "]")
            for tk in sec.get("takeaways", [])[:3]:
                lines.append("• " + str(tk))
            for m in sec.get("metrics", [])[:3]:
                lines.append(f"{m.get('label', '')} = {m.get('value', '')}")
        if t == "metrics":
            for m in sec.get("metrics", [])[:5]:
                lines.append(f"{m.get('label', '')} = {m.get('value', '')}")
        if t == "timeline":
            for s in sec.get("steps", [])[:5]:
                lines.append(f"{s.get('time', '')} {s.get('title', '')}")
        if t == "table":
            lines.append("列: " + str(sec.get("columns", [])))
        y = 7.2
        for ln in lines:
            ax.text(0.8, y, str(ln)[:28], fontsize=9, color="#333333", va="top")
            y -= 0.55
        if t == "figure":
            ax.add_patch(plt.Rectangle((0.8, 1.0), 7.6, 5.5, fill=True,
                                       facecolor="#D6E4F0", edgecolor="#2E74B5"))
            ax.text(4.6, 3.7, "image", ha="center", va="center", fontsize=9, color="#2E74B5")

    for i, ax in enumerate(axs):
        if i < n:
            one(ax, secs[i])
        else:
            ax.axis("off")
    fig.suptitle("组会 PPT 轻量预览（内容结构，非逐像素还原）", fontsize=13, color="#1F3864")
    plt.tight_layout(rect=(0, 0, 1, 0.97))
    plt.savefig(out_png, dpi=130, facecolor="white")
    plt.close(fig)


def write_snapshot(brief_path, brief, th, history_dir, template_path,
                   image_dependencies, source_info, output=None,
                   generation_mode="full", slide=None):
    """写可移植 brief + 模板 + 去重图片 + 预览，成功后发布完整版本。"""
    if generation_mode not in {"full", "slide"}:
        raise ValueError(f"未知历史生成模式：{generation_mode!r}")
    if generation_mode == "full" and not output:
        raise ValueError("全量历史快照需要输出路径")

    # 先完成 PPT，再写历史；任何源文件在此期间变化都不能记为一致快照。
    for label, path, expected in [
        ("brief", brief_path, source_info["brief_sha256"]),
        ("模板", template_path, source_info["template_sha256"]),
        ("生成器", __file__, source_info["generator_sha256"]),
    ]:
        if file_hash(path) != expected:
            raise ValueError(f"{label}在生成过程中发生变化，请重试；未写入历史快照")

    history_dir = os.path.abspath(history_dir)
    versions_dir = os.path.join(history_dir, "versions")
    previews_dir = os.path.join(history_dir, "previews")
    assets_dir = os.path.join(history_dir, "assets")
    os.makedirs(assets_dir, exist_ok=True)
    os.makedirs(versions_dir, exist_ok=True)
    os.makedirs(previews_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    base_id = f"{timestamp}-{source_info['brief_sha256'][:8]}"
    version_id = base_id
    serial = 1
    while (os.path.exists(os.path.join(versions_dir, version_id))
           or os.path.exists(os.path.join(previews_dir, version_id))):
        version_id = f"{base_id}-{serial}"
        serial += 1

    staging = tempfile.mkdtemp(prefix=".pending-", dir=history_dir)
    stage_version = os.path.join(staging, "version")
    stage_preview = os.path.join(staging, "preview")
    os.makedirs(stage_version)
    os.makedirs(stage_preview)
    final_version = os.path.join(versions_dir, version_id)
    final_preview = os.path.join(previews_dir, version_id)
    published = []

    try:
        snapshot_brief = copy.deepcopy(brief)
        for index, dep in enumerate(image_dependencies):
            if file_hash(dep["source"]) != dep["sha256"]:
                raise ValueError(
                    f"第 {dep['section']} 段图片在生成过程中发生变化，请重试；未写入历史快照")
            asset_path = os.path.join(assets_dir, dep["sha256"])
            if os.path.exists(asset_path):
                if not os.path.isfile(asset_path) or file_hash(asset_path) != dep["sha256"]:
                    raise ValueError(f"历史图片资产损坏或与哈希不符：{asset_path}")
            else:
                staged_asset = os.path.join(staging, f"asset-{index}")
                shutil.copyfile(dep["source"], staged_asset)
                if file_hash(staged_asset) != dep["sha256"]:
                    raise ValueError("图片复制后哈希不一致，未写入历史快照")
                # 同哈希即同字节内容；另一进程已写入时无需重复保存。
                if os.path.exists(asset_path):
                    if file_hash(asset_path) != dep["sha256"]:
                        raise ValueError(f"历史图片资产损坏或与哈希不符：{asset_path}")
                else:
                    os.replace(staged_asset, asset_path)
            snapshot_brief["sections"][dep["section"] - 1]["image"] = (
                f"../../assets/{dep['sha256']}")

        snapshot_brief_path = os.path.join(stage_version, "brief.yaml")
        with open(snapshot_brief_path, "w", encoding="utf-8") as f:
            yaml.safe_dump(snapshot_brief, f, allow_unicode=True, sort_keys=False)
        validate(snapshot_brief, stage_version)
        snapshot_brief_hash = file_hash(snapshot_brief_path)

        shutil.copyfile(template_path, os.path.join(stage_version, "template.yaml"))
        if file_hash(os.path.join(stage_version, "template.yaml")) != source_info["template_sha256"]:
            raise ValueError("模板复制后哈希不一致，未写入历史快照")

        manifest = {
            "skill": SNAPSHOT_SKILL,
            "format_version": SNAPSHOT_FORMAT_VERSION,
            "id": version_id,
            "created_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "generation_mode": generation_mode,
            "slide": slide,
            "brief_sha256": source_info["brief_sha256"],
            "snapshot_brief_sha256": snapshot_brief_hash,
            "template": {
                "name": th.name,
                "sha256": source_info["template_sha256"],
            },
            "generator_sha256": source_info["generator_sha256"],
            "images": source_info["images"],
        }
        write_json_atomic(os.path.join(stage_version, "manifest.json"), manifest)
        render_preview(snapshot_brief, th, os.path.join(stage_preview, "preview.png"))

        # 只有所有资源就绪后才公开版本；失败则把已发布目录移回本次临时目录。
        for source, destination in [
            (stage_version, final_version),
            (stage_preview, final_preview),
        ]:
            os.replace(source, destination)
            published.append((source, destination))
        if generation_mode == "full":
            write_json_atomic(
                os.path.join(history_dir, ".last_generated.json"),
                generation_marker(source_info, output),
            )
        return os.path.join(final_version, "brief.yaml")
    except Exception:
        for source, destination in reversed(published):
            if os.path.isdir(destination) and not os.path.exists(source):
                try:
                    os.replace(destination, source)
                except OSError:
                    pass
        raise
    finally:
        # staging 是本函数刚创建的专用临时目录，不含用户原历史。
        shutil.rmtree(staging, ignore_errors=True)


# ---------------- 主流程 ----------------

def build(brief, th, args, base_dir):
    sections = brief["sections"]

    if args.slide is not None and os.path.isfile(args.output):
        prs = Presentation(args.output)
        layout = prs.slide_layouts[6]
        idx = args.slide - 1
        if not (0 <= idx < len(sections)):
            raise ValueError(f"--slide {args.slide} 超出范围（1–{len(sections)}）")
        slide = replace_slide(prs, idx, layout)
        draw(slide, sections[idx], th, base_dir)
        return prs

    prs = Presentation()
    prs.slide_width = Inches(th.W)
    prs.slide_height = Inches(th.H)
    layout = prs.slide_layouts[6]
    if args.slide is not None:
        print("警告：--slide 需要已存在的输出文件；改为全量生成", file=sys.stderr)
    for sec in sections:
        slide = prs.slides.add_slide(layout)
        draw(slide, sec, th, base_dir)
    return prs


def main(argv):
    ap = argparse.ArgumentParser(description="组会 PPT 生成脚本")
    ap.add_argument("input", help="brief.yaml 路径")
    ap.add_argument("-o", "--output", required=True, help="输出 .pptx 路径")
    ap.add_argument("--template", default=None,
                    help="模板名；新快照未指定时使用其归档模板，普通 brief 默认 default")
    ap.add_argument("--slide", type=int, default=None, help="只重建第 N 页并原地替换")
    ap.add_argument("--history-dir", default=None, help="同时写 brief 快照 + 轻量预览的目录")
    ap.add_argument("--template-dir", default=None, help="模板目录（默认脚本上级的 templates/）")
    args = ap.parse_args(argv)

    brief_path = os.path.abspath(args.input)
    if not os.path.isfile(brief_path):
        print(f"错误：找不到 brief 文件 {brief_path}", file=sys.stderr)
        return 2
    try:
        brief, loaded_brief_hash = load_yaml_with_hash(brief_path)
    except (OSError, UnicodeError, yaml.YAMLError) as e:
        print(f"错误：无法读取 brief：{e}", file=sys.stderr)
        return 2
    base_dir = os.path.dirname(brief_path)

    try:
        validate(brief, base_dir)
    except ValueError as e:
        print(f"校验失败：{e}", file=sys.stderr)
        return 2

    try:
        tpl_path = resolve_template_path(brief_path, args.template, args.template_dir)
    except (OSError, ValueError) as e:
        print(f"错误：{e}", file=sys.stderr)
        return 2
    if not os.path.isfile(tpl_path):
        print(f"错误：找不到模板档案 {tpl_path}", file=sys.stderr)
        return 2
    try:
        template_data, loaded_template_hash = load_yaml_with_hash(tpl_path)
        th = Theme(template_data)
        validate_snapshot_inputs(brief_path, brief, tpl_path)
    except (OSError, UnicodeError, yaml.YAMLError, ValueError) as e:
        print(f"错误：{e}", file=sys.stderr)
        return 2
    out_abs = os.path.abspath(args.output)

    source_info = None
    image_dependencies = []
    try:
        if args.history_dir:
            image_dependencies = collect_image_dependencies(brief, base_dir)
            source_info = describe_inputs(
                brief_path, tpl_path, image_dependencies,
                brief_hash=loaded_brief_hash, template_hash=loaded_template_hash)
            if (args.slide is None
                    and should_skip_generation(args.history_dir, out_abs, source_info)):
                print(f"无变化，跳过重建（{args.output}）")
                return 0
        prs = build(brief, th, args, base_dir)
    except (OSError, ValueError) as e:
        print(f"错误：{e}", file=sys.stderr)
        return 2

    out_dir = os.path.dirname(out_abs) or "."
    os.makedirs(out_dir, exist_ok=True)
    prs.save(out_abs)
    print(f"共 {len(prs.slides)} 页 · 模板 {th.name} → {out_abs}")

    if args.history_dir:
        try:
            if args.slide is not None:
                mark_partial_generation(args.history_dir, out_abs, source_info, args.slide)
            else:
                mark_history_pending(args.history_dir, out_abs, source_info)
            snap = write_snapshot(
                brief_path, brief, th, args.history_dir, tpl_path,
                image_dependencies, source_info, output=out_abs,
                generation_mode="slide" if args.slide is not None else "full",
                slide=args.slide,
            )
        except Exception as e:
            print(f"错误：PPT 已保存，但历史快照写入失败：{e}", file=sys.stderr)
            return 1
        print(f"已写快照：{snap}")

    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
