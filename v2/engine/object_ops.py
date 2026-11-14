#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""v2 对象级修改适配层（M1 原型）。

设计约束（规格 §3）：
- 只按身份定位对象，绝不按标题文字、形状名称或页码匹配；
- 只改动明确指向的对象，其余部件保持字节不变；
- 图片替换建立新媒体引用，不覆盖可能被共享的媒体字节；
- 能力不足时返回明确原因，而不是静默降级。
"""
import os
import sys

from lxml import etree
from pptx.util import Emu

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

MARKER_PREFIX = "gmppt:"


class CapabilityError(RuntimeError):
    """能力不足或不支持的操作，附带明确原因。"""


def shape_marker(shape):
    try:
        descr = shape._element._nvXxPr.cNvPr.get("descr") or ""
    except Exception:
        return None
    return descr if descr.startswith(MARKER_PREFIX) else None


def shape_tags(shape):
    try:
        nv_pr = shape._element._nvXxPr.nvPr
        ref = nv_pr.find(".//{%s}custDataLst/{%s}tags" % (P, P))
        if ref is None:
            return {}
        part = shape.part.related_part(ref.get("{%s}id" % R))
        root = etree.fromstring(part.blob)
        return {tag.get("name"): tag.get("val") for tag in root.findall("{%s}tag" % P)}
    except Exception:
        return {}


def find_objects(prs, identity):
    """按身份标记或 tag 值定位对象；返回 [(slide, shape)]。"""
    hits = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape_marker(shape) == identity:
                hits.append((slide, shape, "marker"))
                continue
            tags = shape_tags(shape)
            if identity in tags.values() or identity in tags:
                hits.append((slide, shape, "tags"))
    return hits


def require_one(prs, identity):
    hits = find_objects(prs, identity)
    if not hits:
        raise CapabilityError("找不到身份为 %r 的对象" % identity)
    if len(hits) > 1:
        raise CapabilityError("身份 %r 命中 %d 个对象（疑似复制页导致重复身份），停止写入"
                              % (identity, len(hits)))
    return hits[0]


def set_text(prs, identity, text):
    """整段替换文本，保留首个 run 的格式；多段落时报告能力边界。"""
    slide, shape, _ = require_one(prs, identity)
    if not getattr(shape, "has_text_frame", False) or not shape.has_text_frame:
        raise CapabilityError("对象 %r 不是文本框，无法替换文本" % identity)
    frame = shape.text_frame
    paragraphs = frame.paragraphs
    if len(paragraphs) > 1:
        raise CapabilityError("对象 %r 含 %d 个段落，首版只支持整段单段落改写"
                              % (identity, len(paragraphs)))
    paragraph = paragraphs[0]
    runs = paragraph.runs
    if not runs:
        raise CapabilityError("对象 %r 没有可写入的文本 run" % identity)
    runs[0].text = text
    for extra in runs[1:]:
        extra._r.getparent().remove(extra._r)
    return {"op": "set_text", "identity": identity, "text": text, "slide_id": slide.slide_id}


def replace_picture(prs, identity, image, keep_frame=True):
    """替换图片内容；建立新媒体引用，不触碰共享媒体字节。"""
    slide, shape, _ = require_one(prs, identity)
    if shape.shape_type != 13:  # PICTURE
        raise CapabilityError("对象 %r 不是图片（type=%s）" % (identity, shape.shape_type))
    if not os.path.isfile(image):
        raise CapabilityError("图片文件不存在：%s" % image)
    blip = shape._element.blipFill.blip
    old_rid = blip.get("{%s}embed" % R)
    old_part = slide.part.related_part(old_rid)
    old_size = (shape.left, shape.top, shape.width, shape.height)
    image_part, new_rid = slide.part.get_or_add_image_part(image)
    blip.set("{%s}embed" % R, new_rid)
    if keep_frame:
        shape.left, shape.top, shape.width, shape.height = old_size
    return {"op": "replace_picture", "identity": identity, "image": image,
            "slide_id": slide.slide_id, "new_media": str(image_part.partname),
            "shared_media_untouched": old_part.partname != image_part.partname}


def set_geometry(prs, identity, left=None, top=None, width=None, height=None,
                 keep_aspect=False, unit=914400):
    """更新几何属性；keep_aspect 只在给单一尺寸时按原比例推算。"""
    slide, shape, _ = require_one(prs, identity)
    box = [shape.left, shape.top, shape.width, shape.height]
    ratio = shape.width / shape.height if shape.height else 1.0
    if left is not None:
        box[0] = int(left * unit)
    if top is not None:
        box[1] = int(top * unit)
    if width is not None:
        box[2] = int(width * unit)
        if keep_aspect:
            box[3] = int(box[2] / ratio)
    if height is not None:
        box[3] = int(height * unit)
        if keep_aspect:
            box[2] = int(box[3] * ratio)
    shape.left, shape.top, shape.width, shape.height = box
    return {"op": "set_geometry", "identity": identity, "slide_id": slide.slide_id,
            "frame": [int(v) for v in box]}


def reorder_slides(prs, slide_ids):
    """按给定 slide_id 顺序重排页面；缺项或多项时拒绝执行。"""
    sld_ids = prs.slides._sldIdLst
    current = [int(el.get("id")) for el in sld_ids]
    if sorted(current) != sorted(slide_ids):
        raise CapabilityError("页序参数与现有页面不一致：%s vs %s" % (sorted(slide_ids), sorted(current)))
    lookup = {int(el.get("id")): el for el in sld_ids}
    for el in list(sld_ids):
        sld_ids.remove(el)
    for value in slide_ids:
        sld_ids.append(lookup[value])
    return {"op": "reorder_slides", "order": list(slide_ids)}


OPS = {"set_text": set_text, "replace_picture": replace_picture, "set_geometry": set_geometry}


def apply_plan(prs, plan):
    """按计划顺序执行；任一步失败即抛出，由调用方决定回滚策略。"""
    done = []
    for step in plan:
        op = step.get("op")
        if op == "reorder_slides":
            done.append(reorder_slides(prs, step["order"]))
        elif op in OPS:
            done.append(OPS[op](prs, step["identity"], **{k: v for k, v in step.items()
                                                           if k not in ("op", "identity")}))
        else:
            raise CapabilityError("未知操作 %r" % op)
    return done
