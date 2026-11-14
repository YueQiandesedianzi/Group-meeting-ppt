#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""工作台的可测逻辑：容量提醒、页数估算、最近文件显示名。

这里只放纯函数，不导入 Tkinter、也不导入渲染引擎，便于单元测试。
硬性错误（缺字段、图片不存在、表格行列不一致、目录/时间线/下一步超过 12 项）由
generate_ppt.validate 负责；本模块只产出“不阻塞生成但建议处理”的提醒。
"""
import os

MAX_POINTS = 6
MAX_TITLE = 20
MAX_TAKEAWAYS = 3
MAX_SECTIONS = 24
LONG_POINT = 40


def display_labels(paths):
    """最近文件在左栏的显示名：默认只给文件名，重名时补上级目录。"""
    names = [os.path.basename(p) for p in paths]
    seen = {}
    for n in names:
        seen[n] = seen.get(n, 0) + 1
    labels = []
    for path, name in zip(paths, names):
        if seen[name] > 1:
            parent = os.path.basename(os.path.dirname(os.path.normpath(path)))
            labels.append(f"{name}（{parent}）" if parent else name)
        else:
            labels.append(name)
    return labels


def estimate_pages(brief):
    """brief 里的 sections 数量，也就是预计页数。"""
    secs = brief.get("sections") if isinstance(brief, dict) else None
    return len(secs) if isinstance(secs, list) else 0


def _title_of(sec):
    for key in ("heading", "title"):
        value = sec.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()
    return ""


def compute_warnings(brief):
    """返回建议性提醒列表（有序、可直接展示）。"""
    warnings = []
    secs = brief.get("sections") if isinstance(brief, dict) else None
    if not isinstance(secs, list):
        return warnings
    for i, sec in enumerate(secs, 1):
        if not isinstance(sec, dict):
            continue
        title = _title_of(sec)
        if len(title) > MAX_TITLE:
            warnings.append(f"第 {i} 段标题 {len(title)} 字，建议不超过 {MAX_TITLE} 字")
        typ = sec.get("type")
        if typ == "points":
            pts = sec.get("points")
            if isinstance(pts, list):
                if len(pts) > MAX_POINTS:
                    warnings.append(f"第 {i} 段 points 有 {len(pts)} 条，建议不超过 {MAX_POINTS} 条并拆页")
                long_items = [p for p in pts if isinstance(p, str) and len(p.strip()) > LONG_POINT]
                if long_items:
                    warnings.append(
                        f"第 {i} 段有 {len(long_items)} 条要点超过 {LONG_POINT} 字，可能在文本框内溢出")
        if typ == "figure":
            takeaways = sec.get("takeaways")
            if isinstance(takeaways, list) and len(takeaways) > MAX_TAKEAWAYS:
                warnings.append(
                    f"第 {i} 段 takeaways 有 {len(takeaways)} 条，建议不超过 {MAX_TAKEAWAYS} 条")
    if len(secs) > MAX_SECTIONS:
        warnings.append(f"共 {len(secs)} 页，篇幅较长，建议确认是否需要精简")
    return warnings


def format_preflight(pages, warnings):
    """生成前检查的对话框文案。"""
    lines = [f"预计 {pages} 页。"]
    if warnings:
        lines.append(f"发现 {len(warnings)} 条提醒：")
        lines.extend("• " + w for w in warnings)
        lines.append("")
        lines.append("这些提醒不会阻止生成，是否继续？")
    else:
        lines.append("没有发现需要提醒的问题。")
    return "\n".join(lines)
