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


# ---------------- 大纲与版式骨架 ----------------

SECTION_TYPES = ("cover", "agenda", "section_header", "points", "figure",
                 "metrics", "table", "timeline", "conclusion", "next_steps")

TYPE_LABELS = {
    "cover": "封面",
    "agenda": "目录",
    "section_header": "章节页",
    "points": "要点",
    "figure": "图片",
    "metrics": "指标",
    "table": "表格",
    "timeline": "时间线",
    "conclusion": "结论",
    "next_steps": "下一步",
}

SNIPPETS = {
    "cover": """  - type: cover
    title: 汇报标题
    subtitle: 副标题（可选）
    presenter: 汇报人
    date: 2026-11""",
    "agenda": """  - type: agenda
    heading: 汇报提纲
    items:
      - 第一项
      - 第二项""",
    "section_header": """  - type: section_header
    heading: 章节标题
    subtitle: 一句话说明（可选）""",
    "points": """  - type: points
    heading: 要点标题
    points:
      - 第一点
      - 第二点""",
    "figure": """  - type: figure
    heading: 结果图
    image: 图片文件名.png
    caption: 图注（可选）
    takeaways:
      - 结论要点""",
    "metrics": """  - type: metrics
    heading: 主要指标
    metrics:
      - { label: 指标名, value: 0.0, note: 说明（可选） }""",
    "table": """  - type: table
    heading: 对照表
    columns: [列 1, 列 2]
    rows:
      - [值 1, 值 2]
      - [值 3, 值 4]""",
    "timeline": """  - type: timeline
    heading: 研究路线
    steps:
      - { time: 阶段 1, title: 任务名, detail: 说明（可选） }""",
    "conclusion": """  - type: conclusion
    heading: 小结
    points:
      - 结论一
      - 结论二""",
    "next_steps": """  - type: next_steps
    heading: 下一步计划
    items:
      - { task: 任务名, target: 目标（可选）, deadline: 截止（可选） }""",
}

_SECTION_START = None  # 延迟编译，保持模块导入轻量


def _section_start_pattern():
    global _SECTION_START
    if _SECTION_START is None:
        import re
        _SECTION_START = re.compile(r"^(\s*)-\s")
    return _SECTION_START


def type_label(typ):
    return TYPE_LABELS.get(typ, typ or "未知")


def section_outline(brief):
    """大纲行：每段一条 {index, type, label, title}。"""
    rows = []
    secs = brief.get("sections") if isinstance(brief, dict) else None
    if not isinstance(secs, list):
        return rows
    for i, sec in enumerate(secs, 1):
        if not isinstance(sec, dict):
            rows.append({"index": i, "type": "?", "label": "未知", "title": ""})
            continue
        typ = sec.get("type") if isinstance(sec.get("type"), str) else "?"
        rows.append({"index": i, "type": typ, "label": type_label(typ), "title": _title_of(sec)})
    return rows


def section_spans(text):
    """返回各段在文本中的行区间 [(start, end)]（0 基、半开），失败返回 []。"""
    parts = _split_sections(text)
    if parts is None:
        return []
    head, blocks, tail = parts
    spans = []
    pos = len(head)
    for i, block in enumerate(blocks):
        if i:
            pos += 1  # 重建时插入的空行
        spans.append((pos, pos + len(block)))
        pos += len(block)
    return spans


def _split_sections(text):
    """把文本切成 (sections: 之前及本身的头部, 每段的行列表, 其余尾部)。

    找不到 sections: 返回 None；sections 为空列表时返回块列表 []。
    """
    lines = text.split("\n")
    key = None
    for i, line in enumerate(lines):
        if line.startswith("sections:") or line.startswith("sections :"):
            key = i
            break
    if key is None:
        return None
    pattern = _section_start_pattern()
    indent = None
    spans = []
    for i in range(key + 1, len(lines)):
        line = lines[i]
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        match = pattern.match(line)
        if match:
            if indent is None:
                indent = len(match.group(1))
            if len(match.group(1)) == indent:
                if spans:
                    spans[-1][1] = i
                spans.append([i, len(lines)])
            continue
        if indent is not None and not line[0].isspace():
            if spans:
                spans[-1][1] = i
            break
    if not spans:
        return lines[:key + 1], [], lines[key + 1:]
    for span in spans:
        while span[1] > span[0] + 1 and not lines[span[1] - 1].strip():
            span[1] -= 1
    return lines[:spans[0][0]], [lines[s:e] for s, e in spans], lines[spans[-1][1]:]


def _rebuild(head, blocks, tail):
    body = []
    for i, block in enumerate(blocks):
        if i:
            body.append("")
        body.extend(block)
    return "\n".join(head + body + tail)


def move_section_text(text, index, delta):
    """把第 index 段（0 基）上移/下移 delta 位，返回新文本；越界或无法解析时原样返回。"""
    parts = _split_sections(text)
    if parts is None:
        return text
    head, blocks, tail = parts
    if not (0 <= index < len(blocks)):
        return text
    target = max(0, min(len(blocks) - 1, index + delta))
    if target == index:
        return text
    block = blocks.pop(index)
    blocks.insert(target, block)
    return _rebuild(head, blocks, tail)


def remove_section_text(text, index):
    """删除第 index 段（0 基），返回新文本。"""
    parts = _split_sections(text)
    if parts is None:
        return text
    head, blocks, tail = parts
    if not (0 <= index < len(blocks)):
        return text
    blocks.pop(index)
    return _rebuild(head, blocks, tail)


def append_snippet(text, typ):
    """在 sections 末尾追加一段版式骨架，返回新文本。"""
    if typ not in SNIPPETS:
        raise ValueError(f"未知版式 {typ!r}")
    parts = _split_sections(text)
    if parts is None:
        raise ValueError("brief 里没有 sections: 列表，无法追加版式")
    head, blocks, tail = parts
    blocks.append(SNIPPETS[typ].split("\n"))
    return _rebuild(head, blocks, tail)
