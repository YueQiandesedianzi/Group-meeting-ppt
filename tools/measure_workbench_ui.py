#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""采集工作台界面的客观易用性指标，输出 JSON。

用法：在仓库根目录执行
    python tools/measure_workbench_ui.py

指标口径见 docs/workbench-usability.md。脚本只读界面与源码，不写用户状态文件。
"""
import json
import os
import re
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "scripts"))

import tkinter as tk  # noqa: E402
from tkinter import font as tkfont  # noqa: E402

import yaml  # noqa: E402

import workbench as wb  # noqa: E402
import workbench_logic as wl  # noqa: E402

SAMPLE_PATH = r"C:\Users\HP\Documents\科研\组会\2026 秋季学期\第 12 周进展汇报\brief.yaml"


def count_buttons(frame):
    total = 0
    stack = [frame]
    while stack:
        widget = stack.pop()
        for child in widget.winfo_children():
            if child.winfo_class() in ("TButton", "Button"):
                total += 1
            stack.append(child)
    return total


def main():
    wb.STATE_PATH = os.path.join(tempfile.mkdtemp(prefix="wb-measure-"), "state.json")
    app = wb.Workbench()
    app.withdraw()
    app.update()
    app.load_brief_into_editor(os.path.join(ROOT, "assets", "demo_brief.yaml"))
    app.update()
    app.refresh_outline()
    app.refresh_slide_choices()
    app.update()

    brief = yaml.safe_load(open(os.path.join(ROOT, "assets", "demo_brief.yaml"), encoding="utf-8"))
    rows = wl.section_outline(brief)
    font = tkfont.nametofont("TkDefaultFont")

    title_col = app.outline_tree.column("title", "width")
    longest_title = max((r["title"] for r in rows), key=len)
    values = list(app.slide_combo["values"])
    longest_value = max(values, key=len)
    combo_width = app.slide_combo.winfo_reqwidth()

    source = open(os.path.join(ROOT, "scripts", "workbench.py"), encoding="utf-8").read()
    report = {
        "geometry": app.geometry().split("+")[0],
        "requested_size": "%dx%d" % (app.winfo_reqwidth(), app.winfo_reqheight()),
        "tabs": [app.nb.tab(t, "text") for t in app.nb.tabs()],
        "outline_title_col_px": title_col,
        "longest_title": longest_title,
        "longest_title_px": font.measure(longest_title),
        "title_fits": font.measure(longest_title) <= title_col - 10,
        "slide_combo_reqwidth_px": combo_width,
        "longest_slide_value": longest_value,
        "longest_slide_value_px": font.measure(longest_value),
        "slide_fits": font.measure(longest_value) <= combo_width - 24,
        "left_rail_width_px": 240,
        "longest_path_sample_px": font.measure(SAMPLE_PATH),
        "path_fits": font.measure(SAMPLE_PATH) <= 240 - 12,
        "shortcuts": [s for s in ("<Control-s>", "<Control-o>", "<Control-n>",
                                  "<Control-Return>", "<F5>") if app.bind_all(s)],
        "messagebox_calls": len(re.findall(r"messagebox\.\w+\(", source)),
        "askyesno": len(re.findall(r"askyesno\(", source)),
        "askyesnocancel": len(re.findall(r"askyesnocancel\(", source)),
        "showerror": len(re.findall(r"showerror\(", source)),
        "showwarning": len(re.findall(r"showwarning\(", source)),
        "showinfo": len(re.findall(r"showinfo\(", source)),
        "warnings_for_demo": wl.compute_warnings(brief),
        "demo_pages": wl.estimate_pages(brief),
        "layout_types": len(wl.SECTION_TYPES),
        "buttons": {},
    }
    for index, name in enumerate(report["tabs"]):
        app.nb.select(index)
        app.update()
        report["buttons"][name] = count_buttons(app.nb.nametowidget(app.nb.select()))

    app.destroy()
    tk._default_root = None
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
