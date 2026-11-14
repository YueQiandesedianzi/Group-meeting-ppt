#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""组会 PPT 工作台（Tkinter）：brief 编辑 + 模板 + 生成 + 历史预览。

复用 generate_ppt.py 与 template_wizard.py 的引擎逻辑，不复制布局/校验/渲染代码。
"""
import os
import sys
import json
import subprocess
from types import SimpleNamespace

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

import generate_ppt as g  # noqa: E402
import template_wizard as tw  # noqa: E402
import workbench_logic as wl  # noqa: E402

TDIR = os.path.normpath(os.path.join(SCRIPT_DIR, "..", "templates"))
STATE_PATH = os.path.join(os.path.expanduser("~"), ".bohrium", "ppt-workbench.json")


# ---------------- 非界面工具 ----------------

def template_names(tdir=TDIR):
    if not os.path.isdir(tdir):
        return []
    return sorted(f[:-5] for f in os.listdir(tdir) if f.endswith(".yaml"))


def load_state():
    try:
        with open(STATE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def run_generate(brief_path, output, template_name=None, slide=None, history_dir=None, tdir=None):
    """调用引擎完成校验→渲染→保存→（可选）快照。返回 (页数, 模板名)。"""
    brief_path = os.path.abspath(brief_path)
    brief, loaded_brief_hash = g.load_yaml_with_hash(brief_path)
    base_dir = os.path.dirname(brief_path)
    g.validate(brief, base_dir)  # 失败抛 ValueError
    tpl_path = g.resolve_template_path(brief_path, template_name, tdir)
    if not os.path.isfile(tpl_path):
        raise ValueError(f"找不到模板档案 {tpl_path}")
    template_data, loaded_template_hash = g.load_yaml_with_hash(tpl_path)
    th = g.Theme(template_data)
    g.validate_snapshot_inputs(brief_path, brief, tpl_path)
    g.validate_output_path(output, brief_path, tpl_path, brief, base_dir)
    image_dependencies = g.collect_image_dependencies(brief, base_dir) if history_dir else []
    source_info = (
        g.describe_inputs(
            brief_path, tpl_path, image_dependencies,
            brief_hash=loaded_brief_hash, template_hash=loaded_template_hash)
        if history_dir else None)
    args = SimpleNamespace(slide=slide, output=output)
    prs = g.build(brief, th, args, base_dir)
    out_abs = os.path.abspath(output)
    os.makedirs(os.path.dirname(out_abs) or ".", exist_ok=True)
    prs.save(out_abs)
    if history_dir:
        if slide is not None:
            g.mark_partial_generation(history_dir, out_abs, source_info, slide)
        else:
            g.mark_history_pending(history_dir, out_abs, source_info)
        g.write_snapshot(
            brief_path, brief, th, history_dir, tpl_path,
            image_dependencies, source_info, output=out_abs,
            generation_mode="slide" if slide is not None else "full", slide=slide,
        )
    return len(prs.slides), th.name


def find_latest_preview(history_dir):
    p = os.path.join(history_dir, "previews")
    if not os.path.isdir(p):
        return None
    cands = [os.path.join(p, d, "preview.png") for d in os.listdir(p)
             if os.path.isfile(os.path.join(p, d, "preview.png"))]
    if not cands:
        return None
    cands.sort(key=os.path.getmtime, reverse=True)
    return cands[0]


def list_versions(history_dir):
    """同时列出新版本目录和旧的扁平 YAML，不让临时目录进入历史。"""
    vs = os.path.join(history_dir, "versions")
    out = []
    if not os.path.isdir(vs):
        return out
    for name in sorted(os.listdir(vs)):
        path = os.path.join(vs, name)
        is_legacy = os.path.isfile(path) and name.endswith(".yaml")
        is_snapshot = os.path.isdir(path)
        if not is_legacy and not is_snapshot:
            continue
        vid = name[:-5] if is_legacy else name
        snap = path if is_legacy else os.path.join(path, "brief.yaml")
        if not is_legacy and not os.path.isfile(snap) and not os.path.isfile(
                os.path.join(path, "manifest.json")):
            continue
        pp = os.path.join(history_dir, "previews", vid, "preview.png")
        version = {
            "id": vid,
            "yaml": snap,
            "preview": pp if os.path.isfile(pp) else None,
            "legacy": is_legacy,
            "template_name": None,
            "error": None,
        }
        if is_snapshot:
            try:
                manifest = g.read_snapshot_manifest(path)
                if manifest is None or not os.path.isfile(snap):
                    raise ValueError("版本缺少 manifest.json 或 brief.yaml")
                if not os.path.isfile(os.path.join(path, "template.yaml")):
                    raise ValueError("版本缺少 template.yaml")
                version["template_name"] = manifest.get("template", {}).get("name", "未命名")
            except (OSError, ValueError, AttributeError) as e:
                version["error"] = str(e)
        out.append(version)
    return out


def open_path(path):
    """跨平台打开文件/目录（Windows 用 startfile）。"""
    if os.name == "nt":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def reveal_path(path):
    """在系统文件管理器中定位文件。"""
    path = os.path.abspath(path)
    if os.name == "nt":
        subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
    elif sys.platform == "darwin":
        subprocess.Popen(["open", "-R", path])
    else:
        subprocess.Popen(["xdg-open", os.path.dirname(path)])


# ---------------- 工作台界面 ----------------

class Workbench(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("组会 PPT 工作台")
        self.geometry("1180x760")

        st = load_state()
        self._dirty = False
        self._loading = False
        self._generate_buttons = []
        geometry = st.get("geometry")
        if isinstance(geometry, str) and "x" in geometry:
            try:
                self.geometry(geometry)
            except tk.TclError:
                pass
        self.recent_briefs = [p for p in st.get("recent_briefs", []) if os.path.isfile(p)]
        self.brief_path = st.get("current_brief") or None
        self.template_name = st.get("template") or "default"
        if self.template_name not in template_names():
            self.template_name = "default"

        self.preview_img = None
        self.tpl_colors = dict(tw.DEFAULT_COLORS)
        self.tpl_fonts = dict(tw.DEFAULT_FONTS)
        self.tpl_slide = {"aspect": "16:9", "width_in": 13.333, "height_in": 7.5}
        self.tpl_name = tk.StringVar(value="")
        self.tpl_desc = tk.StringVar(value="")
        self.tpl_default = tk.BooleanVar(value=False)
        self.tpl_cjk = tk.StringVar(value=tw.DEFAULT_FONTS["cjk"])
        self.tpl_latin = tk.StringVar(value=tw.DEFAULT_FONTS["latin"])
        self.tpl_summary = tk.StringVar(value="（未导入）")

        self._build_toolbar()
        self._build_body()
        self._build_statusbar()

        self.refresh_left_briefs()
        self.refresh_left_templates()
        self.refresh_template_combo()
        if self.brief_path:
            self.load_brief_into_editor(self.brief_path)
        tab = st.get("tab")
        if isinstance(tab, int) and 0 <= tab < len(self.nb.tabs()):
            self.nb.select(tab)
        self._bind_shortcuts()
        self._set_title()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------- 布局 ----------
    def _build_toolbar(self):
        bar = ttk.Frame(self, padding=(8, 6))
        bar.pack(fill="x")
        ttk.Button(bar, text="新建 brief", command=self.new_brief).pack(side="left", padx=2)
        ttk.Button(bar, text="打开 brief…", command=self.open_brief).pack(side="left", padx=2)
        ttk.Button(bar, text="刷新", command=self.refresh_all).pack(side="left", padx=2)
        ttk.Button(bar, text="打开输出", command=self.open_output).pack(side="left", padx=2)
        ttk.Button(bar, text="打开所在文件夹", command=self.reveal_output).pack(side="left", padx=2)

    def _build_body(self):
        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=8, pady=4)

        # 左栏
        left = ttk.Frame(body, width=240)
        left.pack(side="left", fill="y", padx=(0, 8))
        left.pack_propagate(False)
        ttk.Label(left, text="项目（brief）").pack(anchor="w")
        self.brief_list = tk.Listbox(left, height=8)
        self.brief_list.pack(fill="x", pady=(2, 8))
        self.brief_list.bind("<Double-Button-1>", lambda e: self.open_selected_brief())
        self.brief_list.bind("<Motion>", self._on_brief_hover)
        self.brief_list.bind("<Leave>", lambda e: self.update_status())
        self.brief_list.bind("<Button-3>", self._show_brief_menu)
        self._brief_menu = tk.Menu(self, tearoff=0)
        self._brief_menu.add_command(label="打开", command=self.open_selected_brief)
        self._brief_menu.add_command(label="在资源管理器中显示", command=self.reveal_selected_brief)
        self._brief_menu.add_separator()
        self._brief_menu.add_command(label="从最近列表移除", command=self.remove_selected_brief)

        ttk.Label(left, text="模板（点击切换）").pack(anchor="w")
        self.tpl_list = tk.Listbox(left, height=10)
        self.tpl_list.pack(fill="both", expand=True, pady=(2, 0))
        self.tpl_list.bind("<<ListboxSelect>>", self.on_template_select)
        self.tpl_list.bind("<Double-Button-1>", self.on_template_select)

        # 主区 Notebook
        self.nb = ttk.Notebook(body)
        self.nb.pack(side="left", fill="both", expand=True)
        self._build_tab_brief()
        self._build_tab_outline()
        self._build_tab_generate()
        self._build_tab_history()
        self._build_tab_template()
        self.nb.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    def _build_statusbar(self):
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w", padding=(6, 2)).pack(fill="x", side="bottom")

    # ---------- Tab 1: brief 编辑 ----------
    def _build_tab_brief(self):
        f = ttk.Frame(self.nb, padding=8)
        self.nb.add(f, text="brief 编辑")
        row = ttk.Frame(f)
        row.pack(fill="x", pady=(0, 4))
        self.brief_path_var = tk.StringVar(value=self.brief_path or "")
        ttk.Entry(row, textvariable=self.brief_path_var).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="打开…", command=self.open_brief).pack(side="left", padx=2)
        ttk.Button(row, text="保存", command=self.save_brief).pack(side="left", padx=2)
        ttk.Button(row, text="另存为…", command=self.save_brief_as).pack(side="left", padx=2)

        tf = ttk.Frame(f)
        tf.pack(fill="both", expand=True)
        self.brief_text = tk.Text(tf, wrap="none", undo=True)
        ys = ttk.Scrollbar(tf, orient="vertical", command=self.brief_text.yview)
        xs = ttk.Scrollbar(tf, orient="horizontal", command=self.brief_text.xview)
        self.brief_text.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.brief_text.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        tf.rowconfigure(0, weight=1)
        tf.columnconfigure(0, weight=1)
        self.brief_text.bind("<<Modified>>", self._on_text_modified)

        row2 = ttk.Frame(f)
        row2.pack(fill="x", pady=(6, 0))
        ttk.Button(row2, text="校验", command=self.validate_brief).pack(side="left")
        self.validate_var = tk.StringVar(value="")
        ttk.Label(row2, textvariable=self.validate_var).pack(side="left", padx=8)

    # ---------- Tab 2: 大纲 ----------
    def _build_tab_outline(self):
        f = ttk.Frame(self.nb, padding=8)
        self.outline_tab = f
        self.nb.add(f, text="大纲")
        top = ttk.Frame(f)
        top.pack(fill="x", pady=(0, 4))
        ttk.Button(top, text="从文本刷新", command=self.refresh_outline).pack(side="left")
        ttk.Label(top, text="双击行可定位到 YAML 原文；右侧按钮直接改写 brief 文本").pack(side="left", padx=8)

        body = ttk.Frame(f)
        body.pack(fill="both", expand=True)

        self.outline_tree = ttk.Treeview(body, columns=("no", "type", "title"),
                                         show="headings", selectmode="browse")
        self.outline_tree.heading("no", text="#")
        self.outline_tree.heading("type", text="版式")
        self.outline_tree.heading("title", text="标题")
        self.outline_tree.column("no", width=44, anchor="center", stretch=False)
        self.outline_tree.column("type", width=96, anchor="center", stretch=False)
        self.outline_tree.column("title", width=360)
        self.outline_tree.bind("<Double-Button-1>", self._on_outline_double)
        self.outline_tree.pack(side="left", fill="both", expand=True)
        yscroll = ttk.Scrollbar(body, orient="vertical", command=self.outline_tree.yview)
        yscroll.pack(side="left", fill="y")
        self.outline_tree.configure(yscrollcommand=yscroll.set)

        right = ttk.Frame(body, width=150)
        right.pack(side="left", fill="y", padx=(10, 0))
        right.pack_propagate(False)
        ttk.Label(right, text="调整顺序").pack(anchor="w", pady=(0, 2))
        ttk.Button(right, text="上移", command=lambda: self.move_outline(-1)).pack(fill="x", pady=1)
        ttk.Button(right, text="下移", command=lambda: self.move_outline(1)).pack(fill="x", pady=1)
        ttk.Button(right, text="删除该段", command=self.remove_outline).pack(fill="x", pady=1)
        ttk.Label(right, text="插入版式（追加末尾）").pack(anchor="w", pady=(8, 2))
        for typ in wl.SECTION_TYPES:
            ttk.Button(right, text="+ " + wl.type_label(typ),
                       command=lambda t=typ: self.insert_outline(t)).pack(fill="x", pady=1)

    # ---------- Tab 3: 生成 ----------
    def _build_tab_generate(self):
        f = ttk.Frame(self.nb, padding=12)
        self.nb.add(f, text="生成")
        form = ttk.Frame(f)
        form.pack(fill="x")

        ttk.Label(form, text="模板").grid(row=0, column=0, sticky="w", pady=4)
        self.template_var = tk.StringVar(value=self.template_name)
        self.template_combo = ttk.Combobox(form, textvariable=self.template_var, state="readonly", width=28)
        self.template_combo.grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(form, text="输出").grid(row=1, column=0, sticky="w", pady=4)
        self.output_var = tk.StringVar(value=load_state().get("output", ""))
        ttk.Entry(form, textvariable=self.output_var, width=42).grid(row=1, column=1, sticky="w")
        ttk.Button(form, text="浏览…", command=self.browse_output).grid(row=1, column=2, padx=4)

        ttk.Label(form, text="历史目录").grid(row=2, column=0, sticky="w", pady=4)
        self.history_var = tk.StringVar(value=load_state().get("history_dir", ""))
        ttk.Entry(form, textvariable=self.history_var, width=42).grid(row=2, column=1, sticky="w")
        ttk.Button(form, text="浏览…", command=self.browse_history).grid(row=2, column=2, padx=4)

        ttk.Label(form, text="单页 N").grid(row=3, column=0, sticky="w", pady=4)
        self.slide_var = tk.StringVar(value="")
        ttk.Entry(form, textvariable=self.slide_var, width=8).grid(row=3, column=1, sticky="w")

        btns = ttk.Frame(f)
        btns.pack(fill="x", pady=8)
        full_btn = ttk.Button(btns, text="全量生成（Ctrl+Enter）", command=lambda: self.do_generate(None))
        full_btn.pack(side="left", padx=2)
        slide_btn = ttk.Button(btns, text="单页重建", command=self.do_generate_slide)
        slide_btn.pack(side="left", padx=2)
        self._generate_buttons = [full_btn, slide_btn]

        self.result_text = tk.Text(f, height=8, wrap="word", state="disabled")
        self.result_text.pack(fill="both", expand=True)

    # ---------- Tab 4: 预览与历史 ----------
    def _build_tab_history(self):
        f = ttk.Frame(self.nb, padding=8)
        self.nb.add(f, text="预览与历史")
        top = ttk.Frame(f)
        top.pack(fill="x", pady=(0, 4))
        ttk.Button(top, text="刷新列表", command=self.refresh_history).pack(side="left")
        ttk.Button(top, text="从该版本重建…", command=self.rebuild_version).pack(side="left", padx=6)

        paned = ttk.PanedWindow(f, orient="horizontal")
        paned.pack(fill="both", expand=True)
        self.version_list = tk.Listbox(paned, width=30)
        self.version_list.bind("<<ListboxSelect>>", lambda e: self.on_version_select())
        paned.add(self.version_list, weight=1)
        self.preview_label = ttk.Label(paned, text="（尚无预览）", anchor="center")
        paned.add(self.preview_label, weight=3)

    # ---------- Tab 5: 模板 ----------
    def _build_tab_template(self):
        f = ttk.Frame(self.nb, padding=12)
        self.nb.add(f, text="模板")
        top = ttk.Frame(f)
        top.pack(fill="x")
        ttk.Button(top, text="导入 .pptx（提取主题）…", command=self.import_pptx).pack(side="left")
        ttk.Label(top, textvariable=self.tpl_summary, wraplength=520).pack(side="left", padx=8)

        form = ttk.Frame(f)
        form.pack(fill="x", pady=10)
        ttk.Label(form, text="模板名").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=self.tpl_name, width=34).grid(row=0, column=1, sticky="w")
        ttk.Label(form, text="描述").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=self.tpl_desc, width=34).grid(row=1, column=1, sticky="w")

        ttk.Label(form, text="配色（点击色块修改）").grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 2))
        cf = ttk.Frame(form)
        cf.grid(row=3, column=0, columnspan=2, sticky="w")
        self.tpl_swatches = {}
        for i, k in enumerate(tw.COLOR_KEYS):
            ttk.Label(cf, text=tw.COLOR_LABELS[k]).grid(row=0, column=i, padx=(0, 2))
            b = tk.Button(cf, text=self.tpl_colors[k], bg=self.tpl_colors[k],
                          fg=tw.Wizard._fg(self.tpl_colors[k]), width=9,
                          command=lambda key=k: self.pick_tpl_color(key))
            b.grid(row=1, column=i, padx=(0, 6))
            self.tpl_swatches[k] = b

        ttk.Label(form, text="字体").grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 2))
        ff = ttk.Frame(form)
        ff.grid(row=5, column=0, columnspan=2, sticky="w")
        ttk.Label(ff, text="中文").grid(row=0, column=0)
        ttk.Entry(ff, textvariable=self.tpl_cjk, width=18).grid(row=0, column=1, padx=(4, 12))
        ttk.Label(ff, text="西文").grid(row=0, column=2)
        ttk.Entry(ff, textvariable=self.tpl_latin, width=18).grid(row=0, column=3, padx=(4, 0))

        ttk.Checkbutton(f, text="设为默认模板", variable=self.tpl_default).pack(anchor="w", pady=8)
        ttk.Button(f, text="保存模板", command=self.save_template).pack(anchor="w")

    # ---------- 左栏 ----------
    def refresh_left_briefs(self):
        self.brief_list.delete(0, "end")
        items = list(self.recent_briefs)
        if self.brief_path and self.brief_path not in items:
            items.insert(0, self.brief_path)
        self._brief_paths = items
        for label in wl.display_labels(items):
            self.brief_list.insert("end", label)

    def refresh_left_templates(self):
        self.tpl_list.delete(0, "end")
        self._tpl_names = template_names()
        for n in self._tpl_names:
            mark = "  ←当前" if n == self.template_name else ""
            self.tpl_list.insert("end", n + mark)

    def on_template_select(self, event=None):
        sel = self.tpl_list.curselection()
        if not sel or not getattr(self, "_tpl_names", []):
            return
        name = self._tpl_names[sel[0]]
        self.template_name = name
        self.template_var.set(name)
        self.refresh_left_templates()
        if name in self._tpl_names:
            index = self._tpl_names.index(name)
            self.tpl_list.selection_clear(0, "end")
            self.tpl_list.selection_set(index)
        self.update_status()

    def _on_brief_hover(self, event=None):
        index = self.brief_list.nearest(event.y) if event is not None else -1
        paths = getattr(self, "_brief_paths", [])
        if 0 <= index < len(paths):
            self.status_var.set(paths[index])

    def _show_brief_menu(self, event):
        paths = getattr(self, "_brief_paths", [])
        if not paths:
            return
        index = self.brief_list.nearest(event.y)
        self.brief_list.selection_clear(0, "end")
        self.brief_list.selection_set(index)
        self._brief_menu.tk_popup(event.x_root, event.y_root)

    def selected_brief_path(self):
        sel = self.brief_list.curselection()
        paths = getattr(self, "_brief_paths", [])
        if not sel or not paths or sel[0] >= len(paths):
            return None
        return paths[sel[0]]

    def reveal_selected_brief(self):
        path = self.selected_brief_path()
        if path:
            reveal_path(path)

    def remove_selected_brief(self):
        path = self.selected_brief_path()
        if not path:
            return
        if path in self.recent_briefs:
            self.recent_briefs.remove(path)
        self.refresh_left_briefs()
        self.update_status()

    def refresh_template_combo(self):
        names = template_names()
        self.template_combo["values"] = names
        if self.template_name not in names:
            self.template_name = names[0] if names else "default"
        self.template_var.set(self.template_name)

    def refresh_all(self):
        self.refresh_left_briefs()
        self.refresh_left_templates()
        self.refresh_template_combo()
        self.refresh_history()
        self.refresh_outline()
        self.update_status()

    def update_status(self):
        self.status_var.set(
            f"brief: {self.brief_path or '（未打开）'} | 模板: {self.template_name} | 输出: {self.output_var.get() or '（未设置）'}")

    # ---------- brief 操作 ----------
    def load_brief_into_editor(self, path):
        if not self._confirm_discard():
            return
        try:
            with open(path, "r", encoding="utf-8") as fh:
                text = fh.read()
        except Exception as e:
            messagebox.showerror("错误", f"读取失败：{e}")
            return
        self.brief_path = path
        self.brief_path_var.set(path)
        self._loading = True
        try:
            self.brief_text.delete("1.0", "end")
            self.brief_text.insert("1.0", text)
            self.brief_text.edit_modified(False)
        finally:
            self._loading = False
        if path not in self.recent_briefs:
            self.recent_briefs.insert(0, path)
        self._dirty = False
        self.refresh_left_briefs()
        self.refresh_outline()
        self._set_title()
        self.update_status()

    def open_brief(self):
        path = filedialog.askopenfilename(filetypes=[("YAML", "*.yaml"), ("所有文件", "*.*")], title="打开 brief")
        if path:
            self.load_brief_into_editor(path)

    def open_selected_brief(self):
        sel = self.brief_list.curselection()
        if sel:
            path = self.brief_list.get(sel[0])
            if os.path.isfile(path):
                self.load_brief_into_editor(path)

    def new_brief(self):
        if not self._confirm_discard():
            return
        path = filedialog.asksaveasfilename(defaultextension=".yaml", filetypes=[("YAML", "*.yaml")], title="新建 brief")
        if not path:
            return
        skeleton = (
            "meta:\n"
            "  title: 组会汇报\n"
            "  presenter: \n"
            "  date: \n"
            "sections:\n"
            "  - type: cover\n"
            "    title: 组会汇报\n"
            "  - type: points\n"
            "    heading: 研究背景\n"
            "    points:\n"
            "      - \n"
        )
        try:
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(skeleton)
        except Exception as e:
            messagebox.showerror("错误", f"写入失败：{e}")
            return
        self.load_brief_into_editor(path)

    def save_brief(self):
        if not self.brief_path:
            return self.save_brief_as()
        try:
            with open(self.brief_path, "w", encoding="utf-8") as fh:
                fh.write(self.brief_text.get("1.0", "end"))
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{e}")
            return False
        self._dirty = False
        self._set_title()
        self.update_status()
        return True

    def save_brief_as(self):
        path = filedialog.asksaveasfilename(defaultextension=".yaml", filetypes=[("YAML", "*.yaml")], title="另存 brief")
        if not path:
            return False
        previous_path = self.brief_path
        self.brief_path = path
        if not self.save_brief():
            self.brief_path = previous_path
            return False
        self.brief_path_var.set(path)
        return True

    def validate_brief(self):
        import yaml
        text = self.brief_text.get("1.0", "end")
        try:
            brief = yaml.safe_load(text)
        except yaml.YAMLError as e:
            mark = getattr(e, "problem_mark", None)
            line = mark.line + 1 if mark is not None else None
            detail = getattr(e, "problem", None) or str(e)
            self.validate_var.set(f"YAML 语法错误：第 {line} 行 {detail}" if line else f"YAML 语法错误：{detail}")
            if line:
                self.brief_text.mark_set("insert", f"{line}.0")
                self.brief_text.see(f"{line}.0")
                self.brief_text.focus_set()
            return
        base_dir = os.path.dirname(os.path.abspath(self.brief_path)) if self.brief_path else "."
        try:
            g.validate(brief, base_dir)
        except ValueError as e:
            self.validate_var.set(f"校验失败：{e}")
            return
        pages = wl.estimate_pages(brief)
        warnings = wl.compute_warnings(brief)
        if warnings:
            self.validate_var.set(f"校验通过（{pages} 段，{len(warnings)} 条提醒）")
            self._set_result("生成前检查：\n" + "\n".join("• " + w for w in warnings) + "\n")
        else:
            self.validate_var.set(f"校验通过（{pages} 段）")

    # ---------- 生成 ----------
    def browse_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("PowerPoint", "*.pptx")], title="输出文件")
        if path:
            self.output_var.set(path)

    def browse_history(self):
        d = filedialog.askdirectory(title="历史目录")
        if d:
            self.history_var.set(d)

    def _ensure_brief_saved(self):
        if not self.brief_path:
            messagebox.showwarning("提示", "请先保存 brief")
            return False
        return self.save_brief()

    def do_generate(self, slide):
        if not self._ensure_brief_saved():
            return
        brief, error = self._current_brief()
        if error:
            messagebox.showerror("无法生成", error)
            return
        pages = wl.estimate_pages(brief)
        warnings = wl.compute_warnings(brief)
        if warnings and not messagebox.askyesno("生成前检查", wl.format_preflight(pages, warnings)):
            self._set_result(f"已取消生成（{len(warnings)} 条提醒未处理）。\n")
            return
        summary = f"预检：{pages} 页" + (f"，{len(warnings)} 条提醒" if warnings else "，无提醒")
        self._set_result(summary + "\n")
        output = self.output_var.get().strip()
        if not output:
            output = os.path.splitext(self.brief_path)[0] + ".pptx"
            self.output_var.set(output)
        template = self.template_var.get()
        history = self.history_var.get().strip() or None
        self._set_busy(True)
        try:
            n, tn = run_generate(self.brief_path, output, template, slide=slide, history_dir=history)
        except Exception as e:
            self._set_result(f"失败：{e}\n")
            messagebox.showerror("生成失败", str(e))
            return
        finally:
            self._set_busy(False)
        self.template_name = template
        self._set_result(f"成功：共 {n} 页 · 模板 {tn} → {output}\n")
        if history:
            self.refresh_history()
        self.refresh_left_templates()
        self.refresh_template_combo()
        self.update_status()

    def do_generate_slide(self):
        s = self.slide_var.get().strip()
        if not s.isdigit():
            messagebox.showwarning("提示", "请输入要重建的页码 N")
            return
        self.do_generate(int(s))

    def _set_result(self, msg):
        self.result_text.configure(state="normal")
        self.result_text.insert("end", msg)
        self.result_text.see("end")
        self.result_text.configure(state="disabled")

    # ---------- 状态、快捷键与预检 ----------
    def _current_brief(self):
        """解析编辑器里的 brief 并做硬性校验，返回 (brief, error)。"""
        import yaml
        text = self.brief_text.get("1.0", "end")
        try:
            brief = yaml.safe_load(text)
        except yaml.YAMLError as e:
            mark = getattr(e, "problem_mark", None)
            where = f"第 {mark.line + 1} 行 " if mark is not None else ""
            return None, f"YAML 语法错误：{where}{getattr(e, 'problem', None) or e}"
        if not isinstance(brief, dict):
            return None, "brief 顶层必须是包含 meta 和 sections 的映射"
        base_dir = os.path.dirname(os.path.abspath(self.brief_path)) if self.brief_path else "."
        try:
            g.validate(brief, base_dir)
        except ValueError as e:
            return None, str(e)
        return brief, None

    def _parsed_brief(self):
        """只做 YAML 解析，不校验，供大纲在内容未填全时仍可显示。"""
        import yaml
        try:
            return yaml.safe_load(self.brief_text.get("1.0", "end")), None
        except yaml.YAMLError as e:
            mark = getattr(e, "problem_mark", None)
            where = f"第 {mark.line + 1} 行 " if mark is not None else ""
            return None, f"YAML 语法错误：{where}{getattr(e, 'problem', None) or e}"

    # ---------- 大纲 ----------
    def refresh_outline(self, select=None):
        if not hasattr(self, "outline_tree"):
            return
        self.outline_tree.delete(*self.outline_tree.get_children())
        brief, error = self._parsed_brief()
        if error or not isinstance(brief, dict):
            self.status_var.set(f"大纲不可用：{error or 'brief 顶层必须是映射'}")
            return
        for row in wl.section_outline(brief):
            self.outline_tree.insert("", "end", iid=str(row["index"]),
                                     values=(row["index"], row["label"], row["title"]))
        children = self.outline_tree.get_children()
        if select is not None and str(select) in children:
            self.outline_tree.selection_set(str(select))
            self.outline_tree.see(str(select))
        self.status_var.set(f"大纲：{len(children)} 段")

    def _outline_selected_index(self):
        sel = self.outline_tree.selection()
        if not sel:
            return None
        try:
            return int(sel[0]) - 1
        except ValueError:
            return None

    def _replace_brief_text(self, text):
        """整段替换编辑器内容，并同步脏标记与大纲。"""
        self._loading = True
        try:
            self.brief_text.delete("1.0", "end")
            self.brief_text.insert("1.0", text)
            self.brief_text.edit_modified(False)
        finally:
            self._loading = False
        self._dirty = True
        self._set_title()

    def move_outline(self, delta):
        index = self._outline_selected_index()
        if index is None:
            messagebox.showinfo("提示", "请先在大纲里选一段")
            return
        text = self.brief_text.get("1.0", "end-1c")
        new_text = wl.move_section_text(text, index, delta)
        if new_text == text:
            return
        self._replace_brief_text(new_text)
        self.refresh_outline(select=index + delta + 1)

    def remove_outline(self):
        index = self._outline_selected_index()
        if index is None:
            messagebox.showinfo("提示", "请先在大纲里选一段")
            return
        if not messagebox.askyesno("删除该段", f"删除第 {index + 1} 段？文本区可用 Ctrl+Z 撤销。"):
            return
        text = self.brief_text.get("1.0", "end-1c")
        new_text = wl.remove_section_text(text, index)
        if new_text == text:
            return
        self._replace_brief_text(new_text)
        self.refresh_outline(select=index + 1)

    def insert_outline(self, typ):
        text = self.brief_text.get("1.0", "end-1c")
        try:
            new_text = wl.append_snippet(text, typ)
        except ValueError as e:
            messagebox.showerror("无法插入", str(e))
            return
        self._replace_brief_text(new_text)
        brief, error = self._parsed_brief()
        rows = wl.section_outline(brief) if not error else []
        self.refresh_outline(select=len(rows) or None)
        if rows:
            self.status_var.set(f"已追加「{wl.type_label(typ)}」骨架，请补齐第 {len(rows)} 段的占位内容")

    def _on_outline_double(self, event=None):
        index = self._outline_selected_index()
        if index is None:
            return
        spans = wl.section_spans(self.brief_text.get("1.0", "end-1c"))
        if not (0 <= index < len(spans)):
            return
        start, end = spans[index]
        self.brief_text.tag_remove("sel", "1.0", "end")
        self.brief_text.tag_add("sel", f"{start + 1}.0", f"{end}.end")
        self.brief_text.mark_set("insert", f"{start + 1}.0")
        self.brief_text.see(f"{start + 1}.0")
        self.brief_text.focus_set()
        self.nb.select(0)

    def _on_tab_changed(self, event=None):
        if hasattr(self, "outline_tab") and self.nb.select() == str(self.outline_tab):
            self.refresh_outline()

    def _set_busy(self, busy):
        for btn in self._generate_buttons:
            btn.config(state="disabled" if busy else "normal")
        self.config(cursor="watch" if busy else "")
        self.update_idletasks()

    def _set_title(self):
        name = os.path.basename(self.brief_path) if self.brief_path else "未打开 brief"
        mark = "● " if self._dirty else ""
        self.title(f"{mark}组会 PPT 工作台 — {name}")

    def _on_text_modified(self, event=None):
        if self._loading:
            self.brief_text.edit_modified(False)
            return
        if self.brief_text.edit_modified():
            self.brief_text.edit_modified(False)
            self._dirty = True
            self._set_title()

    def _confirm_discard(self):
        """返回 True 表示可以继续（无改动、已保存或用户选择放弃）。"""
        if not self._dirty:
            return True
        answer = messagebox.askyesnocancel(
            "未保存的修改",
            "当前 brief 有未保存的修改。\n\n是：保存后继续\n否：放弃修改\n取消：留在当前文件")
        if answer is None:
            return False
        if answer:
            return bool(self.save_brief())
        return True

    @staticmethod
    def _shortcut(action):
        action()
        return "break"

    def _bind_shortcuts(self):
        self.bind_all("<Control-s>", lambda e: self._shortcut(self.save_brief))
        self.bind_all("<Control-o>", lambda e: self._shortcut(self.open_brief))
        self.bind_all("<Control-n>", lambda e: self._shortcut(self.new_brief))
        self.bind_all("<Control-Return>", lambda e: self._shortcut(lambda: self.do_generate(None)))
        self.bind_all("<F5>", lambda e: self._shortcut(self.refresh_all))

    # ---------- 历史与预览 ----------
    def refresh_history(self):
        self.version_list.delete(0, "end")
        hd = self.history_var.get().strip()
        if not hd:
            return
        self._versions = list_versions(hd)
        for v in self._versions:
            label = v["id"]
            if v.get("legacy"):
                label += "（旧版）"
            elif v.get("error"):
                label += "（无法读取）"
            elif v.get("template_name"):
                label += f"（{v['template_name']}）"
            self.version_list.insert("end", label)
        latest = find_latest_preview(hd)
        if latest:
            self.show_preview(latest)

    def on_version_select(self):
        hd = self.history_var.get().strip()
        if not hd:
            return
        sel = self.version_list.curselection()
        if not sel:
            return
        idx = sel[0]
        if hasattr(self, "_versions") and idx < len(self._versions):
            v = self._versions[idx]
            if v["preview"]:
                self.show_preview(v["preview"])

    def show_preview(self, path, max_w=880):
        try:
            img = tk.PhotoImage(file=path)
            w = img.width()
            if w > max_w:
                img = img.subsample(max(1, round(w / max_w)), max(1, round(w / max_w)))
            self.preview_img = img
            self.preview_label.config(image=img, text="")
        except Exception as e:
            self.preview_label.config(image="", text=f"无法显示预览：{e}")

    def rebuild_version(self):
        hd = self.history_var.get().strip()
        if not hd:
            messagebox.showwarning("提示", "请先设置历史目录")
            return
        sel = self.version_list.curselection()
        if not sel or not hasattr(self, "_versions"):
            messagebox.showwarning("提示", "请先选择一个版本")
            return
        v = self._versions[sel[0]]
        if v.get("error"):
            messagebox.showerror("无法重建", v["error"])
            return
        template_name = None
        if v.get("legacy"):
            template_name = self.template_var.get()
            if not messagebox.askyesno(
                    "旧版历史",
                    "该旧版本没有归档图片和模板，仅在原始图片仍存在时可重建。\n"
                    f"将使用当前模板“{template_name}”，是否继续？"):
                return
        out = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("PowerPoint", "*.pptx")],
                                           title="重建输出文件")
        if not out:
            return
        try:
            n, tn = run_generate(v["yaml"], out, template_name)
        except Exception as e:
            messagebox.showerror("重建失败", str(e))
            return
        self._set_result(f"已从 {v['id']} 重建：共 {n} 页 → {out}\n")
        self.output_var.set(out)
        self.update_status()

    # ---------- 模板 ----------
    def import_pptx(self):
        path = filedialog.askopenfilename(filetypes=[("PowerPoint", "*.pptx")], title="选择 PPT 文件")
        if not path:
            return
        try:
            theme = tw.extract_theme(path)
        except Exception as e:
            messagebox.showerror("错误", f"解析失败：{e}")
            return
        self.tpl_colors = theme["colors"]
        self.tpl_fonts = theme["fonts"]
        self.tpl_slide = theme["slide"]
        self.tpl_name.set(os.path.splitext(os.path.basename(path))[0])
        self.tpl_desc.set("")
        self.tpl_cjk.set(theme["fonts"]["cjk"])
        self.tpl_latin.set(theme["fonts"]["latin"])
        self.tpl_default.set(False)
        for k, b in self.tpl_swatches.items():
            b.config(text=self.tpl_colors[k], bg=self.tpl_colors[k], fg=tw.Wizard._fg(self.tpl_colors[k]))
        s = theme["slide"]
        self.tpl_summary.set(
            f"已提取：{s['aspect']}（{s['width_in']}″ × {s['height_in']}″）· "
            f"主色 {theme['colors']['primary']} · 强调 {theme['colors']['accent']}")

    def pick_tpl_color(self, key):
        _, hexc = colorchooser.askcolor(color=self.tpl_colors[key], title=tw.COLOR_LABELS[key])
        if hexc:
            self.tpl_colors[key] = hexc
            self.tpl_swatches[key].config(text=hexc, bg=hexc, fg=tw.Wizard._fg(hexc))

    def save_template(self):
        import yaml
        name = self.tpl_name.get().strip()
        try:
            out = tw.template_output_path(TDIR, name)
        except ValueError as e:
            messagebox.showwarning("提示", str(e))
            return
        data = {
            "name": name,
            "description": self.tpl_desc.get().strip(),
            "slide": self.tpl_slide,
            "colors": self.tpl_colors,
            "fonts": {"cjk": self.tpl_cjk.get().strip() or tw.DEFAULT_FONTS["cjk"],
                      "latin": self.tpl_latin.get().strip() or tw.DEFAULT_FONTS["latin"]},
            "default": bool(self.tpl_default.get()),
        }
        if os.path.exists(out) and not messagebox.askyesno("覆盖", f"{out} 已存在，是否覆盖？"):
            return
        try:
            with open(out, "w", encoding="utf-8") as fh:
                yaml.safe_dump(data, fh, allow_unicode=True, sort_keys=False)
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{e}")
            return
        if data["default"]:
            self.template_name = name
        self.refresh_left_templates()
        self.refresh_template_combo()
        self.tpl_summary.set(f"已保存：{out}")
        self.update_status()

    # ---------- 其他 ----------
    def open_output(self):
        out = self.output_var.get().strip()
        if not out or not os.path.isfile(out):
            messagebox.showwarning("提示", "输出文件不存在")
            return
        open_path(out)

    def reveal_output(self):
        out = self.output_var.get().strip()
        if not out or not os.path.isfile(out):
            messagebox.showwarning("提示", "输出文件不存在")
            return
        reveal_path(out)

    def save_state(self):
        st = {
            "recent_briefs": self.recent_briefs[:20],
            "current_brief": self.brief_path or "",
            "template": self.template_name,
            "output": self.output_var.get(),
            "history_dir": self.history_var.get(),
            "geometry": self.geometry(),
            "tab": self.nb.index(self.nb.select()) if self.nb.tabs() else 0,
        }
        try:
            os.makedirs(os.path.dirname(STATE_PATH), exist_ok=True)
            with open(STATE_PATH, "w", encoding="utf-8") as f:
                json.dump(st, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def on_close(self):
        if not self._confirm_discard():
            return
        self.save_state()
        self.destroy()


def main():
    Workbench().mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())

