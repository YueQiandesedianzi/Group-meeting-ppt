#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""组会 PPT 工作台（Tkinter）：brief 编辑 + 模板 + 生成 + 历史预览。

复用 generate_ppt.py 与 template_wizard.py 的引擎逻辑，不复制布局/校验/渲染代码。
"""
import os
import sys
import json
from types import SimpleNamespace

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

import generate_ppt as g  # noqa: E402
import template_wizard as tw  # noqa: E402

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


# ---------------- 工作台界面 ----------------

class Workbench(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("组会 PPT 工作台")
        self.geometry("1180x760")

        st = load_state()
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
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------- 布局 ----------
    def _build_toolbar(self):
        bar = ttk.Frame(self, padding=(8, 6))
        bar.pack(fill="x")
        ttk.Button(bar, text="新建 brief", command=self.new_brief).pack(side="left", padx=2)
        ttk.Button(bar, text="打开 brief…", command=self.open_brief).pack(side="left", padx=2)
        ttk.Button(bar, text="刷新", command=self.refresh_all).pack(side="left", padx=2)
        ttk.Button(bar, text="打开输出…", command=self.open_output).pack(side="left", padx=2)

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
        ttk.Label(left, text="模板").pack(anchor="w")
        self.tpl_list = tk.Listbox(left, height=10)
        self.tpl_list.pack(fill="both", expand=True, pady=(2, 0))

        # 主区 Notebook
        self.nb = ttk.Notebook(body)
        self.nb.pack(side="left", fill="both", expand=True)
        self._build_tab_brief()
        self._build_tab_generate()
        self._build_tab_history()
        self._build_tab_template()

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

        row2 = ttk.Frame(f)
        row2.pack(fill="x", pady=(6, 0))
        ttk.Button(row2, text="校验", command=self.validate_brief).pack(side="left")
        self.validate_var = tk.StringVar(value="")
        ttk.Label(row2, textvariable=self.validate_var).pack(side="left", padx=8)

    # ---------- Tab 2: 生成 ----------
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
        ttk.Button(btns, text="全量生成", command=lambda: self.do_generate(None)).pack(side="left", padx=2)
        ttk.Button(btns, text="单页重建", command=self.do_generate_slide).pack(side="left", padx=2)

        self.result_text = tk.Text(f, height=8, wrap="word", state="disabled")
        self.result_text.pack(fill="both", expand=True)

    # ---------- Tab 3: 预览与历史 ----------
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

    # ---------- Tab 4: 模板 ----------
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
        for p in items:
            self.brief_list.insert("end", p)

    def refresh_left_templates(self):
        self.tpl_list.delete(0, "end")
        for n in template_names():
            mark = "  *" if n == self.template_name else ""
            self.tpl_list.insert("end", n + mark)

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
        self.update_status()

    def update_status(self):
        self.status_var.set(
            f"brief: {self.brief_path or '（未打开）'} | 模板: {self.template_name} | 输出: {self.output_var.get() or '（未设置）'}")

    # ---------- brief 操作 ----------
    def load_brief_into_editor(self, path):
        try:
            with open(path, "r", encoding="utf-8") as fh:
                text = fh.read()
        except Exception as e:
            messagebox.showerror("错误", f"读取失败：{e}")
            return
        self.brief_path = path
        self.brief_path_var.set(path)
        self.brief_text.delete("1.0", "end")
        self.brief_text.insert("1.0", text)
        if path not in self.recent_briefs:
            self.recent_briefs.insert(0, path)
        self.refresh_left_briefs()
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
            self.save_brief_as()
            return
        try:
            with open(self.brief_path, "w", encoding="utf-8") as fh:
                fh.write(self.brief_text.get("1.0", "end"))
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{e}")
            return
        self.update_status()

    def save_brief_as(self):
        path = filedialog.asksaveasfilename(defaultextension=".yaml", filetypes=[("YAML", "*.yaml")], title="另存 brief")
        if not path:
            return
        self.brief_path = path
        self.save_brief()
        self.brief_path_var.set(path)

    def validate_brief(self):
        import yaml
        text = self.brief_text.get("1.0", "end")
        try:
            brief = yaml.safe_load(text)
        except Exception as e:
            self.validate_var.set(f"YAML 语法错误：{e}")
            return
        base_dir = os.path.dirname(os.path.abspath(self.brief_path)) if self.brief_path else "."
        try:
            g.validate(brief, base_dir)
            n = len(brief.get("sections", []))
            self.validate_var.set(f"校验通过（{n} 段）")
        except ValueError as e:
            self.validate_var.set(f"校验失败：{e}")

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
        self.save_brief()
        return True

    def do_generate(self, slide):
        if not self._ensure_brief_saved():
            return
        output = self.output_var.get().strip()
        if not output:
            output = os.path.splitext(self.brief_path)[0] + ".pptx"
            self.output_var.set(output)
        template = self.template_var.get()
        history = self.history_var.get().strip() or None
        try:
            n, tn = run_generate(self.brief_path, output, template, slide=slide, history_dir=history)
        except Exception as e:
            self._set_result(f"失败：{e}\n")
            messagebox.showerror("生成失败", str(e))
            return
        self.template_name = template
        self._set_result(f"成功：共 {n} 页 · 模板 {tn} → {output}\n")
        if history:
            self.refresh_history()
        self.refresh_left_templates()
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
        if not name or name.lower() == "default":
            messagebox.showwarning("提示", "请填写模板名（不能与内置 default 重名）")
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
        out = os.path.join(TDIR, name + ".yaml")
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
        os.startfile(out)  # Windows

    def save_state(self):
        st = {
            "recent_briefs": self.recent_briefs[:20],
            "current_brief": self.brief_path or "",
            "template": self.template_name,
            "output": self.output_var.get(),
            "history_dir": self.history_var.get(),
        }
        try:
            os.makedirs(os.path.dirname(STATE_PATH), exist_ok=True)
            with open(STATE_PATH, "w", encoding="utf-8") as f:
                json.dump(st, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def on_close(self):
        self.save_state()
        self.destroy()


def main():
    Workbench().mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
