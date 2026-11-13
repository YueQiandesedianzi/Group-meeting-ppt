#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""模板向导（Tkinter）：首次使用导入/选择/配置模板（主题级）。"""
import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, colorchooser, messagebox
import zipfile

import yaml
from lxml import etree

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

DEFAULT_COLORS = {
    "primary": "#1F3864",
    "accent": "#2E74B5",
    "light": "#D6E4F0",
    "text": "#333333",
    "muted": "#808080",
    "line": "#C9C9C9",
    "background": "#FFFFFF",
}
DEFAULT_FONTS = {"cjk": "微软雅黑", "latin": "Calibri"}
COLOR_KEYS = ["primary", "accent", "light", "text", "muted", "line", "background"]
COLOR_LABELS = {
    "primary": "主色", "accent": "强调色", "light": "浅底", "text": "正文",
    "muted": "次要", "line": "分隔线", "background": "背景",
}


def template_output_path(tdir, name):
    """Return a safe template filename confined to the template directory."""
    if (not name or name.lower() == "default" or name in {".", ".."}
            or name != name.strip() or name.endswith(".")
            or any(c in name for c in '/\\:*?"<>|')
            or any(ord(c) < 32 for c in name)):
        raise ValueError("模板名无效：请使用普通文件名，不能使用 default 或路径字符")
    stem = name.split(".", 1)[0].upper()
    if stem in {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)),
                *(f"LPT{i}" for i in range(1, 10))}:
        raise ValueError("模板名不能使用系统保留名称")
    directory = os.path.realpath(tdir)
    path = os.path.join(directory, name + ".yaml")
    if os.path.commonpath([directory, os.path.realpath(path)]) != directory:
        raise ValueError("模板文件必须位于模板目录内")
    return path


def extract_theme(path):
    """主题级提取：页面尺寸 + 配色 + 字体。失败时保留默认值。"""
    from pptx import Presentation

    prs = Presentation(path)
    w = prs.slide_width / 914400.0
    h = prs.slide_height / 914400.0
    ratio = w / h
    aspect = "16:9" if abs(ratio - 16 / 9) < 0.05 else ("4:3" if abs(ratio - 4 / 3) < 0.05 else "custom")

    colors = dict(DEFAULT_COLORS)
    fonts = dict(DEFAULT_FONTS)

    try:
        with zipfile.ZipFile(path) as z:
            names = [n for n in z.namelist() if n.startswith("ppt/theme/theme1.xml")]
            if names:
                root = etree.fromstring(z.read(names[0]))
                scheme = root.find(".//a:clrScheme", NS)
                if scheme is not None:
                    def col(tag, default):
                        el = scheme.find("a:" + tag, NS)
                        if el is not None:
                            srgb = el.find(".//a:srgbClr", NS)
                            if srgb is not None and srgb.get("val"):
                                return "#" + srgb.get("val")
                        return default

                    colors["text"] = col("dk1", colors["text"])
                    colors["background"] = col("lt1", colors["background"])
                    colors["primary"] = col("dk2", col("accent1", colors["primary"]))
                    colors["accent"] = col("accent1", colors["accent"])
                    colors["light"] = col("lt2", colors["light"])
                fscheme = root.find(".//a:fontScheme", NS)
                if fscheme is not None:
                    def fcol(kind, sub):
                        el = fscheme.find(f"a:{kind}/a:{sub}", NS)
                        return el.get("typeface") if el is not None else None

                    latin = fcol("majorFont", "latin") or fcol("minorFont", "latin")
                    ea = fcol("majorFont", "ea") or fcol("minorFont", "ea")
                    if latin:
                        fonts["latin"] = latin
                    if ea:
                        fonts["cjk"] = ea
    except Exception:
        pass

    return {
        "slide": {"aspect": aspect, "width_in": round(w, 3), "height_in": round(h, 3)},
        "colors": colors,
        "fonts": fonts,
    }


class Wizard(tk.Tk):
    def __init__(self, tdir):
        super().__init__()
        self.tdir = tdir
        self.title("组会 PPT 模板向导")
        self.geometry("660x600")
        self.data = None
        self.frame = None
        self.show_entry()

    def clear(self):
        if self.frame is not None:
            self.frame.destroy()

    # ---- 入口 ----
    def show_entry(self):
        self.clear()
        f = ttk.Frame(self, padding=16)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="模板向导", font=("", 16, "bold")).pack(anchor="w", pady=(0, 4))
        ttk.Label(f, text="选择模板来源（模板将保存到 templates/ 目录）").pack(anchor="w", pady=(0, 12))
        ttk.Button(f, text="使用内置 default", command=self.use_default).pack(fill="x", pady=4)
        ttk.Button(f, text="导入 .pptx 作为模板", command=self.show_import).pack(fill="x", pady=4)
        ttk.Button(f, text="加载已有模板档案", command=self.show_load).pack(fill="x", pady=4)
        ttk.Button(f, text="退出", command=self.destroy).pack(fill="x", pady=12)
        self.frame = f

    def use_default(self):
        messagebox.showinfo("模板", "将使用内置 default 模板。")
        self.destroy()

    def show_load(self):
        path = filedialog.askopenfilename(
            initialdir=self.tdir, filetypes=[("YAML", "*.yaml")], title="选择模板档案")
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as fh:
                data = yaml.safe_load(fh) or {}
        except Exception as e:
            messagebox.showerror("错误", f"读取失败：{e}")
            return
        self.data = {
            "name": data.get("name", ""),
            "description": data.get("description", ""),
            "slide": data.get("slide") or {"aspect": "16:9", "width_in": 13.333, "height_in": 7.5},
            "colors": dict(DEFAULT_COLORS, **(data.get("colors") or {})),
            "fonts": dict(DEFAULT_FONTS, **(data.get("fonts") or {})),
            "default": bool(data.get("default", False)),
        }
        self.show_config()

    # ---- 导入 ----
    def show_import(self):
        self.clear()
        f = ttk.Frame(self, padding=16)
        f.pack(fill="both", expand=True)
        self.import_preview = tk.StringVar(value="（未选择文件）")
        ttk.Label(f, text="导入 .pptx（提取配色 / 字体 / 页面尺寸）", font=("", 13, "bold")).pack(anchor="w")
        ttk.Button(f, text="选择 .pptx 文件…", command=self.pick_pptx).pack(anchor="w", pady=8)
        ttk.Label(f, textvariable=self.import_preview, wraplength=600, justify="left").pack(anchor="w", pady=4)
        nav = ttk.Frame(f)
        nav.pack(fill="x", pady=12)
        ttk.Button(nav, text="上一步", command=self.show_entry).pack(side="left")
        ttk.Button(nav, text="下一步：配置", command=self.after_import).pack(side="right")
        self.frame = f

    def pick_pptx(self):
        path = filedialog.askopenfilename(filetypes=[("PowerPoint", "*.pptx")], title="选择 PPT 文件")
        if not path:
            return
        try:
            theme = extract_theme(path)
        except Exception as e:
            messagebox.showerror("错误", f"解析失败：{e}")
            return
        self.data = {
            "name": os.path.splitext(os.path.basename(path))[0],
            "description": "",
            "slide": theme["slide"],
            "colors": theme["colors"],
            "fonts": theme["fonts"],
            "default": False,
        }
        s = theme["slide"]
        self.import_preview.set(
            f"已提取：{s['aspect']}（{s['width_in']}″ × {s['height_in']}″）\n"
            f"字体：中文 {theme['fonts']['cjk']} · 西文 {theme['fonts']['latin']}\n"
            f"主色 {theme['colors']['primary']} · 强调 {theme['colors']['accent']}"
        )

    def after_import(self):
        if self.data is None:
            messagebox.showwarning("提示", "请先选择 .pptx 文件")
            return
        self.show_config()

    # ---- 配置 ----
    def show_config(self):
        self.clear()
        f = ttk.Frame(self, padding=16)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="配置模板", font=("", 13, "bold")).grid(row=0, column=0, columnspan=2, sticky="w")

        ttk.Label(f, text="模板名").grid(row=1, column=0, sticky="w", pady=(10, 2))
        self.name_var = tk.StringVar(value=self.data.get("name", ""))
        ttk.Entry(f, textvariable=self.name_var, width=42).grid(row=1, column=1, sticky="w")

        ttk.Label(f, text="描述").grid(row=2, column=0, sticky="w", pady=(8, 2))
        self.desc_var = tk.StringVar(value=self.data.get("description", ""))
        ttk.Entry(f, textvariable=self.desc_var, width=42).grid(row=2, column=1, sticky="w")

        ttk.Label(f, text="配色（点击色块修改）").grid(row=3, column=0, columnspan=2, sticky="w", pady=(12, 2))
        cf = ttk.Frame(f)
        cf.grid(row=4, column=0, columnspan=2, sticky="w")
        self.swatches = {}
        for i, k in enumerate(COLOR_KEYS):
            ttk.Label(cf, text=COLOR_LABELS[k]).grid(row=0, column=i, padx=(0, 2))
            b = tk.Button(cf, text=self.data["colors"][k], bg=self.data["colors"][k],
                          fg=self._fg(self.data["colors"][k]), width=9,
                          command=lambda key=k: self.pick_color(key))
            b.grid(row=1, column=i, padx=(0, 6))
            self.swatches[k] = b

        ttk.Label(f, text="字体").grid(row=5, column=0, columnspan=2, sticky="w", pady=(12, 2))
        ff = ttk.Frame(f)
        ff.grid(row=6, column=0, columnspan=2, sticky="w")
        ttk.Label(ff, text="中文").grid(row=0, column=0)
        self.cjk_var = tk.StringVar(value=self.data["fonts"].get("cjk", ""))
        ttk.Entry(ff, textvariable=self.cjk_var, width=20).grid(row=0, column=1, padx=(4, 12))
        ttk.Label(ff, text="西文").grid(row=0, column=2)
        self.latin_var = tk.StringVar(value=self.data["fonts"].get("latin", ""))
        ttk.Entry(ff, textvariable=self.latin_var, width=20).grid(row=0, column=3, padx=(4, 0))

        self.default_var = tk.BooleanVar(value=self.data.get("default", False))
        ttk.Checkbutton(f, text="设为默认模板", variable=self.default_var).grid(
            row=7, column=0, columnspan=2, sticky="w", pady=(12, 2))

        nav = ttk.Frame(f)
        nav.grid(row=8, column=0, columnspan=2, sticky="ew", pady=14)
        ttk.Button(nav, text="上一步", command=self.show_entry).pack(side="left")
        ttk.Button(nav, text="保存模板", command=self.save).pack(side="right")
        self.frame = f

    @staticmethod
    def _fg(hexcolor):
        try:
            r, g, b = int(hexcolor[1:3], 16), int(hexcolor[3:5], 16), int(hexcolor[5:7], 16)
            return "white" if (r * 0.299 + g * 0.587 + b * 0.114) < 128 else "black"
        except Exception:
            return "black"

    def pick_color(self, key):
        _, hexc = colorchooser.askcolor(color=self.data["colors"][key], title=COLOR_LABELS[key])
        if hexc:
            self.data["colors"][key] = hexc
            self.swatches[key].config(bg=hexc, text=hexc, fg=self._fg(hexc))

    def save(self):
        name = self.name_var.get().strip()
        try:
            out = template_output_path(self.tdir, name)
        except ValueError as e:
            messagebox.showwarning("提示", str(e))
            return
        self.data["name"] = name
        self.data["description"] = self.desc_var.get().strip()
        self.data["fonts"] = {
            "cjk": self.cjk_var.get().strip() or DEFAULT_FONTS["cjk"],
            "latin": self.latin_var.get().strip() or DEFAULT_FONTS["latin"],
        }
        self.data["default"] = bool(self.default_var.get())
        if os.path.exists(out):
            if not messagebox.askyesno("覆盖", f"{out} 已存在，是否覆盖？"):
                return
        try:
            with open(out, "w", encoding="utf-8") as fh:
                yaml.safe_dump(self.data, fh, allow_unicode=True, sort_keys=False)
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{e}")
            return
        messagebox.showinfo("完成", f"已保存模板：{out}")
        self.destroy()


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    tdir = os.path.normpath(os.path.join(script_dir, "..", "templates"))
    os.makedirs(tdir, exist_ok=True)
    Wizard(tdir).mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())

