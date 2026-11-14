#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""工作台 UI 行为测试：驱动真实 Tk 实例，但不进入 mainloop。

- 没有可用显示环境时整组跳过（探测在子进程里做，不污染当前进程）。
- 整个测试类共用一个 Tk 根：多根解释器会让 tkinter.Variable.__del__ 报
  "main thread is not in main loop"，所以用例之间靠 reset_app() 复位，而不是重建。
- 每个用例使用独立临时目录，不读取或改写用户的 .bohrium/ppt-workbench.json。
"""
import os
import subprocess
import sys
import tempfile
import time
import unittest
from unittest.mock import patch

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))

PROJECT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEMO_BRIEF = os.path.join(PROJECT, "assets", "demo_brief.yaml")

SAMPLE = """meta:
  title: 测试
sections:
  - type: cover
    title: 封面

  - type: points
    heading: 方法
    points:
      - A

  - type: conclusion
    heading: 小结
    points:
      - B

notes: keep me
"""


def _tk_available():
    code = "import tkinter;r=tkinter.Tk();r.withdraw();r.destroy();print('tk-ok')"
    try:
        done = subprocess.run([sys.executable, "-c", code], capture_output=True, timeout=90)
    except Exception:
        return False
    return done.returncode == 0 and b"tk-ok" in done.stdout


TK_OK = _tk_available()
if TK_OK:
    import tkinter as tk
    import workbench as wb
    import workbench_logic as wl

STATE_DIR = tempfile.mkdtemp(prefix="wb-ui-state-")


@unittest.skipUnless(TK_OK, "需要可用的 Tk 显示环境")
class WorkbenchUITests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.state_patch = patch.object(wb, "STATE_PATH", os.path.join(STATE_DIR, "state.json"))
        cls.state_patch.start()
        cls.app = wb.Workbench()
        cls.app.withdraw()
        cls.app.update()

    @classmethod
    def tearDownClass(cls):
        try:
            cls.app.destroy()
        except Exception:
            pass
        tk._default_root = None
        cls.state_patch.stop()

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="wb-ui-case-")
        self.reset_app()

    # ---------- 夹具 ----------
    def reset_app(self):
        app = self.app
        if getattr(app, "_outline_job", None):
            try:
                app.after_cancel(app._outline_job)
            except Exception:
                pass
            app._outline_job = None
        app._dirty = False
        app._loading = True
        try:
            app.brief_text.delete("1.0", "end")
        finally:
            app._loading = False
        app.brief_path = None
        app.brief_path_var.set("")
        app.output_var.set("")
        app.history_var.set("")
        app.slide_var.set("")
        app.template_name = "default"
        app.template_var.set("default")
        app._versions = []
        app.version_list.delete(0, "end")
        app.result_text.configure(state="normal")
        app.result_text.delete("1.0", "end")
        app.result_text.configure(state="disabled")
        app.nb.select(0)
        app.refresh_outline()
        app.refresh_slide_choices()
        app.update()

    def load_text(self, text):
        self.app.brief_text.delete("1.0", "end")
        self.app.brief_text.insert("1.0", text)
        self.app.brief_text.edit_modified(False)
        self.app.update()

    def outline_rows(self):
        return [self.app.outline_tree.item(i, "values")
                for i in self.app.outline_tree.get_children()]

    # ---------- 结构 ----------
    def test_five_tabs_in_reading_order(self):
        names = [self.app.nb.tab(t, "text") for t in self.app.nb.tabs()]
        self.assertEqual(names, ["brief 编辑", "大纲", "生成", "预览与历史", "模板"])

    def test_default_geometry_and_shortcuts(self):
        self.assertEqual(self.app.geometry().split("+")[0], "1180x760")
        for seq in ("<Control-s>", "<Control-o>", "<Control-n>", "<Control-Return>", "<F5>"):
            self.assertTrue(self.app.bind_all(seq), f"{seq} 未绑定快捷键")

    # ---------- 未保存流程 ----------
    def test_dirty_marker_and_saved_state(self):
        brief = os.path.join(self.tmp, "b.yaml")
        open(brief, "w", encoding="utf-8").write(SAMPLE)
        self.app.load_brief_into_editor(brief)
        self.app.update()
        self.assertNotIn("●", self.app.title())

        self.app.brief_text.insert("end", "\n# 备注\n")
        self.app.update()
        self.assertTrue(self.app._dirty)
        self.assertIn("●", self.app.title())

        self.assertTrue(self.app.save_brief())
        self.assertFalse(self.app._dirty)
        self.assertNotIn("●", self.app.title())
        self.assertIn("# 备注", open(brief, encoding="utf-8").read())

    def test_cancel_discard_keeps_current_text(self):
        self.load_text(SAMPLE)
        self.app._dirty = True
        with patch.object(wb.messagebox, "askyesnocancel", return_value=None) as ask:
            self.assertFalse(self.app._confirm_discard())
        ask.assert_called_once()
        self.assertIn("notes: keep me", self.app.brief_text.get("1.0", "end"))

    # ---------- 大纲 ----------
    def test_outline_rows_and_move(self):
        self.load_text(SAMPLE)
        self.app.refresh_outline()
        self.assertEqual([r[1] for r in self.outline_rows()], ["封面", "要点", "结论"])
        self.app.outline_tree.selection_set("1")
        self.app.move_outline(1)
        types = [s["type"] for s in __import__("yaml").safe_load(
            self.app.brief_text.get("1.0", "end"))["sections"]]
        self.assertEqual(types, ["points", "cover", "conclusion"])
        self.assertEqual(self.app.outline_tree.selection(), ("2",))

    def test_outline_insert_appends_and_selects_last(self):
        self.load_text(SAMPLE)
        self.app.refresh_outline()
        self.app.insert_outline("table")
        self.assertEqual(self.outline_rows()[-1][1], "表格")
        self.assertEqual(self.app.outline_tree.selection(), ("4",))
        self.assertTrue(self.app._dirty)

    def test_outline_remove_keeps_tail(self):
        self.load_text(SAMPLE)
        self.app.refresh_outline()
        with patch.object(wb.messagebox, "askyesno", return_value=True):
            self.app.outline_tree.selection_set("2")
            self.app.remove_outline()
        self.assertIn("notes: keep me", self.app.brief_text.get("1.0", "end"))
        self.assertEqual([r[1] for r in self.outline_rows()], ["封面", "结论"])

    def test_debounce_rebuilds_outline_after_edit(self):
        self.load_text(SAMPLE)
        self.app.refresh_outline()
        before = len(self.outline_rows())
        head, middle, tail = SAMPLE.partition("notes: keep me")
        four = head + "  - type: conclusion\n    heading: 追加\n    points:\n      - x\n\n" + middle + tail
        self.load_text(four)
        self.app.refresh_outline()
        self.assertEqual(len(self.outline_rows()), before + 1)

        # 删掉第 4 段：防抖到期前大纲保持 4 行，之后回落到 3 行。
        text = self.app.brief_text.get("1.0", "end-1c")
        start, end = wl.section_spans(text)[3]
        self.app.brief_text.delete(f"{start + 1}.0", f"{end}.end")
        self.app.update()
        self.assertEqual(len(self.outline_rows()), before + 1)
        time.sleep(0.6)
        self.app.update()
        self.assertEqual(len(self.outline_rows()), before)
        self.assertIsNone(self.app._outline_job)

    def test_broken_yaml_reports_and_recovers(self):
        self.load_text(SAMPLE)
        self.app.refresh_outline()
        self.assertEqual(len(self.outline_rows()), 3)

        self.app.brief_text.delete("1.0", "end")
        self.app.brief_text.insert("1.0", "meta: [broken\n")
        self.app.update()
        time.sleep(0.6)
        self.app.update()
        self.assertEqual(len(self.outline_rows()), 0)
        self.assertIn("大纲不可用", self.app.status_var.get())

        self.load_text(SAMPLE)
        self.app.refresh_outline()
        self.assertEqual(len(self.outline_rows()), 3)

    # ---------- 单页选择器 ----------
    def test_slide_choices_and_parse(self):
        self.app.load_brief_into_editor(DEMO_BRIEF)
        self.app.update()
        values = list(self.app.slide_combo["values"])
        self.assertEqual(len(values), 12)
        self.assertTrue(values[0].startswith("1. 封面"))

        self.app.slide_var.set(values[4])
        with patch.object(self.app, "do_generate") as generate:
            self.app.do_generate_slide()
        generate.assert_called_once_with(5)

    # ---------- 历史 ----------
    def test_history_label_status_zoom_and_delete(self):
        history = os.path.join(self.tmp, "history")
        output = os.path.join(self.tmp, "demo.pptx")
        wb.run_generate(DEMO_BRIEF, output, "lab", history_dir=history)

        self.app.history_var.set(history)
        self.app.refresh_history()
        label = self.app.version_list.get(0)
        self.assertIn("lab", label)
        self.assertIn("12 页", label)
        self.assertIn("完整生成", label)
        self.assertIn("完整", self.app.history_status_var.get())

        self.app.show_preview(self.app._versions[0]["preview"])
        width = self.app.preview_img.width()
        self.assertGreater(width, 0)
        self.app.zoom_preview(1)
        self.assertGreater(self.app.preview_img.width(), width)
        self.app.fit_preview()
        self.assertEqual(self.app.preview_img.width(), width)

        with patch.object(wb.messagebox, "askyesno", return_value=True):
            self.app.version_list.selection_set(0)
            self.app.delete_version()
        self.assertEqual(os.listdir(os.path.join(history, "versions")), [])

    def test_delete_rejects_unsafe_version_id(self):
        self.app.history_var.set(os.path.join(self.tmp, "history"))
        self.app._versions = [{"id": "../escape", "preview": None}]
        self.app.version_list.insert("end", "bad")
        self.app.version_list.selection_set(0)
        with patch.object(wb.messagebox, "showerror") as error, \
                patch.object(wb.messagebox, "askyesno", return_value=True):
            self.app.delete_version()
        error.assert_called_once()

    # ---------- 生成状态 ----------
    def test_busy_toggles_buttons_progress_and_abandon(self):
        self.app._set_busy(True)
        self.assertEqual([str(b["state"]) for b in self.app._generate_buttons], ["disabled", "disabled"])
        self.assertEqual(str(self.app.abandon_btn["state"]), "normal")
        self.assertEqual(self.app["cursor"], "watch")
        self.app._set_busy(False)
        self.assertEqual([str(b["state"]) for b in self.app._generate_buttons], ["normal", "normal"])
        self.assertEqual(str(self.app.abandon_btn["state"]), "disabled")

    def test_clean_brief_generates_without_dialog(self):
        """demo brief 无提醒时，全量生成不应弹任何确认框。"""
        self.app.load_brief_into_editor(DEMO_BRIEF)
        self.app.update()
        output = os.path.join(self.tmp, "out.pptx")
        self.app.output_var.set(output)
        self.app.history_var.set("")
        with patch.object(wb.messagebox, "askyesno", side_effect=AssertionError("不该弹确认框")):
            self.app._start_generation(None, output, "lab", None)
            deadline = time.time() + 90
            while self.app._generating and time.time() < deadline:
                self.app.update()
                time.sleep(0.05)
            self.app.update()
        self.assertFalse(self.app._generating)
        self.assertTrue(os.path.isfile(output))
        self.assertIn("成功", self.app.result_text.get("1.0", "end"))

    def test_abandon_generation_stops_waiting(self):
        self.app._generating = True
        self.app._gen_abandon = False
        self.app._set_busy(True)
        self.app.abandon_generation()
        self.assertFalse(self.app._generating)
        self.assertTrue(self.app._gen_abandon)
        self.assertEqual([str(b["state"]) for b in self.app._generate_buttons], ["normal", "normal"])

    # ---------- 状态持久化（子进程验证恢复） ----------
    def test_state_roundtrip_keeps_geometry_and_tab(self):
        self.app.nb.select(2)
        self.app.update()
        self.app.save_state()
        self.assertTrue(os.path.isfile(wb.STATE_PATH))

        probe = os.path.join(self.tmp, "probe.py")
        with open(probe, "w", encoding="utf-8") as fh:
            fh.write(
                "import os, sys\n"
                "sys.path.insert(0, os.path.join(os.environ['WB_PROJECT'], 'group-meeting-ppt', 'scripts'))\n"
                "import workbench as wb\n"
                "wb.STATE_PATH = os.environ['WB_STATE_PATH']\n"
                "print('STATE', wb.STATE_PATH)\n"
                "app = wb.Workbench()\n"
                "app.withdraw()\n"
                "app.update()\n"
                "print('TAB', app.nb.index(app.nb.select()))\n"
                "print('GEO', app.geometry().split('+')[0])\n"
                "app.destroy()\n")
        done = subprocess.run([sys.executable, probe], capture_output=True, text=True,
                              encoding="utf-8", cwd=os.path.dirname(PROJECT),
                              env=dict(os.environ, PYTHONIOENCODING="utf-8",
                                       WB_STATE_PATH=wb.STATE_PATH,
                                       WB_PROJECT=os.path.dirname(PROJECT)), timeout=120)
        self.assertEqual(done.returncode, 0, done.stderr)
        self.assertIn("TAB 2", done.stdout)
        self.assertIn("GEO 1180x760", done.stdout)

    def tearDown(self):
        self.reset_app()


if __name__ == "__main__":
    unittest.main()
