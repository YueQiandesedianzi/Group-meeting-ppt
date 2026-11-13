# -*- coding: utf-8 -*-
"""历史快照回归测试；所有输入输出均位于独立的临时目录。"""
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest import mock
import zipfile

import yaml
from PIL import Image
from pptx import Presentation


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))
import generate_ppt as g  # noqa: E402
import workbench as wb  # noqa: E402


class HistoryTests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory(prefix="group-meeting-history-")
        self.root = Path(self.tempdir.name)
        self.addCleanup(self.tempdir.cleanup)

    def write_brief(self, path, sections):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(
            yaml.safe_dump(
                {"meta": {"title": "历史测试"}, "sections": sections},
                allow_unicode=True,
                sort_keys=False,
            ),
            encoding="utf-8",
        )
        return path

    def make_template(self, directory, name="archive", primary="#123456"):
        directory.mkdir(parents=True, exist_ok=True)
        path = directory / f"{name}.yaml"
        path.write_text(
            yaml.safe_dump(
                {
                    "name": name,
                    "slide": {"aspect": "16:9", "width_in": 13.333, "height_in": 7.5},
                    "colors": {
                        "primary": primary,
                        "accent": "#2E74B5",
                        "light": "#D6E4F0",
                        "text": "#333333",
                        "muted": "#808080",
                        "line": "#C9C9C9",
                        "background": "#FFFFFF",
                    },
                    "fonts": {"cjk": "微软雅黑", "latin": "Calibri"},
                },
                allow_unicode=True,
                sort_keys=False,
            ),
            encoding="utf-8",
        )
        return path

    def cli(self, brief, output, history=None, template=None, template_dir=None,
            slide=None, check=True):
        cmd = [sys.executable, "-B", str(SCRIPTS / "generate_ppt.py"),
               str(brief), "-o", str(output)]
        if history is not None:
            cmd += ["--history-dir", str(history)]
        if template is not None:
            cmd += ["--template", str(template)]
        if template_dir is not None:
            cmd += ["--template-dir", str(template_dir)]
        if slide is not None:
            cmd += ["--slide", str(slide)]
        env = os.environ.copy()
        env["MPLCONFIGDIR"] = str(self.root / "matplotlib")
        env["PYTHONDONTWRITEBYTECODE"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        result = subprocess.run(
            cmd, text=True, encoding="utf-8", capture_output=True, env=env)
        if check and result.returncode != 0:
            self.fail(f"CLI exited {result.returncode}\n{result.stdout}\n{result.stderr}")
        return result

    def direct_snapshot(self, brief_path, template_path, history, output,
                        generation_mode="full", slide=None, preview_effect=None):
        brief = g.load_yaml(str(brief_path))
        dependencies = g.collect_image_dependencies(brief, str(brief_path.parent))
        info = g.describe_inputs(str(brief_path), str(template_path), dependencies)
        theme = g.Theme(g.load_yaml(str(template_path)))
        effect = self.fake_preview if preview_effect is None else preview_effect
        Path(output).write_bytes(b"test generated output")
        with mock.patch.object(g, "render_preview", side_effect=effect):
            return Path(g.write_snapshot(
                str(brief_path), brief, theme, str(history), str(template_path),
                dependencies, info, output=str(output),
                generation_mode=generation_mode, slide=slide,
            ))

    @staticmethod
    def fake_preview(brief, theme, output):
        Path(output).write_bytes(b"test preview placeholder")

    def assert_cover_color(self, output, expected):
        prs = Presentation(str(output))
        for shape in prs.slides[0].shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if run.text == "历史测试":
                        self.assertEqual(str(run.font.color.rgb), expected.lstrip("#"))
                        return
        self.fail("封面标题不存在")

    def test_isolated_cli_and_workbench_rebuild_with_archived_inputs(self):
        original = self.root / "original"
        external_image = self.root / "external" / "red.png"
        external_image.parent.mkdir(parents=True)
        Image.new("RGB", (80, 60), "red").save(external_image)
        brief = self.write_brief(original / "brief.yaml", [
            {"type": "cover", "title": "历史测试"},
            {"type": "figure", "heading": "图", "image": "../external/red.png"},
        ])
        template_dir = self.root / "custom-templates"
        template = self.make_template(template_dir, primary="#123456")
        history = self.root / "history"
        output = self.root / "source.pptx"
        self.cli(brief, output, history, template.stem, template_dir)

        versions = list((history / "versions").iterdir())
        self.assertEqual(len(versions), 1)
        version = versions[0]
        self.assertTrue((version / "brief.yaml").is_file())
        self.assertTrue((version / "template.yaml").is_file())
        self.assertTrue((history / "previews" / version.name / "preview.png").is_file())
        manifest = json.loads((version / "manifest.json").read_text(encoding="utf-8"))
        self.assertEqual(manifest["template"]["name"], template.stem)
        self.assertTrue(all(set(entry) == {"section", "sha256"}
                            for entry in manifest["images"]))
        saved_brief = g.load_yaml(str(version / "brief.yaml"))
        self.assertFalse(os.path.isabs(saved_brief["sections"][1]["image"]))
        archived_image = Path(g.image_path(saved_brief["sections"][1]["image"], str(version)))
        self.assertEqual(archived_image.parent, history / "assets")

        # 显式参数仍可覆盖历史模板，但默认应使用归档模板。
        override = self.make_template(template_dir, "override", "#ABCDEF")
        overridden = self.root / "overridden.pptx"
        self.cli(version / "brief.yaml", overridden, template=override.stem,
                 template_dir=template_dir)
        self.assert_cover_color(overridden, "#ABCDEF")
        directory_only = self.root / "directory-only.pptx"
        self.cli(version / "brief.yaml", directory_only, template_dir=template_dir)
        self.assert_cover_color(directory_only, "#123456")
        workbench_history = self.root / "workbench-history"
        count, _ = wb.run_generate(
            str(brief), str(self.root / "workbench-source.pptx"), template.stem,
            history_dir=str(workbench_history), tdir=str(template_dir))
        self.assertEqual(count, 2)
        self.assertEqual(len(wb.list_versions(str(workbench_history))), 1)

        detached = self.root / "detached-history"
        shutil.copytree(history, detached)
        shutil.rmtree(original)
        shutil.rmtree(external_image.parent)
        shutil.rmtree(template_dir)
        rebuilt = self.root / "rebuilt-cli.pptx"
        self.cli(detached / "versions" / version.name / "brief.yaml", rebuilt)
        self.assertEqual(len(Presentation(str(rebuilt)).slides), 2)
        self.assert_cover_color(rebuilt, "#123456")
        with zipfile.ZipFile(rebuilt) as package:
            media = [name for name in package.namelist() if name.startswith("ppt/media/")]
            self.assertEqual(len(media), 1)
            embedded_hash = hashlib.sha256(package.read(media[0])).hexdigest()
        self.assertEqual(embedded_hash, manifest["images"][0]["sha256"])

        workbench_rebuilt = self.root / "rebuilt-workbench.pptx"
        count, template_name = wb.run_generate(
            str(detached / "versions" / version.name / "brief.yaml"),
            str(workbench_rebuilt),
        )
        self.assertEqual((count, template_name), (2, "archive"))
        self.assert_cover_color(workbench_rebuilt, "#123456")
        listed = wb.list_versions(str(detached))
        self.assertEqual(len(listed), 1)
        self.assertFalse(listed[0]["legacy"])
        self.assertEqual(listed[0]["template_name"], "archive")

    def test_content_dedup_and_image_changes_preserve_old_assets(self):
        brief = self.root / "project" / "brief.yaml"
        self.root.joinpath("project").mkdir()
        first_image = self.root / "project" / "a.png"
        second_image = self.root / "project" / "b.jpg"
        Image.new("RGB", (60, 40), "red").save(first_image)
        shutil.copyfile(first_image, second_image)
        first_hash = g.file_hash(str(first_image))
        brief = self.write_brief(brief, [
            {"type": "figure", "heading": "图1", "image": "a.png"},
            {"type": "figure", "heading": "图2", "image": "b.jpg"},
        ])
        template = self.make_template(self.root / "templates")
        history = self.root / "history"
        first = self.direct_snapshot(brief, template, history, self.root / "one.pptx")
        data = yaml.safe_load(brief.read_text(encoding="utf-8"))
        data["sections"][0]["caption"] = "只新增说明"
        brief.write_text(yaml.safe_dump(data, allow_unicode=True), encoding="utf-8")
        second = self.direct_snapshot(brief, template, history, self.root / "two.pptx")
        self.assertNotEqual(first, second)
        self.assertEqual(len(list((history / "assets").iterdir())), 1)

        Image.new("RGB", (60, 40), "blue").save(first_image)
        second_hash = g.file_hash(str(first_image))
        self.direct_snapshot(brief, template, history, self.root / "three.pptx")
        assets = {path.name for path in (history / "assets").iterdir()}
        self.assertEqual(assets, {first_hash, second_hash})
        old_saved = g.load_yaml(str(first))
        self.assertTrue(Path(g.image_path(
            old_saved["sections"][0]["image"], str(first.parent))).is_file())
        manifest = json.loads(first.with_name("manifest.json").read_text(encoding="utf-8"))
        self.assertEqual([entry["sha256"] for entry in manifest["images"]], [first_hash] * 2)

    def test_change_detection_and_legacy_marker(self):
        image = self.root / "project" / "figure.png"
        image.parent.mkdir()
        Image.new("RGB", (60, 40), "red").save(image)
        brief = self.write_brief(self.root / "project" / "brief.yaml", [
            {"type": "cover", "title": "历史测试"},
            {"type": "figure", "heading": "图", "image": "figure.png"},
        ])
        template = self.make_template(self.root / "templates", "changing")
        history = self.root / "history"
        output = self.root / "out.pptx"
        self.cli(brief, output, history, template.stem, template.parent)
        first_marker = (history / ".last_generated.json").read_bytes()
        result = self.cli(brief, output, history, template.stem, template.parent)
        self.assertIn("无变化", result.stdout)
        self.assertEqual(first_marker, (history / ".last_generated.json").read_bytes())

        self.make_template(self.root / "templates", "changing", "#654321")
        result = self.cli(brief, output, history, template.stem, template.parent)
        self.assertNotIn("无变化", result.stdout)
        Image.new("RGB", (60, 40), "green").save(image)
        result = self.cli(brief, output, history, template.stem, template.parent)
        self.assertNotIn("无变化", result.stdout)
        self.assert_cover_color(output, "#654321")

        other_output = self.root / "other-output.pptx"
        result = self.cli(brief, other_output, history, template.stem, template.parent)
        self.assertNotIn("无变化", result.stdout)
        legacy_history = self.root / "legacy-history"
        legacy_history.mkdir()
        legacy_history.joinpath(".last_generated_hash").write_text(
            g.file_hash(str(brief)), encoding="utf-8")
        other_output.write_bytes(b"not a PPTX")
        result = self.cli(brief, other_output, legacy_history, template.stem, template.parent)
        self.assertNotIn("无变化", result.stdout)
        self.assertEqual(len(Presentation(str(other_output)).slides), 2)

        info = g.describe_inputs(str(brief), str(template), [])
        fingerprint = g.generation_fingerprint(info, str(output))
        changed_generator = dict(info, generator_sha256="different")
        self.assertNotEqual(fingerprint, g.generation_fingerprint(changed_generator, str(output)))

    def test_partial_never_marks_full_and_snapshots_are_unique(self):
        brief = self.write_brief(self.root / "brief.yaml", [
            {"type": "points", "heading": "第一页", "points": ["旧一"]},
            {"type": "points", "heading": "第二页", "points": ["旧二"]},
        ])
        template = self.make_template(self.root / "templates")
        history = self.root / "history"
        output = self.root / "out.pptx"
        self.cli(brief, output, history, template.stem, template.parent)
        old_marker = (history / ".last_generated.json").read_bytes()
        data = yaml.safe_load(brief.read_text(encoding="utf-8"))
        data["sections"][0]["points"] = ["新一"]
        data["sections"][1]["points"] = ["新二"]
        brief.write_text(yaml.safe_dump(data, allow_unicode=True), encoding="utf-8")

        self.cli(brief, output, history, template.stem, template.parent, slide=1)
        partial_marker = json.loads((history / ".last_generated.json").read_text(encoding="utf-8"))
        self.assertEqual(partial_marker["status"], "partial")
        self.assertTrue(partial_marker["requires_full_rebuild"])
        self.assertNotEqual(old_marker, (history / ".last_generated.json").read_bytes())
        self.cli(brief, output, history, template.stem, template.parent, slide=1)
        second_partial_marker = json.loads(
            (history / ".last_generated.json").read_text(encoding="utf-8"))
        self.assertEqual(second_partial_marker["status"], "partial")
        self.assertTrue(second_partial_marker["requires_full_rebuild"])
        self.assertEqual(second_partial_marker["output_sha256"], g.file_hash(str(output)))
        self.assertEqual(
            {key: value for key, value in second_partial_marker.items() if key != "output_sha256"},
            {key: value for key, value in partial_marker.items() if key != "output_sha256"},
        )
        partial_snapshots = sorted((history / "versions").iterdir())[1:]
        self.assertEqual(len(partial_snapshots), 2)
        for version in partial_snapshots:
            manifest = json.loads(version.joinpath("manifest.json").read_text(encoding="utf-8"))
            self.assertEqual(manifest["generation_mode"], "slide")
            self.assertEqual(manifest["slide"], 1)

        result = self.cli(brief, output, history, template.stem, template.parent)
        self.assertNotIn("无变化", result.stdout)
        prs = Presentation(str(output))
        texts = [shape.text for shape in prs.slides[1].shapes if shape.has_text_frame]
        self.assertTrue(any("新二" in text for text in texts))
        self.assertNotEqual(old_marker, (history / ".last_generated.json").read_bytes())

    def test_restoring_old_brief_after_partial_still_requires_full_build(self):
        original_sections = [{"type": "points", "heading": "旧", "points": ["旧"]}]
        brief = self.write_brief(self.root / "brief.yaml", original_sections)
        template = self.make_template(self.root / "templates")
        history = self.root / "history"
        output = self.root / "out.pptx"
        self.cli(brief, output, history, template.stem, template.parent)
        old_brief_hash = g.file_hash(str(brief))
        data = yaml.safe_load(brief.read_text(encoding="utf-8"))
        data["sections"][0]["points"] = ["临时改动"]
        brief.write_text(yaml.safe_dump(data, allow_unicode=True), encoding="utf-8")
        self.cli(brief, output, history, template.stem, template.parent, slide=1)
        brief.write_text(yaml.safe_dump(
            {"meta": {"title": "历史测试"}, "sections": original_sections},
            allow_unicode=True, sort_keys=False), encoding="utf-8")
        self.assertEqual(g.file_hash(str(brief)), old_brief_hash)
        result = self.cli(brief, output, history, template.stem, template.parent)
        self.assertNotIn("无变化", result.stdout)
        prs = Presentation(str(output))
        texts = [shape.text for shape in prs.slides[0].shapes if shape.has_text_frame]
        self.assertTrue(any("旧" in text for text in texts))
        self.assertFalse(any("临时改动" in text for text in texts))

    def test_failed_history_or_marker_write_cannot_skip_stale_output(self):
        for failure_target in (
                "write_snapshot", "mark_history_pending", "mark_partial_generation"):
            with self.subTest(failure_target=failure_target):
                project = self.root / failure_target
                original_sections = [
                    {"type": "points", "heading": "一", "points": ["原始内容一"]},
                    {"type": "points", "heading": "二", "points": ["原始内容二"]},
                ]
                brief = self.write_brief(project / "brief.yaml", original_sections)
                original_bytes = brief.read_bytes()
                template = self.make_template(project / "templates")
                history = project / "history"
                output = project / "out.pptx"
                self.cli(brief, output, history, template.stem, template.parent)
                initial_marker = json.loads(
                    (history / ".last_generated.json").read_text(encoding="utf-8"))
                data = yaml.safe_load(brief.read_text(encoding="utf-8"))
                data["sections"][0]["points"] = ["失败后的新内容"]
                brief.write_text(yaml.safe_dump(data, allow_unicode=True), encoding="utf-8")
                args = [str(brief), "-o", str(output), "--history-dir", str(history),
                        "--template", template.stem, "--template-dir", str(template.parent)]
                is_partial = failure_target == "mark_partial_generation"
                if is_partial:
                    args += ["--slide", "1"]
                with mock.patch.object(
                        g, failure_target, side_effect=RuntimeError("模拟历史提交失败")):
                    self.assertEqual(g.main(args), 1)
                prs = Presentation(str(output))
                changed_text = "\n".join(
                    shape.text for shape in prs.slides[0].shapes if shape.has_text_frame)
                self.assertIn("失败后的新内容", changed_text)
                failed_marker = json.loads(
                    (history / ".last_generated.json").read_text(encoding="utf-8"))
                if failure_target == "write_snapshot":
                    self.assertEqual(failed_marker["status"], "history_pending")
                    self.assertEqual(failed_marker["last_full"], initial_marker)
                else:
                    self.assertEqual(failed_marker["status"], "full")
                    self.assertNotEqual(failed_marker["output_sha256"], g.file_hash(str(output)))

                brief.write_bytes(original_bytes)
                result = self.cli(brief, output, history, template.stem, template.parent)
                self.assertNotIn("无变化", result.stdout)
                prs = Presentation(str(output))
                restored_text = "\n".join(
                    shape.text for shape in prs.slides[0].shapes if shape.has_text_frame)
                self.assertIn("原始内容一", restored_text)
                self.assertNotIn("失败后的新内容", restored_text)

    def test_failed_snapshot_keeps_previous_marker_and_published_versions(self):
        brief = self.write_brief(self.root / "brief.yaml", [
            {"type": "points", "heading": "旧", "points": ["旧"]},
        ])
        template = self.make_template(self.root / "templates")
        history = self.root / "history"
        self.direct_snapshot(brief, template, history, self.root / "out.pptx")
        marker = history / ".last_generated.json"
        old_marker = marker.read_bytes()
        old_versions = {path.name for path in (history / "versions").iterdir()}
        old_previews = {path.name for path in (history / "previews").iterdir()}

        with self.assertRaisesRegex(RuntimeError, "模拟预览失败"):
            self.direct_snapshot(
                brief, template, history, self.root / "out.pptx",
                preview_effect=RuntimeError("模拟预览失败"))
        self.assertEqual(marker.read_bytes(), old_marker)
        self.assertEqual({path.name for path in (history / "versions").iterdir()}, old_versions)
        self.assertFalse(list(history.glob(".pending-*")))

        with mock.patch.object(g, "render_preview", side_effect=self.fake_preview):
            with mock.patch.object(g, "generation_marker",
                                   side_effect=OSError("模拟标记写入失败")):
                with self.assertRaisesRegex(OSError, "模拟标记写入失败"):
                    self.direct_snapshot(brief, template, history, self.root / "out.pptx")
        self.assertEqual(marker.read_bytes(), old_marker)
        self.assertEqual({path.name for path in (history / "versions").iterdir()}, old_versions)
        self.assertEqual({path.name for path in (history / "previews").iterdir()}, old_previews)
        self.assertFalse(list(history.glob(".pending-*")))

    def test_corrupted_snapshot_resources_are_rejected(self):
        image = self.root / "image.png"
        Image.new("RGB", (40, 30), "blue").save(image)
        brief = self.write_brief(self.root / "brief.yaml", [
            {"type": "cover", "title": "历史测试"},
            {"type": "figure", "heading": "图", "image": "image.png"},
        ])
        template = self.make_template(self.root / "templates")
        history = self.root / "history"
        self.cli(brief, self.root / "source.pptx", history, template.stem, template.parent)
        version = next((history / "versions").iterdir())
        snapshot_brief = version / "brief.yaml"
        manifest = json.loads((version / "manifest.json").read_text(encoding="utf-8"))
        good_template_hash = manifest["template"]["sha256"]
        version.joinpath("template.yaml").write_text(
            version.joinpath("template.yaml").read_text(encoding="utf-8") + "\n# changed\n",
            encoding="utf-8")
        result = self.cli(snapshot_brief, self.root / "bad-template.pptx", check=False)
        self.assertEqual(result.returncode, 2)
        self.assertIn("历史模板", result.stderr)
        self.assertIn("哈希", result.stderr)
        # 显式使用完整的新模板时，不需要读取已损坏的归档模板。
        result = self.cli(snapshot_brief, self.root / "override.pptx",
                          template="lab", template_dir=ROOT / "templates")
        self.assertTrue(result.returncode == 0, result.stderr)

        # 隔离图片损坏用例，先恢复模板，避免模板校验先行终止。
        shutil.copyfile(template, version / "template.yaml")
        Image.new("RGB", (40, 30), "green").save(
            history / "assets" / manifest["images"][0]["sha256"], format="PNG")
        result = self.cli(snapshot_brief, self.root / "bad-image.pptx", check=False)
        self.assertEqual(result.returncode, 2)
        self.assertIn("历史图片", result.stderr)
        self.assertIn("哈希", result.stderr)

        no_image_brief = self.write_brief(self.root / "plain" / "brief.yaml", [
            {"type": "points", "heading": "文字", "points": ["A"]},
        ])
        plain_history = self.root / "plain-history"
        self.cli(no_image_brief, self.root / "plain.pptx", plain_history,
                 template.stem, template.parent)
        plain_version = next((plain_history / "versions").iterdir())
        plain_snapshot = plain_version / "brief.yaml"
        self.assertEqual(g.validate_snapshot_inputs(
            str(plain_snapshot), g.load_yaml(str(plain_snapshot)),
            str(plain_version / "template.yaml")), None)
        # 该正常快照使用与测试早期损坏副本不同的独立模板。
        current_template_hash = json.loads(
            plain_version.joinpath("manifest.json").read_text(encoding="utf-8")
        )["template"]["sha256"]
        self.assertEqual(current_template_hash, g.file_hash(str(template)))
        self.assertEqual(current_template_hash, good_template_hash)

    def test_legacy_listing_and_explicit_template_rebuild(self):
        history = self.root / "history"
        versions = history / "versions"
        versions.mkdir(parents=True)
        preview = history / "previews" / "old" / "preview.png"
        preview.parent.mkdir(parents=True)
        preview.write_bytes(b"old preview")
        image = self.root / "external.png"
        Image.new("RGB", (30, 20), "blue").save(image)
        legacy = self.write_brief(versions / "old.yaml", [
            {"type": "figure", "heading": "旧图", "image": "../../external.png"},
        ])
        listed = wb.list_versions(str(history))
        self.assertEqual(len(listed), 1)
        self.assertTrue(listed[0]["legacy"])
        self.assertEqual(listed[0]["preview"], str(preview))
        template_dir = self.root / "templates"
        template = self.make_template(template_dir, "current", "#010203")
        output = self.root / "legacy-rebuilt.pptx"
        count, name = wb.run_generate(
            str(legacy), str(output), template.stem, tdir=str(template_dir))
        self.assertEqual((count, name), (1, "current"))
        image.unlink()
        with self.assertRaisesRegex(ValueError, "图片不存在"):
            wb.run_generate(str(legacy), str(self.root / "missing.pptx"),
                            template.stem, tdir=str(template_dir))

    def test_demo_all_layouts_and_validation_still_work(self):
        output = self.root / "demo.pptx"
        self.cli(ROOT / "assets" / "demo_brief.yaml", output, self.root / "demo-history")
        self.assertEqual(len(Presentation(str(output)).slides), 12)
        all_brief = self.write_brief(self.root / "all.yaml", [
            {"type": "cover", "title": "历史测试"},
            {"type": "agenda", "items": ["一", "二"]},
            {"type": "section_header", "heading": "章节"},
            {"type": "points", "heading": "要点", "points": ["A", "B"]},
            {"type": "figure", "heading": "图片",
             "image": str(ROOT / "assets" / "demo_figure.png")},
            {"type": "metrics", "heading": "指标",
             "metrics": [{"label": "准确率", "value": "91%"}]},
            {"type": "table", "heading": "表格",
             "columns": ["方法", "结果"], "rows": [["A", "91%"]]},
            {"type": "timeline", "heading": "路线",
             "steps": [{"time": "Q1", "title": "实验"}]},
            {"type": "conclusion", "heading": "结论", "points": ["完成"]},
            {"type": "next_steps", "heading": "计划", "items": [{"task": "复查"}]},
        ])
        all_output = self.root / "all.pptx"
        self.cli(all_brief, all_output)
        self.assertEqual(len(Presentation(str(all_output)).slides), 10)

        invalid = yaml.safe_load(all_brief.read_text(encoding="utf-8"))
        invalid["sections"][6]["rows"] = [["A", "91%", "extra"]]
        all_brief.write_text(yaml.safe_dump(invalid, allow_unicode=True), encoding="utf-8")
        result = self.cli(all_brief, self.root / "bad.pptx", check=False)
        self.assertEqual(result.returncode, 2)
        self.assertIn("长度", result.stderr)
        invalid["sections"][6]["rows"] = [["A", "91%"]]
        invalid["sections"][4]["image"] = "missing.png"
        all_brief.write_text(yaml.safe_dump(invalid, allow_unicode=True), encoding="utf-8")
        result = self.cli(all_brief, self.root / "missing.pptx", check=False)
        self.assertEqual(result.returncode, 2)
        self.assertIn("图片不存在", result.stderr)


if __name__ == "__main__":
    unittest.main()
