import os
import sys
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
import template_wizard as tw
import workbench as wb


class WorkbenchSaveTests(unittest.TestCase):
    def test_failed_save_stops_generation(self):
        obj = SimpleNamespace(brief_path="existing.yaml", save_brief=Mock(return_value=False))
        self.assertFalse(wb.Workbench._ensure_brief_saved(obj))

    def test_disk_error_returns_false(self):
        obj = SimpleNamespace(brief_path="existing.yaml", brief_text=Mock())
        obj.brief_text.get.return_value = "content"
        with patch("builtins.open", side_effect=OSError("disk full")), patch.object(wb.messagebox, "showerror"):
            self.assertFalse(wb.Workbench.save_brief(obj))

    def test_failed_save_as_restores_previous_path(self):
        obj = SimpleNamespace(brief_path="old.yaml", save_brief=Mock(return_value=False),
                              brief_path_var=Mock())
        with patch.object(wb.filedialog, "asksaveasfilename", return_value="new.yaml"):
            self.assertFalse(wb.Workbench.save_brief_as(obj))
        self.assertEqual(obj.brief_path, "old.yaml")
        obj.brief_path_var.set.assert_not_called()


class TemplatePathTests(unittest.TestCase):
    def test_names_must_be_plain_portable_filenames(self):
        directory = os.path.dirname(__file__)
        for name in ("", "default", "../outside", "..\\outside", "/tmp/outside",
                     "C:\\outside", "CON", "NUL.txt", "bad.", "a:b"):
            with self.subTest(name=name), self.assertRaises(ValueError):
                tw.template_output_path(directory, name)
        self.assertEqual(tw.template_output_path(directory, "实验室"),
                         os.path.join(os.path.realpath(directory), "实验室.yaml"))

    def test_existing_symlink_cannot_escape(self):
        directory = os.path.realpath(os.path.dirname(__file__))
        outside = os.path.dirname(directory)
        realpath = os.path.realpath
        def resolve(path):
            if path.endswith("linked.yaml"):
                return os.path.join(outside, "other.yaml")
            return realpath(path)
        with patch.object(tw.os.path, "realpath", side_effect=resolve):
            with self.assertRaises(ValueError):
                tw.template_output_path(directory, "linked")


if __name__ == "__main__":
    unittest.main()
