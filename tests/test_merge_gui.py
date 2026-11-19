from pathlib import Path
from tempfile import TemporaryDirectory
import time
import unittest

try:
    import tkinter as tk
except ImportError as error:
    raise unittest.SkipTest(f"Tkinter unavailable: {error}") from error

from merge_gui import MergeWindow, validate_paths
from test_merge_regressions import paragraph, write_docx


class MergeWindowTests(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.folder = Path(directory.name) / "中文 示例"
        self.folder.mkdir()
        self.copies = self.folder / "部门副本"
        self.copies.mkdir()
        self.base = self.folder / "原稿.docx"
        self.output = self.folder / "合稿.docx"
        write_docx(self.base, [paragraph("Original report")])

        try:
            self.root = tk.Tk()
        except tk.TclError as error:
            self.skipTest(f"Tk display unavailable: {error}")
        self.root.withdraw()
        self.addCleanup(self.root.destroy)
        self.ui = MergeWindow(self.root)
        self.ui.original.set(str(self.base))
        self.ui.copies.set(str(self.copies))
        self.ui.output.set(str(self.output))

    def wait_for_merge(self):
        deadline = time.monotonic() + 20
        while self.ui.running and time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.01)
        self.assertFalse(self.ui.running, "GUI merge timed out")

    def assert_actions(self, enabled):
        self.assertTrue(all(button.instate(["!disabled" if enabled else "disabled"])
                            for button in self.ui.actions))

    def test_empty_folder_stops_before_launch_and_input_overwrite_is_rejected(self):
        self.ui.start_button.invoke()
        self.assertFalse(self.ui.running)
        self.assertIn("没有可合并", self.ui.status.cget("text"))
        self.assert_actions(False)

        copy = self.copies / "副本.docx"
        write_docx(copy, [paragraph("Revised report")])
        with self.assertRaisesRegex(ValueError, "覆盖"):
            validate_paths(str(self.base), str(self.copies), str(copy))

    def test_gui_generates_and_enables_all_three_results(self):
        write_docx(self.copies / "副本.docx", [paragraph("Revised report")])
        self.ui.start_button.invoke()
        self.assertTrue(self.ui.running)
        self.assertTrue(self.ui.start_button.instate(["disabled"]))
        self.wait_for_merge()

        self.assertEqual(self.ui.status.cget("text"), "已生成合稿")
        self.assert_actions(True)
        self.assertIn("处理完成", self.ui.console.get("1.0", "end"))
        self.assertTrue(self.output.is_file())
        self.assertTrue(self.output.with_name("合稿_conflicts.txt").is_file())
        self.assertTrue(self.output.with_name("合稿_merged_changes_log.txt").is_file())

    def test_gui_reports_corrupt_copy_and_reenables_retry(self):
        (self.copies / "有损副本.docx").write_bytes(b"invalid")
        self.ui.start_button.invoke()
        self.wait_for_merge()

        self.assertIn("合并失败", self.ui.status.cget("text"))
        self.assertIn("有损副本.docx", self.ui.console.get("1.0", "end"))
        self.assertTrue(self.ui.start_button.instate(["!disabled"]))
        self.assert_actions(False)
        self.assertFalse(self.output.exists())


if __name__ == "__main__":
    unittest.main()
