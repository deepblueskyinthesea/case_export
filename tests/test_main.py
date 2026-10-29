"""main の連番・削除対象・書き込み完了の判定（Windows なしで確かめられる部分）。ダミーデータだけを使う。"""

import contextlib
import io
import tempfile
import unittest
from pathlib import Path

import main


class NumberingTest(unittest.TestCase):
    def test_file_name(self):
        self.assertEqual(main.file_name(1), "001.xls")
        self.assertEqual(main.file_name(12), "012.xls")

    def test_next_number_skips_existing(self):
        with tempfile.TemporaryDirectory() as d:
            (Path(d) / "001.xls").write_bytes(b"")
            (Path(d) / "002.xls").write_bytes(b"")
            self.assertEqual(main.next_number(d, 1), 3)
            self.assertEqual(main.next_number(d, 5), 5)


class NumberedFilesTest(unittest.TestCase):
    def test_only_three_digit_xls(self):
        with tempfile.TemporaryDirectory() as d:
            names = ["001.xls", "012.XLS", "1.xls", "0001.xls", "abc.xls",
                     "001.xlsx", "ダミー患者.xls", "001.xls.bak"]
            for name in names:
                (Path(d) / name).write_bytes(b"")
            (Path(d) / "003.xls").mkdir()  # フォルダは消さない
            found = [p.name for p in main.numbered_files(d)]
            self.assertEqual(sorted(found), ["001.xls", "012.XLS"])


class TargetTest(unittest.TestCase):
    def test_is_target_ignores_case(self):
        self.assertTrue(main.is_target(r"C:\SSI\Excel\手術\手術依頼箋.xls"))
        self.assertTrue(main.is_target(r"c:\ssi\excel\手術\手術依頼箋.XLS"))
        self.assertFalse(main.is_target(r"C:\Users\dummy\Desktop\001.xls"))


class StabilityTest(unittest.TestCase):
    def test_counts_same_hash(self):
        self.assertEqual(main.update_stability("a", 1, "a", True), ("a", 2))
        self.assertEqual(main.update_stability("a", 2, "b", True), ("b", 1))

    def test_not_ready_resets(self):
        self.assertEqual(main.update_stability("a", 2, "a", False), ("a", 0))

    def test_hash_depends_on_values(self):
        self.assertEqual(main.content_hash((("ダミー", 1),)), main.content_hash((("ダミー", 1),)))
        self.assertNotEqual(main.content_hash((("ダミー", 1),)), main.content_hash((("ダミー", 2),)))


class FakeSheet:
    def __init__(self, book):
        self.book = book

    @property
    def UsedRange(self):
        return type("Range", (), {"Value": self.book.values})()


class FakeWorkbook:
    def __init__(self, values):
        self.values = values
        self.closed = False

    def Worksheets(self, name):
        assert name == main.DATA_SHEET
        return FakeSheet(self)

    def SaveCopyAs(self, path):
        Path(path).write_bytes(b"dummy")

    def Close(self, save_changes):
        assert save_changes is False
        self.closed = True


class FakeWorkbooks:
    def __init__(self, count):
        self.Count = count


class FakeApp:
    Ready = True

    def __init__(self, other_books=0):
        self.Workbooks = FakeWorkbooks(other_books)  # 閉じたあとに残るブックの数
        self.quit = False

    def Quit(self):
        self.quit = True


class WatcherCheckTest(unittest.TestCase):
    def run_checks(self, watcher, hwnd, pending, times):
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            for _ in range(times):
                if hwnd in watcher.pending:
                    watcher.check(hwnd, pending)
        return out.getvalue()

    def test_saves_after_stable_and_closes(self):
        with tempfile.TemporaryDirectory() as d:
            w = main.Watcher(Path(d))
            wb = FakeWorkbook((("ダミー", 1),))
            p = w.pending[1] = main.Pending(FakeApp(), wb)
            self.run_checks(w, 1, p, main.STABLE_COUNT - 1)
            self.assertFalse((Path(d) / "001.xls").exists())
            out = self.run_checks(w, 1, p, 1)
            self.assertTrue((Path(d) / "001.xls").exists())
            self.assertTrue(wb.closed)
            self.assertNotIn(1, w.pending)
            self.assertNotIn(1, w.skipped)
            self.assertIn("001.xls を保存しました", out)
            self.assertNotIn("ダミー", out)

    def test_quits_excel_when_no_books_left(self):
        with tempfile.TemporaryDirectory() as d:
            w = main.Watcher(Path(d))
            app = FakeApp()
            w.apps[1] = app
            p = w.pending[1] = main.Pending(app, FakeWorkbook((("ダミー", 1),)))
            self.run_checks(w, 1, p, main.STABLE_COUNT)
            self.assertTrue(app.quit)
            self.assertNotIn(1, w.apps)

    def test_keeps_excel_with_other_books(self):
        with tempfile.TemporaryDirectory() as d:
            w = main.Watcher(Path(d))
            app = FakeApp(other_books=1)
            p = w.pending[1] = main.Pending(app, FakeWorkbook((("ダミー", 1),)))
            self.run_checks(w, 1, p, main.STABLE_COUNT)
            self.assertTrue((Path(d) / "001.xls").exists())
            self.assertFalse(app.quit)

    def test_waits_while_values_change(self):
        with tempfile.TemporaryDirectory() as d:
            w = main.Watcher(Path(d))
            wb = FakeWorkbook((("ダミー", 0),))
            p = w.pending[1] = main.Pending(FakeApp(), wb)
            for i in range(main.STABLE_COUNT * 2):
                wb.values = (("ダミー", i),)
                w.check(1, p)
            self.assertFalse((Path(d) / "001.xls").exists())

    def test_duplicate_warns_but_saves(self):
        with tempfile.TemporaryDirectory() as d:
            w = main.Watcher(Path(d))
            for hwnd in (1, 2):
                p = w.pending[hwnd] = main.Pending(FakeApp(), FakeWorkbook((("ダミー", 1),)))
                out = self.run_checks(w, hwnd, p, main.STABLE_COUNT)
            self.assertTrue((Path(d) / "002.xls").exists())
            self.assertIn("001.xls と内容が同じです", out)
            self.assertEqual(w.count, 2)


class MainTest(unittest.TestCase):
    def test_refuses_non_windows(self):
        if main.sys.platform == "win32":
            self.skipTest("Windows では実際に見張りを始めてしまう")
        self.assertEqual(main.main(), 1)


if __name__ == "__main__":
    unittest.main()
