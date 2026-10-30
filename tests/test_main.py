"""main の連番・削除対象・書き込み完了の判定（Windows なしで確かめられる部分）。ダミーデータだけを使う。"""

import contextlib
import io
import tempfile
import unittest
from pathlib import Path
from unittest import mock

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

    def Range(self, address):
        return type("Range", (), {"Value": self.book.cells.get(address)})()


class FakeWorkbook:
    def __init__(self, values, cells=None):
        self.values = values
        self.cells = cells or {}  # データシートのセル番地 -> 値
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


class SettingsTest(unittest.TestCase):
    GRID = (40, 400, 990, 700)  # 患者一覧の画面上の範囲（ダミー）

    def test_make_settings_from_first_row(self):
        s = main.make_settings((100, 420), self.GRID, 96)
        self.assertEqual(s, {"version": main.SETTINGS_VERSION, "dpi": 96, "x": 60, "y": 20})

    def test_make_settings_rejects_outside(self):
        with self.assertRaises(ValueError):
            main.make_settings((1000, 420), self.GRID, 96)

    def test_row_point_follows_grid(self):
        s = {"dpi": 96, "x": 60, "y": 20, "row_height": 20.5}
        self.assertEqual(main.row_point(s, self.GRID, 0), (100, 420))
        self.assertEqual(main.row_point(s, self.GRID, 2), (100, 461))
        moved = (140, 300, 1090, 600)  # ダイアログを動かしても、一覧の左上からの位置は同じ
        self.assertEqual(main.row_point(s, moved, 2), (200, 361))

    def test_row_point_stops_at_bottom(self):
        s = {"dpi": 96, "x": 60, "y": 20, "row_height": 20}
        self.assertIsNotNone(main.row_point(s, self.GRID, 13))  # y=680
        self.assertIsNone(main.row_point(s, self.GRID, 14))     # y=700（下端）

    def test_load_settings(self):
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / main.SETTINGS_NAME
            self.assertIsNone(main.load_settings(path))  # ない
            path.write_text("{壊れている", encoding="utf-8")
            self.assertIsNone(main.load_settings(path))
            path.write_text('{"dpi": 96, "x": 60, "y": 20, "row_height": 20.5}', encoding="utf-8")
            self.assertIsNone(main.load_settings(path))  # 行番号を数えて測っていた古い形式
            path.write_text('{"version": 2, "dpi": 96, "x": 60}', encoding="utf-8")
            self.assertIsNone(main.load_settings(path))  # 項目が足りない
            path.write_text('{"version": 2, "dpi": 96, "x": 60, "y": 20}', encoding="utf-8")
            self.assertEqual(main.load_settings(path)["y"], 20)


WHITE, BLUE, PINK = (255, 255, 255), (0, 120, 215), (255, 210, 210)
LINE, HEADER, DARK, BORDER, BLACK = (192, 192, 192), (240, 240, 240), (100, 100, 100), (64, 64, 64), (0, 0, 0)


def grid_pixels(colors, row_height=20, header=24, width=80, line=1, dark=60,
                scrollbar=0, blue_covers_line=False, dark_color=DARK, hscroll=0):
    """患者一覧の画面を模したダミーの色の並び。返すのは (pixels, 1 行目の上端の y)。"""
    def pixel_row(color, text=None):
        row = [color] * width
        row[40] = LINE  # 縦の区切り線
        if text:
            row[10:30] = [text] * 20  # 文字（少数）
        if scrollbar:
            row[width - 2 - scrollbar:width - 2] = [HEADER] * scrollbar
        row[:2] = row[-2:] = [BORDER] * 2  # 外枠
        return row

    pixels = [[BORDER] * width for _ in range(2)]
    pixels += [pixel_row(HEADER, BLACK if 5 < y < 15 else None) for y in range(header)]
    pixels.append(pixel_row(LINE))
    top = len(pixels)
    for i, color in enumerate(colors):
        lines = 0 if (i == 0 and blue_covers_line) else line
        text = WHITE if color == BLUE else BLACK
        pixels += [pixel_row(color, text if 5 < y < 14 else None) for y in range(row_height - lines)]
        pixels += [pixel_row(LINE) for _ in range(lines)]
    pixels += [pixel_row(dark_color) for _ in range(dark)]
    pixels += [pixel_row(HEADER) for _ in range(hscroll)]  # 横のスクロールバー
    pixels += [[BORDER] * width for _ in range(2)]
    return pixels, top


class AnalyzeGridTest(unittest.TestCase):
    def analyze(self, pixels, first_y):
        return main.analyze_grid(main.row_colors(pixels), first_y)

    def test_rows_and_height(self):
        pixels, top = grid_pixels([BLUE, WHITE, PINK, WHITE, WHITE, WHITE, WHITE])
        self.assertEqual(self.analyze(pixels, top + 10), (top, 20.0, 7))

    def test_ignores_scrollbar_and_border(self):
        pixels, top = grid_pixels([BLUE, WHITE, WHITE], scrollbar=15, row_height=23)
        self.assertEqual(self.analyze(pixels, top + 3), (top, 23.0, 3))

    def test_single_patient(self):
        pixels, top = grid_pixels([BLUE])
        self.assertEqual(self.analyze(pixels, top + 10), (top, 20.0, 1))

    def test_selection_covering_line(self):
        pixels, top = grid_pixels([BLUE, WHITE, WHITE, WHITE], blue_covers_line=True)
        self.assertEqual(self.analyze(pixels, top + 10), (top, 20.0, 4))

    def test_first_y_on_dense_text(self):
        pixels, top = grid_pixels([BLUE, WHITE, WHITE])
        for y in (top + 8, top + 28):  # 文字だらけの列（一番多い色が半分に届かない）
            pixels[y] = [BORDER] * 2 + [WHITE] * 26 + [BLACK] * 25 + [BLUE] * 25 + [BORDER] * 2
        self.assertIsNone(main.row_colors(pixels)[top + 8])
        self.assertEqual(self.analyze(pixels, top + 8), (top, 20.0, 3))

    def test_no_dark_area(self):
        pixels, top = grid_pixels([BLUE, WHITE, WHITE], dark=0)
        with self.assertRaises(ValueError):
            self.analyze(pixels, top + 10)

    def test_horizontal_scrollbar_below_dark(self):
        pixels, top = grid_pixels([BLUE, WHITE, PINK, WHITE], hscroll=17)
        self.assertEqual(self.analyze(pixels, top + 10), (top, 20.0, 4))

    def test_standard_gray(self):
        pixels, top = grid_pixels([BLUE, WHITE, WHITE], dark_color=(171, 171, 171), hscroll=17)
        self.assertEqual(self.analyze(pixels, top + 10), (top, 20.0, 3))

    def test_no_dark_area_above_scrollbar(self):
        pixels, top = grid_pixels([BLUE, WHITE, WHITE], dark=0, hscroll=17)
        with self.assertRaises(ValueError):
            self.analyze(pixels, top + 10)

    def test_describe_bands(self):
        pixels, top = grid_pixels([BLUE, WHITE], dark=10, hscroll=17)
        lines = main.describe_bands(main.row_colors(pixels))
        self.assertIn("高さ  10  (100, 100, 100)", lines[-3])
        self.assertIn("高さ  16  (240, 240, 240)", lines[-2])  # 外枠ぶん 1 列減る
        self.assertIn("読めない", lines[-1])

    def test_no_lines_between_rows(self):
        pixels, top = grid_pixels([BLUE, WHITE, WHITE, WHITE, WHITE], line=0)
        with self.assertRaises(ValueError):
            self.analyze(pixels, top + 10)

    def test_first_y_outside_rows(self):
        pixels, top = grid_pixels([BLUE, WHITE])
        with self.assertRaises(ValueError):
            self.analyze(pixels, 0)  # 外枠の上


class ExcludeTest(unittest.TestCase):
    def test_only_local_anesthesia_alone(self):
        self.assertTrue(main.is_excluded("局所麻酔", None))
        self.assertTrue(main.is_excluded("局所麻酔", ""))
        self.assertTrue(main.is_excluded(None, "局所麻酔"))
        self.assertFalse(main.is_excluded("麻酔科一任", None))
        self.assertFalse(main.is_excluded("全身麻酔", "局所麻酔"))
        self.assertFalse(main.is_excluded("局所麻酔（ダミー）", None))
        self.assertFalse(main.is_excluded(None, None))


class RepeatTest(unittest.TestCase):
    def run_book(self, w, hwnd, values, anesthesia="麻酔科一任"):
        app = FakeApp()
        wb = FakeWorkbook(values, {"B14": anesthesia})
        p = w.pending[hwnd] = main.Pending(app, wb)
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            for _ in range(main.STABLE_COUNT):
                w.check(hwnd, p)
        return wb, app

    def test_same_as_previous_ends_without_saving(self):
        with tempfile.TemporaryDirectory() as d:
            w = main.Watcher(Path(d), end_on_repeat=True)
            self.run_book(w, 1, (("ダミー1", 1),))
            self.run_book(w, 2, (("ダミー2", 2),))
            self.assertFalse(w.repeated)
            wb, app = self.run_book(w, 3, (("ダミー2", 2),))  # 最後の患者がもう一度開いた
            self.assertTrue(w.repeated)
            self.assertEqual(w.count, 2)
            self.assertFalse((Path(d) / "003.xls").exists())
            self.assertTrue(wb.closed)
            self.assertTrue(app.quit)
            self.assertEqual(w.finished, 3)
            self.assertEqual(w.last_done, 3)

    def test_local_anesthesia_skipped_and_numbers_packed(self):
        with tempfile.TemporaryDirectory() as d:
            w = main.Watcher(Path(d), end_on_repeat=True)
            self.run_book(w, 1, (("ダミー1", 1),))
            wb, app = self.run_book(w, 2, (("ダミー2", 2),), anesthesia="局所麻酔")
            self.run_book(w, 3, (("ダミー3", 3),))
            self.assertTrue(wb.closed)
            self.assertTrue(app.quit)
            self.assertEqual((w.count, w.excluded), (2, 1))
            self.assertEqual(sorted(x.name for x in Path(d).iterdir()), ["001.xls", "002.xls"])
            self.assertEqual(w.finished, 3)

    def test_end_detected_when_last_is_local(self):
        with tempfile.TemporaryDirectory() as d:
            w = main.Watcher(Path(d), end_on_repeat=True)
            self.run_book(w, 1, (("ダミー1", 1),))
            self.run_book(w, 2, (("ダミー2", 2),), anesthesia="局所麻酔")
            self.assertFalse(w.repeated)
            self.run_book(w, 3, (("ダミー2", 2),), anesthesia="局所麻酔")  # 最後の患者がもう一度開いた
            self.assertTrue(w.repeated)
            self.assertEqual((w.count, w.excluded), (1, 1))

    def test_semi_auto_excludes_too(self):
        with tempfile.TemporaryDirectory() as d:
            w = main.Watcher(Path(d))
            self.run_book(w, 1, (("ダミー", 1),), anesthesia="局所麻酔")
            self.assertEqual((w.count, w.excluded), (0, 1))

    def test_semi_auto_saves_repeat(self):
        with tempfile.TemporaryDirectory() as d:
            w = main.Watcher(Path(d))
            self.run_book(w, 1, (("ダミー", 1),))
            self.run_book(w, 2, (("ダミー", 1),))
            self.assertFalse(w.repeated)
            self.assertEqual(w.count, 2)


class PollTest(unittest.TestCase):
    def test_connect_failure_reported_once(self):
        def fail(hwnd):
            raise TypeError("ダミー")

        with tempfile.TemporaryDirectory() as d:
            w = main.Watcher(Path(d))
            out = io.StringIO()
            with mock.patch.object(main, "excel_windows", return_value=[1]), \
                    mock.patch("inspect_ui.excel_app_from_window", fail), \
                    contextlib.redirect_stdout(out):
                for _ in range(main.MAX_FAILURES * 2):
                    w.poll()
            self.assertEqual(out.getvalue().count("Excel に接続できません（TypeError）"), 1)
            self.assertNotIn("ダミー", out.getvalue())


class MainTest(unittest.TestCase):
    def test_refuses_non_windows(self):
        if main.sys.platform == "win32":
            self.skipTest("Windows では実際に見張りを始めてしまう")
        self.assertEqual(main.main(), 1)


if __name__ == "__main__":
    unittest.main()
