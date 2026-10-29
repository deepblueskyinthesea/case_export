"""inspect_ui の伏せ字まわり（Windows なしで確かめられる部分）。ダミーの文字列だけを使う。"""

import unittest

import inspect_ui as iu


class MaskTest(unittest.TestCase):
    def test_hides_text_but_keeps_length(self):
        self.assertEqual(iu.mask("ダミー患者"), "<伏せ字 5文字>")

    def test_shows_keyword_hint_only(self):
        result = iu.mask("手術予定表発行 - ダミー患者")
        self.assertNotIn("ダミー", result)
        self.assertIn("含む:手術,予定表,発行", result)

    def test_allow_exact_match_only(self):
        allow = frozenset({"発行プレビュー"})
        self.assertEqual(iu.mask("発行プレビュー", allow), '"発行プレビュー"')
        self.assertTrue(iu.mask("発行プレビュー ダミー", allow).startswith("<伏せ字"))

    def test_show_truncates_and_flattens_newlines(self):
        self.assertEqual(iu.mask("OK\nNG", show=True), '"OK NG"')
        self.assertEqual(iu.mask("あ" * 70, show=True, limit=5), '"あああああ…"')

    def test_empty(self):
        self.assertEqual(iu.mask(""), '""')
        self.assertEqual(iu.mask(None), '""')


class IdentTest(unittest.TestCase):
    def test_identifier_shown(self):
        self.assertEqual(iu.safe_ident("btnPrint_1"), "btnPrint_1")

    def test_digits_and_japanese_hidden(self):
        self.assertTrue(iu.safe_ident("0012345").startswith("<伏せ字"))
        self.assertTrue(iu.safe_ident("ダミー").startswith("<伏せ字"))


class PathTest(unittest.TestCase):
    def test_ssi_template_shown(self):
        allow = frozenset(iu.DEFAULT_ALLOW)
        self.assertEqual(iu.mask_path(r"C:\SSI\Excel\手術\手術依頼箋.xls", allow),
                         r"C:\SSI\Excel\手術\手術依頼箋.xls")

    def test_other_path_hidden_with_extension(self):
        result = iu.mask_path(r"C:\Users\dummy\Desktop\ダミー患者.xls")
        self.assertNotIn("dummy", result)
        self.assertNotIn("ダミー", result)
        self.assertIn("拡張子 .xls", result)

    def test_cmdline_keeps_switches(self):
        cmd = r'"C:\Program Files\Microsoft Office\EXCEL.EXE" /automation -Embedding'
        self.assertEqual(iu.summarize_cmdline(cmd), "/automation -Embedding")

    def test_cmdline_hides_file(self):
        cmd = r'"EXCEL.EXE" "C:\Users\dummy\ダミー.xls"'
        self.assertNotIn("ダミー", iu.summarize_cmdline(cmd))


class PlanChildrenTest(unittest.TestCase):
    def test_few_rows_all_shown(self):
        types = ["Header"] + ["DataItem"] * 5
        self.assertEqual(iu.plan_children(types), (list(range(6)), 0))

    def test_many_rows_trimmed(self):
        types = ["Header"] + ["DataItem"] * 10 + ["ScrollBar"]
        shown, skipped = iu.plan_children(types)
        self.assertEqual(shown, [0, 1, 2, 3, 11])
        self.assertEqual(skipped, 7)


class MainTest(unittest.TestCase):
    def test_refuses_non_windows(self):
        if iu.sys.platform == "win32":
            self.skipTest("Windows では実際に画面を調べてしまう")
        self.assertEqual(iu.main(["--delay", "0"]), 1)


if __name__ == "__main__":
    unittest.main()
