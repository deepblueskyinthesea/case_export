"""電子カルテと Excel の画面の作りを調べ、伏せ字にしたレポートを書き出す（実機用・読み取りのみ）。

使い方（院内の Windows 端末で）:
    python inspect_ui.py              # 10 秒後に前面のウィンドウを調べる
    python inspect_ui.py --delay 20   # 待ち時間を変える
    python inspect_ui.py --allow 発行 # 伏せ字にしない文字列を足す（複数可）

カウントダウン中に、調べたい画面を前面に出す。次の場面で 1 回ずつ実行する。
  1. 「主術依頼箋出力」ボタンが見えている電子カルテの画面
  2. 主術予定表発行ダイアログ
  3. 患者を右クリックしてメニューを開いたところ
  4. 発行プレビューで Excel が開いたところ
  5. その Excel を手で閉じたあと（Excel が裏に残るかを見る）

レポートは inspect_report_<日時>.txt に出る。ボタン・メニュー・列見出し以外の文字
（ウィンドウのタイトル、一覧の行、入力欄など）は文字数だけにして伏せる。
クリックやキー入力は一切しない。
"""

from __future__ import annotations

import argparse
import datetime
import os
import platform
import re
import sys
import time
import traceback
from pathlib import Path

# 名前をそのまま出してよい UI 部品（ボタン名・メニュー名・列見出しなど）
SHOW_NAME_TYPES = {
    "Button", "SplitButton", "MenuItem", "Menu", "MenuBar", "TabItem", "Tab",
    "ToolBar", "CheckBox", "RadioButton", "Header", "HeaderItem", "Hyperlink",
}
# 一覧の行。最初の数件だけ詳しく出し、残りは件数だけにする
ROW_TYPES = {"DataItem", "ListItem", "TreeItem"}

DEFAULT_ALLOW = {
    "主術依頼箋出力", "主術予定表発行", "発行プレビュー",
    "主術依頼箋", "手術依頼箋", "主術依頼箋.xls", "手術依頼箋.xls",
    "Excel", "Microsoft Excel",
}
# 伏せ字にした文字列に含まれていたら、含まれていることだけ示す語
KEYWORDS = ("主術", "手術", "依頼箋", "予定表", "発行", "プレビュー", "Excel")

UIA_PATTERNS = (
    "Invoke", "Selection", "SelectionItem", "Grid", "Table", "Scroll", "ScrollItem",
    "ExpandCollapse", "Value", "VirtualizedItem", "ItemContainer", "LegacyIAccessible",
)

_IDENT_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_.\-]{0,63}")
_FLAG_RE = re.compile(r"[/-][A-Za-z][A-Za-z0-9]*")

RPC_E_CALL_REJECTED = -2147418111  # Excel が取り込み中


def mask(text, allow=frozenset(), show=False, limit=60):
    """文字列を伏せる。allow に完全一致するか show のときだけそのまま出す。"""
    if text is None or text == "":
        return '""'
    text = str(text)
    if show or text in allow:
        shown = text if len(text) <= limit else text[:limit] + "…"
        return '"' + shown.replace("\r", " ").replace("\n", " ") + '"'
    hits = [k for k in KEYWORDS if k in text]
    hint = f" 含む:{','.join(hits)}" if hits else ""
    return f"<伏せ字 {len(text)}文字{hint}>"


def safe_ident(text, allow=frozenset()):
    """AutomationId など。英字で始まる識別子らしいものだけそのまま出す。"""
    if text and _IDENT_RE.fullmatch(text):
        return text
    return mask(text, allow)


def mask_path(path, allow=frozenset()):
    """ファイルパス。C:\\SSI 配下のフォルダと allow のファイル名だけ出す。"""
    if not path:
        return '""'
    path = str(path)
    folder, sep, name = path.rpartition("\\")
    if not sep:
        folder, name = "", path
    if folder.upper().startswith("C:\\SSI"):
        shown_folder = folder
    else:
        shown_folder = mask(folder, allow) if folder else ""
    ext = os.path.splitext(name)[1]
    shown_name = name if name in allow else mask(name, allow) + (f"（拡張子 {ext}）" if ext else "")
    return f"{shown_folder}\\{shown_name}" if shown_folder else shown_name


def summarize_cmdline(cmd, allow=frozenset()):
    """コマンドライン。/automation などのスイッチは出し、ファイル名は伏せる。"""
    if not cmd:
        return "(取得できず)"
    tokens = re.findall(r'"[^"]*"|\S+', cmd)
    parts = []
    for token in tokens[1:]:
        token = token.strip('"')
        if _FLAG_RE.fullmatch(token):
            parts.append(token)
        else:
            parts.append(mask_path(token, allow))
    return " ".join(parts) if parts else "(引数なし)"


def plan_children(types, keep=3, threshold=5):
    """子要素のうち詳しく出す番号と、省略する行の数を返す。

    行（ROW_TYPES）が threshold 件を超えたら、最初の keep 件だけ出す。
    """
    rows = [i for i, t in enumerate(types) if t in ROW_TYPES]
    if len(rows) <= threshold:
        return list(range(len(types))), 0
    skipped = set(rows[keep:])
    return [i for i in range(len(types)) if i not in skipped], len(skipped)


class Report:
    def __init__(self):
        self.lines = []

    def section(self, title):
        self.lines += ["", "=" * 70, title, "=" * 70]

    def add(self, line=""):
        self.lines.append(line)

    def text(self):
        return "\n".join(self.lines) + "\n"


def attempt(fn, default="取得失敗"):
    """COM や UIA の呼び出し。Excel が取り込み中なら少し待ってやり直す。"""
    for _ in range(5):
        try:
            return fn()
        except Exception as e:  # noqa: BLE001 - 調査用なので何が起きても続ける
            if getattr(e, "hresult", None) == RPC_E_CALL_REJECTED:
                time.sleep(0.5)
                continue
            return f"{default}({type(e).__name__})"
    return f"{default}(取り込み中)"


def attempt_count(fn):
    """件数を取る。取れなければ 0。"""
    value = attempt(fn)
    return value if isinstance(value, int) else 0


# ---------------------------------------------------------------- Windows 側


def process_table():
    """pid -> (親 pid, exe 名, コマンドライン)"""
    import win32com.client

    table = {}
    try:
        wmi = win32com.client.GetObject("winmgmts:")
        query = "SELECT ProcessId, ParentProcessId, Name, CommandLine FROM Win32_Process"
        for p in wmi.ExecQuery(query):
            table[int(p.ProcessId)] = (int(p.ParentProcessId or 0), p.Name or "", p.CommandLine or "")
    except Exception as e:  # noqa: BLE001
        print(f"プロセス一覧を取得できませんでした: {type(e).__name__}")
    return table


def top_windows(visible_only=True):
    import win32gui
    import win32process

    result = []

    def callback(hwnd, _):
        if not visible_only or win32gui.IsWindowVisible(hwnd):
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            result.append((hwnd, pid))
        return True

    win32gui.EnumWindows(callback, None)
    return result


def fmt_rect(r):
    return f"({r.left},{r.top})-({r.right},{r.bottom})"


def environment_section(rep, fg_hwnd):
    import ctypes

    rep.section("環境")
    rep.add(f"日時: {datetime.datetime.now():%Y-%m-%d %H:%M:%S}")
    rep.add(f"OS: {platform.platform()}")
    rep.add(f"Python: {sys.version.split()[0]} ({'64' if sys.maxsize > 2**32 else '32'}bit)")
    try:
        import pywinauto
        rep.add(f"pywinauto: {pywinauto.__version__}")
    except Exception as e:  # noqa: BLE001
        rep.add(f"pywinauto: 読み込めず({type(e).__name__})")
    user32 = ctypes.windll.user32
    rep.add(f"画面: {user32.GetSystemMetrics(0)}x{user32.GetSystemMetrics(1)}"
            f"  モニター数: {user32.GetSystemMetrics(80)}")
    rep.add(f"DPI（システム）: {attempt(lambda: user32.GetDpiForSystem())}")
    rep.add(f"DPI（前面ウィンドウ）: {attempt(lambda: user32.GetDpiForWindow(fg_hwnd))}")


def windows_section(rep, procs, allow):
    import win32gui

    rep.section("表示中のウィンドウ一覧")
    for hwnd, pid in top_windows():
        title = win32gui.GetWindowText(hwnd)
        if not title and win32gui.GetClassName(hwnd) != "#32768":
            continue
        exe = procs.get(pid, (0, "?", ""))[1]
        rep.add(f"hwnd={hwnd} pid={pid} exe={exe} class={win32gui.GetClassName(hwnd)!r} "
                f"title={mask(title, allow)}")


def uia_patterns(info):
    from pywinauto import uia_defines

    found, extra = [], []
    for name in UIA_PATTERNS:
        try:
            iface = uia_defines.get_elem_interface(info.element, name)
        except Exception:  # noqa: BLE001 - その操作に対応していない
            continue
        found.append(name)
        if name == "Grid":
            extra.append(f"行数={attempt(lambda: iface.CurrentRowCount)} "
                         f"列数={attempt(lambda: iface.CurrentColumnCount)}")
        elif name == "Selection":
            extra.append(f"複数選択={attempt(lambda: bool(iface.CurrentCanSelectMultiple))}")
        elif name == "LegacyIAccessible":
            extra.append(f"役割={attempt(lambda: iface.CurrentRole)}")
    return found, extra


def describe_uia(info, allow):
    ct = attempt(lambda: info.control_type, "?")
    parts = [
        str(ct),
        f"class={attempt(lambda: info.class_name)!r}",
        f"name={mask(attempt(lambda: info.name, ''), allow, show=ct in SHOW_NAME_TYPES)}",
        f"id={safe_ident(attempt(lambda: info.automation_id, ''), allow)}",
        f"fw={attempt(lambda: info.framework_id)}",
        f"rect={attempt(lambda: fmt_rect(info.rectangle))}",
        f"enabled={attempt(lambda: info.enabled)}",
    ]
    handle = attempt(lambda: info.handle, 0)
    if handle:
        parts.append(f"hwnd={handle}")
    found, extra = uia_patterns(info)
    if found:
        parts.append("操作=" + ",".join(found))
    parts += extra
    return " ".join(parts)


def dump_uia(rep, hwnd, args, allow):
    from pywinauto.uia_element_info import UIAElementInfo

    count = 0

    def walk(info, depth):
        nonlocal count
        if count >= args.max_elements:
            return
        count += 1
        rep.add("  " * depth + describe_uia(info, allow))
        if depth >= args.depth:
            return
        try:
            children = info.children()
        except Exception as e:  # noqa: BLE001
            rep.add("  " * (depth + 1) + f"(子要素を取得できず: {type(e).__name__})")
            return
        types = [attempt(lambda c=c: c.control_type, "?") for c in children]
        shown, skipped = plan_children(types)
        for i in shown:
            walk(children[i], depth + 1)
        if skipped:
            rows = sum(t in ROW_TYPES for t in types)
            rep.add("  " * (depth + 1) + f"(行は全部で {rows} 件。{skipped} 件は省略)")

    walk(UIAElementInfo(hwnd), 0)
    if count >= args.max_elements:
        rep.add(f"(要素が {args.max_elements} 件を超えたので打ち切り)")


# 件数を返すメッセージ（クラス名の一部で判定）
COUNT_MESSAGES = (
    ("SysListView32", 0x1004),  # LVM_GETITEMCOUNT
    ("SysTreeView32", 0x1105),  # TVM_GETCOUNT
    ("ListBox", 0x018B),        # LB_GETCOUNT
    ("ComboBox", 0x0146),       # CB_GETCOUNT
)


def win32_item_count(hwnd, class_name):
    import win32con
    import win32gui

    for key, msg in COUNT_MESSAGES:
        if key.lower() in class_name.lower():
            return attempt(lambda: win32gui.SendMessageTimeout(
                hwnd, msg, 0, 0, win32con.SMTO_ABORTIFHUNG, 1000)[1])
    return None


def menu_items(hwnd, allow):
    """ポップアップメニュー（#32768）の項目名。"""
    import ctypes

    user32 = ctypes.windll.user32
    hmenu = user32.SendMessageW(hwnd, 0x01E1, 0, 0)  # MN_GETHMENU
    if not hmenu:
        return ["(メニューを取得できず)"]
    items = []
    for i in range(user32.GetMenuItemCount(hmenu)):
        buf = ctypes.create_unicode_buffer(256)
        user32.GetMenuStringW(hmenu, i, buf, 256, 0x400)  # MF_BYPOSITION
        items.append(mask(buf.value, allow, show=True) if buf.value else "(区切り線か画像)")
    return items


def dump_win32(rep, hwnd, args, allow):
    from pywinauto.win32_element_info import HwndElementInfo

    count = 0

    def walk(info, depth):
        nonlocal count
        if count >= args.max_elements:
            return
        count += 1
        cls = attempt(lambda: info.class_name, "?")
        text = attempt(lambda: info.name, "")
        line = (f"{cls!r} text={mask(text, allow, show='button' in str(cls).lower())} "
                f"ctrl_id={attempt(lambda: info.control_id)} "
                f"rect={attempt(lambda: fmt_rect(info.rectangle))} "
                f"enabled={attempt(lambda: info.enabled)} hwnd={info.handle}")
        n = win32_item_count(info.handle, str(cls))
        if n is not None:
            line += f" 件数={n}"
        rep.add("  " * depth + line)
        if cls == "#32768":
            for item in menu_items(info.handle, allow):
                rep.add("  " * (depth + 1) + f"メニュー項目: {item}")
        if depth >= args.depth:
            return
        try:
            children = info.children()
        except Exception as e:  # noqa: BLE001
            rep.add("  " * (depth + 1) + f"(子ウィンドウを取得できず: {type(e).__name__})")
            return
        for child in children:
            walk(child, depth + 1)

    walk(HwndElementInfo(hwnd), 0)


def target_section(rep, fg_hwnd, procs, args, allow):
    import win32gui
    import win32process

    _, fg_pid = win32process.GetWindowThreadProcessId(fg_hwnd)
    exe = procs.get(fg_pid, (0, "?", ""))[1]
    rep.section("前面のウィンドウ")
    rep.add(f"hwnd={fg_hwnd} pid={fg_pid} exe={exe} class={win32gui.GetClassName(fg_hwnd)!r}")
    parent = procs.get(fg_pid, (0, "", ""))[0]
    rep.add(f"親プロセス: {procs.get(parent, (0, '(終了済み)', ''))[1]} (pid={parent})")

    if fg_pid == os.getpid() or exe.lower() in {
            "windowsterminal.exe", "conhost.exe", "openconsole.exe",
            "cmd.exe", "powershell.exe", "pwsh.exe"}:
        rep.add("※ 前面がこのツールのコンソールのままでした。画面の中身は調べていません。")
        print("\n※ 前面がコンソールのままでした。カウントダウン中に調べたい画面をクリックしてください。")
        return
    if win32gui.GetClassName(fg_hwnd) == "XLMAIN":
        rep.add("前面は Excel なので、画面の中身は調べず下の「Excel」の項目にまとめます。")
        return

    # 同じプロセスの表示中のウィンドウ（ダイアログやメニューを含む）と、ポップアップメニュー
    targets = [h for h, pid in top_windows()
               if pid == fg_pid or win32gui.GetClassName(h) == "#32768"]
    for hwnd in targets:
        title = mask(win32gui.GetWindowText(hwnd), allow)
        if not args.no_uia:
            rep.section(f"UIA の部品構成: hwnd={hwnd} title={title}")
            try:
                dump_uia(rep, hwnd, args, allow)
            except Exception as e:  # noqa: BLE001
                rep.add(f"(UIA で調べられず: {type(e).__name__}: {e})")
        if not args.no_win32:
            rep.section(f"Win32 の部品構成: hwnd={hwnd} title={title}")
            try:
                dump_win32(rep, hwnd, args, allow)
            except Exception as e:  # noqa: BLE001
                rep.add(f"(Win32 で調べられず: {type(e).__name__}: {e})")


# ---------------------------------------------------------------- Excel 側

OBJID_NATIVEOM = 0xFFFFFFF0


def excel_app_from_window(xlmain):
    """XLMAIN ウィンドウから、その Excel の Application を得る（複数起動していても取れる）。"""
    import ctypes
    from ctypes import wintypes

    import pythoncom
    import win32com.client.dynamic
    import win32gui

    class GUID(ctypes.Structure):
        _fields_ = [("Data1", ctypes.c_ulong), ("Data2", ctypes.c_ushort),
                    ("Data3", ctypes.c_ushort), ("Data4", ctypes.c_ubyte * 8)]

    desk = win32gui.FindWindowEx(xlmain, 0, "XLDESK", None)
    book = win32gui.FindWindowEx(desk, 0, "EXCEL7", None) if desk else 0
    if not book:
        return None
    iid_idispatch = GUID(0x00020400, 0, 0, (ctypes.c_ubyte * 8)(0xC0, 0, 0, 0, 0, 0, 0, 0x46))
    func = ctypes.oledll.oleacc.AccessibleObjectFromWindow
    func.argtypes = [wintypes.HWND, wintypes.DWORD, ctypes.POINTER(GUID),
                     ctypes.POINTER(ctypes.c_void_p)]
    ptr = ctypes.c_void_p()
    func(book, OBJID_NATIVEOM, ctypes.byref(iid_idispatch), ctypes.byref(ptr))
    window = win32com.client.dynamic.Dispatch(
        pythoncom.ObjectFromAddress(ptr.value, pythoncom.IID_IDispatch))
    return window.Application


def describe_workbook(rep, i, wb, allow):
    full_name = attempt(lambda: wb.FullName, "")
    rep.add(f"  ブック{i}: {mask_path(full_name, allow)}")
    rep.add(f"    Saved={attempt(lambda: wb.Saved)} ReadOnly={attempt(lambda: wb.ReadOnly)} "
            f"FileFormat={attempt(lambda: wb.FileFormat)} "
            f"HasVBProject={attempt(lambda: wb.HasVBProject)} "
            f"シート数={attempt(lambda: wb.Worksheets.Count)}")
    rep.add(f"    外部リンク数={attempt(lambda: len(wb.LinkSources(1) or ()))} "  # xlExcelLinks
            f"接続数={attempt(lambda: wb.Connections.Count)} "
            f"名前定義数={attempt(lambda: wb.Names.Count)}")
    if "依頼箋" not in str(full_name):
        return
    # ひな形のブックだけ、シートと名前定義の中身（セル番地）まで出す
    for j in range(1, attempt_count(lambda: wb.Worksheets.Count) + 1):
        ws = wb.Worksheets(j)
        used = attempt(lambda: ws.UsedRange)
        rep.add(f"    シート{j}: name={mask(attempt(lambda: ws.Name, ''), allow, show=True)} "
                f"使用範囲={attempt(lambda: used.Address)} "
                f"数式あり={attempt(lambda: used.HasFormula)}")  # None は一部だけ数式
    names = wb.Names
    for j in range(1, min(attempt_count(lambda: names.Count), 50) + 1):
        n = names(j)
        rep.add(f"    名前定義: {mask(attempt(lambda: n.Name, ''), allow, show=True)} "
                f"= {attempt(lambda: n.RefersTo)}")


def excel_section(rep, procs, allow):
    import win32gui

    rep.section("Excel")
    windowed = {pid for _, pid in top_windows(visible_only=False)}
    excel_pids = [pid for pid, (_, name, _) in procs.items() if name.lower() == "excel.exe"]
    if not excel_pids:
        rep.add("EXCEL.EXE は動いていません。")
    for pid in excel_pids:
        ppid, _, cmd = procs[pid]
        parent = procs.get(ppid, (0, "(終了済み)", ""))[1]
        rep.add(f"EXCEL.EXE pid={pid} 親={parent}(pid={ppid}) "
                f"引数={summarize_cmdline(cmd, allow)} ウィンドウあり={pid in windowed}")

    seen = set()
    for hwnd, pid in top_windows(visible_only=False):
        if win32gui.GetClassName(hwnd) != "XLMAIN":
            continue
        rep.add("")
        rep.add(f"Excel ウィンドウ hwnd={hwnd} pid={pid} "
                f"表示={bool(win32gui.IsWindowVisible(hwnd))}")
        try:
            app = excel_app_from_window(hwnd)
        except Exception as e:  # noqa: BLE001
            rep.add(f"  COM で接続できず: {type(e).__name__}")
            continue
        if app is None:
            rep.add("  ブックが開いていないので COM で接続できず")
            continue
        app_hwnd = attempt(lambda: app.Hwnd)
        if app_hwnd in seen:
            continue
        seen.add(app_hwnd)
        rep.add(f"  Version={attempt(lambda: app.Version)} Visible={attempt(lambda: app.Visible)} "
                f"UserControl={attempt(lambda: app.UserControl)} "
                f"Interactive={attempt(lambda: app.Interactive)} "
                f"Ready={attempt(lambda: app.Ready)} "
                f"ブック数={attempt(lambda: app.Workbooks.Count)}")
        for i in range(1, attempt_count(lambda: app.Workbooks.Count) + 1):
            try:
                describe_workbook(rep, i, app.Workbooks(i), allow)
            except Exception as e:  # noqa: BLE001
                rep.add(f"  ブック{i}: 調べられず({type(e).__name__})")


# ---------------------------------------------------------------- 起動


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="電子カルテと Excel の画面の作りを調べる（読み取りのみ）")
    p.add_argument("--delay", type=int, default=10, help="調べ始めるまでの秒数（既定 10）")
    p.add_argument("--depth", type=int, default=15, help="部品の階層をたどる深さ（既定 15）")
    p.add_argument("--max-elements", type=int, default=3000, help="1 画面で出す部品の上限")
    p.add_argument("--out", default=".", help="レポートを書き出すフォルダ（既定 今のフォルダ）")
    p.add_argument("--allow", action="append", default=[],
                   help="伏せ字にしない文字列（完全一致。複数指定可）")
    p.add_argument("--no-uia", action="store_true", help="UIA で調べない")
    p.add_argument("--no-win32", action="store_true", help="Win32 で調べない")
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    if sys.platform != "win32":
        print("このツールは Windows 専用です。")
        return 1

    # nuitka の exe では、ビルド PC で作った comtypes の補助ファイルを
    # 型ライブラリの更新日時が違ってもそのまま使わせる（PyInstaller の exe と同じ扱い）
    if "__compiled__" in globals():
        sys.frozen = True
    import pywinauto  # noqa: F401 - 先に読み込んで DPI の扱いをそろえる
    import win32gui

    allow = frozenset(DEFAULT_ALLOW | set(args.allow))
    print("調べたい画面（電子カルテ・ダイアログ・右クリックメニュー・Excel）を前面に出してください。")
    print("マウスやキーボードの操作はしません。")
    for left in range(args.delay, 0, -1):
        print(f"\r  あと {left} 秒 ", end="", flush=True)
        time.sleep(1)
    print()

    fg_hwnd = win32gui.GetForegroundWindow()
    procs = process_table()
    rep = Report()
    rep.add("case-export 画面調査レポート（患者名などは伏せ字。共有前に中身を確認すること）")
    environment_section(rep, fg_hwnd)
    target_section(rep, fg_hwnd, procs, args, allow)
    excel_section(rep, procs, allow)
    windows_section(rep, procs, allow)

    out = Path(args.out) / f"inspect_report_{datetime.datetime.now():%Y%m%d_%H%M%S}.txt"
    out.write_text(rep.text(), encoding="utf-8-sig")
    print(f"レポートを書き出しました: {out.resolve()}")
    print("共有する前に、患者名などが含まれていないか中身を確認してください。")
    return 0


if __name__ == "__main__":
    # exe をダブルクリックで起動したとき、結果を読む前にウィンドウが閉じないようにする
    try:
        exit_code = main()
    except Exception:  # noqa: BLE001
        traceback.print_exc()
        exit_code = 1
    input("Enter キーで閉じます。")
    sys.exit(exit_code)
