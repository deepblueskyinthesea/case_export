"""手術依頼箋の Excel を見張り、デスクトップ\\症例ファイル に 001.xls から連番で保存する（第 2 段階・半自動）。

使い方（院内の Windows 端末で）:
  1. main.exe を起動する（前回の NNN.xls を消してよいか聞かれるので Enter）
  2. 電子カルテで患者を順に右クリック →「発行プレビュー」する
  3. Excel が開くと自動で保存して Excel を閉じるので、次の患者に進む
  4. 終わったら Ctrl+C かウィンドウを閉じる

画面には保存したファイル名と件数だけを出し、患者名は出さない。
"""

from __future__ import annotations

import hashlib
import re
import sys
import time
import traceback
from pathlib import Path

TEMPLATE_PATH = r"C:\SSI\Excel\手術\手術依頼箋.xls"
DATA_SHEET = "データ"  # 電子カルテが値を書き込むシート
OUTPUT_FOLDER_NAME = "症例ファイル"
POLL_SECONDS = 0.5
STABLE_COUNT = 3  # データシートの内容がこの回数続けて同じなら書き込み完了とみなす
MAX_FAILURES = 10  # 続けて読めなかったらその Excel はあきらめる

_NUMBERED_RE = re.compile(r"\d{3}\.xls", re.IGNORECASE)

# Excel が取り込み中のときに返すエラー（少し待てば通る）
BUSY_HRESULTS = {-2147418111, -2147417846}  # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER


def numbered_files(folder):
    """フォルダ内の NNN.xls（数字 3 桁）だけを返す。"""
    return sorted(p for p in Path(folder).iterdir()
                  if p.is_file() and _NUMBERED_RE.fullmatch(p.name))


def file_name(number):
    return f"{number:03d}.xls"


def next_number(folder, start):
    """start 以上で、まだファイルがない番号。"""
    number = start
    while (Path(folder) / file_name(number)).exists():
        number += 1
    return number


def is_target(full_name):
    return str(full_name).lower() == TEMPLATE_PATH.lower()


def content_hash(values):
    """セルの値（Range.Value の入れ子のタプル）から内容の指紋を作る。"""
    return hashlib.sha256(repr(values).encode("utf-8")).hexdigest()


def update_stability(prev_hash, stable, new_hash, ready):
    """(新しい指紋, 続けて同じだった回数) を返す。Excel が応答できないときは数え直す。"""
    if not ready:
        return new_hash, 0
    if new_hash == prev_hash:
        return new_hash, stable + 1
    return new_hash, 1


def call(fn, tries=20, wait=0.25):
    """COM の呼び出し。Excel が取り込み中なら少し待ってやり直す。"""
    for i in range(tries):
        try:
            return fn()
        except Exception as e:  # noqa: BLE001
            if getattr(e, "hresult", None) in BUSY_HRESULTS and i < tries - 1:
                time.sleep(wait)
                continue
            raise


# ---------------------------------------------------------------- Windows 側


def desktop_dir():
    """デスクトップの場所を Windows に聞く（OneDrive に移していても正しく取れる）。"""
    import ctypes

    buf = ctypes.create_unicode_buffer(260)
    if ctypes.windll.shell32.SHGetFolderPathW(None, 0x0010, None, 0, buf) != 0:  # CSIDL_DESKTOPDIRECTORY
        return Path.home() / "Desktop"
    return Path(buf.value)


def prepare_folder(folder):
    folder.mkdir(exist_ok=True)
    old = numbered_files(folder)
    if not old:
        return
    print(f"前回のファイルが {len(old)} 件あります（{old[0].name} ～ {old[-1].name}）。")
    input("Enter で削除して開始します（やめるときは Ctrl+C）: ")
    for path in old:
        path.unlink()
    print(f"{len(old)} 件削除しました。")


def excel_windows():
    import win32gui

    hwnd = 0
    while True:
        hwnd = win32gui.FindWindowEx(0, hwnd, "XLMAIN", None)
        if not hwnd:
            return
        yield hwnd


def find_target(app):
    for i in range(1, call(lambda: app.Workbooks.Count) + 1):
        wb = call(lambda: app.Workbooks(i))
        if is_target(call(lambda: wb.FullName)):
            return wb
    return None


class Pending:
    """見つけたが、まだ書き込み完了を待っているブック。"""

    def __init__(self, app, wb):
        self.app = app
        self.wb = wb
        self.last_hash = None
        self.stable = 0
        self.failures = 0


class Watcher:
    def __init__(self, folder):
        self.folder = folder
        self.apps = {}      # Excel ウィンドウ -> Application（接続はなるべく 1 ウィンドウ 1 回）
        self.pending = {}   # Excel ウィンドウ -> Pending
        self.skipped = set()  # 保存や読み取りに失敗し、ブックが開いたままかもしれない Excel
        self.connect_failures = {}  # Excel ウィンドウ -> 続けて接続に失敗した回数
        self.saved = {}     # 内容の指紋 -> ファイル名
        self.number = 0
        self.count = 0

    def poll(self):
        from inspect_ui import excel_app_from_window

        for hwnd in list(excel_windows()):
            if hwnd in self.skipped or hwnd in self.pending:
                continue
            try:
                app = self.apps.get(hwnd)
                if app is None:
                    app = excel_app_from_window(hwnd)
                    if app is None:  # まだブックが開いていない
                        continue
                    self.apps[hwnd] = app
                wb = find_target(app)
            except Exception as e:  # noqa: BLE001 - Excel の起動中など。次の回に接続し直す
                self.apps.pop(hwnd, None)
                n = self.connect_failures[hwnd] = self.connect_failures.get(hwnd, 0) + 1
                if n == MAX_FAILURES:  # 起動中の一時的な失敗は黙って待ち、続くときだけ一度知らせる
                    print(f"※ Excel に接続できません（{type(e).__name__}）。"
                          "保存されなければ、この患者は手で保存してください。")
                continue
            self.connect_failures.pop(hwnd, None)
            if wb is not None:  # 手術依頼箋ではないブックには触らない
                self.pending[hwnd] = Pending(app, wb)

        for hwnd, p in list(self.pending.items()):
            self.check(hwnd, p)

    def done(self, hwnd, skip=False):
        self.pending.pop(hwnd, None)
        if skip:  # 同じブックを何度も保存しないように、この Excel はもう見ない
            self.skipped.add(hwnd)

    def check(self, hwnd, p):
        try:
            ready = bool(call(lambda: p.app.Ready))
            values = call(lambda: p.wb.Worksheets(DATA_SHEET).UsedRange.Value)
        except Exception:  # noqa: BLE001
            p.failures += 1
            if p.failures >= MAX_FAILURES:
                print("※ 開いた Excel を読み取れなかったので飛ばしました。この患者は手で保存してください。")
                self.done(hwnd, skip=True)
            return
        p.failures = 0
        p.last_hash, p.stable = update_stability(p.last_hash, p.stable, content_hash(values), ready)
        if p.stable >= STABLE_COUNT:
            self.save(hwnd, p)

    def save(self, hwnd, p):
        self.number = next_number(self.folder, self.number + 1)
        target = self.folder / file_name(self.number)
        try:
            call(lambda: p.wb.SaveCopyAs(str(target)))
            if not target.exists():
                raise OSError("保存したファイルが見つかりません")
        except Exception as e:  # noqa: BLE001
            print(f"※ {target.name} の保存に失敗しました（{type(e).__name__}）。この患者は手で保存してください。")
            self.done(hwnd, skip=True)
            return
        self.count += 1
        print(f"{target.name} を保存しました（{self.count} 件目）")
        same = self.saved.setdefault(p.last_hash, target.name)
        if same != target.name:
            print(f"  ※ {same} と内容が同じです。同じ患者を 2 回開いていないか確認してください。")
        try:
            call(lambda: p.wb.Close(False))  # 保存せずに閉じる（手で閉じるのと同じ）
        except Exception:  # noqa: BLE001
            print("  ※ ブックを閉じられませんでした。手で閉じてください。")
            self.done(hwnd, skip=True)
            return
        self.quit_if_empty(hwnd, p.app)
        self.done(hwnd)

    def quit_if_empty(self, hwnd, app):
        """ブックが残っていなければ Excel ごと終了する（手で × を押すのと同じ）。"""
        self.apps.pop(hwnd, None)  # 接続を持ったままだと終了を止めてしまうので手放す
        try:
            if call(lambda: app.Workbooks.Count) == 0:
                call(lambda: app.Quit())
        except Exception:  # noqa: BLE001
            print("  ※ Excel を終了できませんでした。手で閉じてください。")


def watch(folder):
    import pythoncom

    watcher = Watcher(folder)
    print("見張っています。電子カルテで患者を右クリック →「発行プレビュー」してください。")
    print("終わったら Ctrl+C で終了します。")
    try:
        while True:
            pythoncom.PumpWaitingMessages()
            watcher.poll()
            time.sleep(POLL_SECONDS)
    except KeyboardInterrupt:
        pass
    return watcher.count


def main():
    if sys.platform != "win32":
        print("このツールは Windows 専用です。")
        return 1
    # nuitka の exe では comtypes の型ライブラリ日時の比較を避ける（CLAUDE.md 参照）
    if "__compiled__" in globals():
        sys.frozen = True

    folder = desktop_dir() / OUTPUT_FOLDER_NAME
    print(f"保存先: {folder}")
    try:
        prepare_folder(folder)
    except KeyboardInterrupt:
        print("\n中止しました。")
        return 1
    except OSError as e:
        print(f"前回のファイルを削除できませんでした（{type(e).__name__}）。開いていないか確認してください。")
        return 1

    count = watch(folder)
    print(f"\n終了しました。保存した件数: {count} 件")
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
