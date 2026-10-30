"""手術予定表発行の全患者の Excel を、デスクトップ\\症例ファイル に 001.xls から連番で保存する（第 3 段階・自動）。

使い方（院内の Windows 端末で、電子カルテを開いた状態で）:
  1. main.exe を起動する（前回の NNN.xls を消してよいか聞かれるので Enter）
  2. 手術予定表発行ダイアログを手で開く（「手術依頼箋出力」ボタン。先に開いておいてもよい）
  3. 初回（と表示の拡大率を変えたとき）は、1 行目の位置を測る（マウスを置いて Enter。settings.json に記録）
  4. Enter で開始すると、1 行目を選んで一覧の色から行の高さと人数を読み取り、
     1 行目から順に右クリック →「発行プレビュー」→ 保存して Excel を閉じるを人数分繰り返す。
     マウスを動かすと止まる。一覧を読み取れなければ半自動に切り替える。
ダイアログに患者一覧が見つからないときは、第 2 段階の半自動（Excel を見張って保存するだけ）に切り替える。

画面には保存したファイル名と件数だけを出し、患者名は出さない。
"""

from __future__ import annotations

import hashlib
import json
import re
import sys
import time
import traceback
import warnings
from pathlib import Path

TEMPLATE_PATH = r"C:\SSI\Excel\手術\手術依頼箋.xls"
DATA_SHEET = "データ"  # 電子カルテが値を書き込むシート
# 依頼箋シートの麻酔の欄は =IF(データ!B14="","",データ!B14)&"  "&IF(データ!C14="","",データ!C14)
ANESTHESIA_CELLS = ("B14", "C14")
EXCLUDED_ANESTHESIA = "局所麻酔"  # 麻酔の欄がちょうどこれなら保存しない（ほかの文字が付けば残す）
OUTPUT_FOLDER_NAME = "症例ファイル"
POLL_SECONDS = 0.5
STABLE_COUNT = 3  # データシートの内容がこの回数続けて同じなら書き込み完了とみなす
MAX_FAILURES = 10  # 続けて読めなかったらその Excel はあきらめる

# 電子カルテの画面（第 1 段階の調査レポートで確かめた名前）
DIALOG_CLASS = "ThunderRT6FormDC"
DIALOG_TITLE = "手術予定表発行"
GRID_CLASS = "AfxOleControl42"  # 患者一覧（中の行は部品として取れないので座標で右クリックする）
MENU_CLASS = "#32768"
MENU_ITEM = "発行プレビュー"

SETTINGS_NAME = "settings.json"  # 1 行目の位置（main.exe と同じフォルダ）
SETTINGS_VERSION = 2  # 1 は行番号を数えて測っていた形式（測り直す）
SETTINGS_KEYS = ("dpi", "x", "y")
MIN_ROW_HEIGHT = 5   # これより低い行は読み間違いとみなす
# 患者一覧の色の読み取り（横一列ごとの「一番多い色」を上から並べて、行・区切り線・下の空きを見分ける）
BORDER_INSET = 3      # 一覧の外枠として読み飛ばす幅（px）
COLOR_SHARE = 0.5     # 横一列のうちこの割合以上を占める色を、その列の色とみなす
MAX_LINE_HEIGHT = 3   # これ以下の高さの色の帯は区切り線とみなす
ROW_TOLERANCE = 2     # 行の始まりの位置のずれの許容（px）
DARK_LEVEL = 200      # 一覧の下の空き（ダークグレー）とみなす明るさの上限（0～255）
GRAY_SPREAD = 16      # 赤・緑・青の差がこれ以下の色を灰色とみなす
ROW_MARGIN = 3       # 一覧の下端からこれより近い位置は右クリックしない
MAX_ROWS = 30        # 念のための上限（一覧は最大 12 人くらい）
DIALOG_SETTLE = 2    # ダイアログが出てから一覧が埋まるまで待つ秒数
MENU_TIMEOUT = 5     # 右クリックしてからメニューが出るまで待つ秒数
EXCEL_TIMEOUT = 60   # 発行プレビューを押してから Excel を保存し終えるまで待つ秒数
CLOSE_TIMEOUT = 10   # 保存した Excel の画面が消えるまで待つ秒数

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


def is_excluded(b14, c14):
    """麻酔の欄（B14 と C14 を空白 2 つでつないだもの）が局所麻酔だけか。"""
    text = f"{'' if b14 is None else b14}  {'' if c14 is None else c14}"
    return text.strip() == EXCLUDED_ANESTHESIA


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


def make_settings(point, grid_rect, dpi):
    """1 行目に置いたマウスの位置（画面座標）を、患者一覧の左上からの位置にする。"""
    left, top, right, bottom = grid_rect
    x, y = point
    if not (left <= x < right and top <= y < bottom):
        raise ValueError("マウスが患者一覧の外にありました")
    return {"version": SETTINGS_VERSION, "dpi": dpi, "x": x - left, "y": y - top}


def row_colors(pixels):
    """横一列ごとの「一番多い色」。どの色も COLOR_SHARE に届かない列と外枠の列は None。

    pixels は [行][列] -> (r, g, b)。文字や縦の区切り線、外枠、スクロールバーは少数なので無視される。
    """
    from collections import Counter

    colors = []
    height = len(pixels)
    for y, line in enumerate(pixels):
        inner = line[BORDER_INSET:len(line) - BORDER_INSET]
        if y < BORDER_INSET or y >= height - BORDER_INSET or not inner:
            colors.append(None)
            continue
        color, n = Counter(inner).most_common(1)[0]
        colors.append(color if n >= len(inner) * COLOR_SHARE else None)
    return colors


def color_bands(colors, start, end):
    """colors[start:end] の、同じ色が続く帯 (色, 始まり, 終わり) の並び。"""
    bands = []
    y = start
    while y < end:
        s = y
        while y < end and colors[y] == colors[s]:
            y += 1
        bands.append((colors[s], s, y))
    return bands


def band_edge(colors, y, step, color):
    """y から step（-1 か 1）の向きに見た、color の帯の端の y。

    途中の読めない列（None。文字が多い列など）は、その先も同じ色なら帯に含める。
    """
    edge = i = y
    while 0 <= i + step < len(colors):
        i += step
        if colors[i] == color:
            edge = i
        elif colors[i] is not None:
            break
    return edge


def is_dark_gray(color):
    """一覧の下の空き（ダークグレー）の色か。患者の行の色（白・ピンク・青）は灰色ではない。"""
    return color is not None and max(color) - min(color) <= GRAY_SPREAD and sum(color) / 3 <= DARK_LEVEL


def describe_bands(colors):
    """読み取れなかったときに見せる、上からの色の帯の並び（y・高さ・色）。色の数値だけで患者情報は含まない。"""
    return [f"  y={s:4} 高さ {e - s:3}  {'読めない' if color is None else color}"
            for color, s, e in color_bands(colors, 0, len(colors))]


def analyze_grid(colors, first_y):
    """一覧の列ごとの色から (1 行目の上端, 行の高さ, 人数) を求める。読み取れなければ ValueError。

    first_y は 1 行目（選んで青くなった行）の中の y。1 行目の帯の上端から、区切り線の色で行を分け、
    下の空き（ダークグレー）の手前までを数える。空きは 1 行目より下で一番下にある、行 1 つ分以上の
    高さのダークグレーの帯（その下の横のスクロールバーなどは飛ばす）。行の始まりが等間隔に並ばなければ
    読み違いとみなす。
    """
    n = len(colors)
    y = next((y for y in sorted(range(n), key=lambda y: abs(y - first_y))
              if abs(y - first_y) <= MAX_LINE_HEIGHT and colors[y] is not None), None)
    if y is None:
        raise ValueError("1 行目の色を読み取れませんでした")
    first = colors[y]
    top = band_edge(colors, y, -1, first)
    first_end = band_edge(colors, y, 1, first) + 1

    dark = next((s for color, s, e in reversed(color_bands(colors, first_end, n))
                 if e - s >= MIN_ROW_HEIGHT and is_dark_gray(color)), None)
    if dark is None:
        raise ValueError("一覧の下のダークグレーが見つかりませんでした")
    dark = max(band_edge(colors, dark, -1, colors[dark]), first_end)  # 途中の読めない列の上も含める

    bands = color_bands(colors, first_end, dark)
    lines = [c for c, s, e in bands if c is not None and e - s <= MAX_LINE_HEIGHT]
    line_color = max(set(lines), key=lines.count) if lines else None
    rows = [(top, first_end)]
    start = first_end
    for color, s, e in bands:
        if color == line_color and e - s <= MAX_LINE_HEIGHT:
            rows.append((start, s))
            start = e
    rows.append((start, dark))
    rows = [r for r in rows if r[1] - r[0] > MAX_LINE_HEIGHT]  # 二重線の片側などの細い切れ端

    count = len(rows)
    height = (dark - top) / count
    if height < MIN_ROW_HEIGHT or count > MAX_ROWS:
        raise ValueError("行の高さを読み取れませんでした")
    for k, (s, _) in enumerate(rows):
        if abs(s - (top + k * height)) > ROW_TOLERANCE:
            raise ValueError("行の高さがそろっていませんでした")
    return top, height, count


def row_point(settings, grid_rect, index):
    """index 行目（0 始まり）を右クリックする画面座標。一覧の下端を越えるなら None。"""
    left, top, right, bottom = grid_rect
    x = left + settings["x"]
    y = top + round(settings["y"] + settings["row_height"] * index)
    if y >= bottom - ROW_MARGIN:
        return None
    return x, y


def load_settings(path):
    """記録した行の位置を読む。ないか壊れていれば None（測り直す）。"""
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    if not isinstance(data, dict) or data.get("version") != SETTINGS_VERSION:
        return None
    if not all(isinstance(data.get(k), (int, float)) and not isinstance(data.get(k), bool)
               for k in SETTINGS_KEYS):
        return None
    return data


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
    def __init__(self, folder, end_on_repeat=False):
        self.folder = folder
        self.end_on_repeat = end_on_repeat  # 自動のとき、直前と同じ内容が開いたら一覧の終わりとみなす
        self.apps = {}      # Excel ウィンドウ -> Application（接続はなるべく 1 ウィンドウ 1 回）
        self.pending = {}   # Excel ウィンドウ -> Pending
        self.skipped = set()  # 保存や読み取りに失敗し、ブックが開いたままかもしれない Excel
        self.connect_failures = {}  # Excel ウィンドウ -> 続けて接続に失敗した回数
        self.saved = {}     # 内容の指紋 -> ファイル名
        self.number = 0
        self.count = 0
        self.excluded = 0
        self.last_seen = None   # 直前に開いた（保存か除外した）内容の指紋
        self.repeated = False   # 直前と同じ内容が開いた（一覧の終わり）
        self.finished = 0       # 保存・飛ばし・終わりの判定のどれかで片付けたブックの数
        self.last_done = None   # 最後に片付けた Excel ウィンドウ

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
        self.finished += 1
        self.last_done = hwnd
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
        if p.stable < STABLE_COUNT:
            return
        if self.end_on_repeat and p.last_hash == self.last_seen:
            self.repeated = True  # 一覧の下の空いた行を右クリックすると、最後の患者がもう一度開く
            self.close(hwnd, p)
            return
        try:
            sheet = call(lambda: p.wb.Worksheets(DATA_SHEET))
            b14, c14 = (call(lambda a=a: sheet.Range(a).Value) for a in ANESTHESIA_CELLS)
        except Exception:  # noqa: BLE001
            p.failures += 1  # 次の回に読み直す
            return
        self.last_seen = p.last_hash
        if is_excluded(b14, c14):
            self.excluded += 1
            print(f"局所麻酔なので保存せずに閉じました（除外 {self.excluded} 件目）")
            self.close(hwnd, p)
        else:
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
        self.close(hwnd, p)

    def close(self, hwnd, p):
        """ブックを保存せずに閉じ、ほかにブックがなければ Excel も終了する。"""
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
    return watcher.count, watcher.excluded


# ---------------------------------------------------------------- 自動操作（第 3 段階）


def set_dpi_aware():
    """座標を実際の画面の点で扱う（拡大率が 100% 以外でもずれないように）。"""
    import ctypes

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
    except Exception:  # noqa: BLE001 - 古い Windows
        ctypes.windll.user32.SetProcessDPIAware()


def system_dpi():
    import ctypes

    try:
        return int(ctypes.windll.user32.GetDpiForSystem())
    except Exception:  # noqa: BLE001 - 古い Windows
        user32, gdi32 = ctypes.windll.user32, ctypes.windll.gdi32
        hdc = user32.GetDC(0)
        try:
            return int(gdi32.GetDeviceCaps(hdc, 88))  # LOGPIXELSX
        finally:
            user32.ReleaseDC(0, hdc)


def program_dir():
    """main.exe のあるフォルダ（設定ファイルを置く）。"""
    return Path(sys.argv[0]).resolve().parent


def find_dialog():
    import win32gui

    hwnd = 0
    while True:
        hwnd = win32gui.FindWindowEx(0, hwnd, DIALOG_CLASS, None)
        if not hwnd:
            return None
        if win32gui.IsWindowVisible(hwnd) and DIALOG_TITLE in win32gui.GetWindowText(hwnd):
            return hwnd


def find_grid(dialog):
    import win32gui

    found = []

    def visit(hwnd, _):
        if win32gui.GetClassName(hwnd) == GRID_CLASS and win32gui.IsWindowVisible(hwnd):
            found.append(hwnd)
        return True

    win32gui.EnumChildWindows(dialog, visit, None)
    return found[0] if len(found) == 1 else None


def wait_dialog():
    """手術予定表発行ダイアログが開くまで待つ（開くのは手で行う）。"""
    dialog = find_dialog()
    if dialog:
        print("手術予定表発行ダイアログが開いているので、そのまま使います。")
        return dialog
    print("手術予定表発行ダイアログを開いてください（開くと先に進みます。やめるときは Ctrl+C）…")
    while not dialog:
        time.sleep(POLL_SECONDS)
        dialog = find_dialog()
    time.sleep(DIALOG_SETTLE)
    return dialog


def cursor_pos():
    import win32api

    return tuple(win32api.GetCursorPos())


def measure(grid, dpi):
    import win32gui

    print("患者一覧の 1 行目の位置を測ります。クリックはせず、マウスを置いてからこの画面で Enter を押してください。")
    print("（このウィンドウが患者一覧に重なっていたら、ずらしてください）")
    while True:
        input("1 行目の患者の名前か日付の上にマウスを置いて Enter（チェックボックスは避けてください）: ")
        try:
            return make_settings(cursor_pos(), win32gui.GetWindowRect(grid), dpi)
        except ValueError as e:
            print(f"※ {e}。もう一度測ります。")


def ensure_settings(grid):
    path = program_dir() / SETTINGS_NAME
    dpi = system_dpi()
    settings = load_settings(path)
    if settings is not None and settings["dpi"] == dpi:
        print(f"1 行目の位置は {path.name} を使います（測り直すときはこのファイルを削除してください）。")
        return settings
    if settings is not None:
        print("表示の拡大率が前回測ったときと違うので、測り直します。")
    elif path.exists():
        print(f"{path.name} が古い形式か壊れているので、測り直します。")
    settings = measure(grid, dpi)
    try:
        path.write_text(json.dumps(settings, indent=2), encoding="utf-8")
        print(f"1 行目の位置を {path.name} に記録しました。")
    except OSError as e:
        print(f"※ 1 行目の位置を記録できませんでした（{type(e).__name__}）。次回も測ります。")
    return settings


def bring_to_front(dialog):
    try:
        from pywinauto.controls.hwndwrapper import HwndWrapper

        HwndWrapper(dialog).set_focus()
    except Exception:  # noqa: BLE001 - 前に出せなくても、隠れていなければ右クリックできる
        pass


def grid_visible_at(grid, point):
    """その点で一番上にあるのが患者一覧か（ほかの画面に隠れていないか）。"""
    import win32gui

    hwnd = win32gui.WindowFromPoint(point)
    return hwnd == grid or bool(win32gui.IsChild(grid, hwnd))


def find_menu(pid):
    import win32gui
    import win32process

    hwnd = 0
    while True:
        hwnd = win32gui.FindWindowEx(0, hwnd, MENU_CLASS, None)
        if not hwnd:
            return None
        if (win32gui.IsWindowVisible(hwnd)
                and win32process.GetWindowThreadProcessId(hwnd)[1] == pid):
            return hwnd


def choose_preview(pid):
    """右クリックで出たメニューの「発行プレビュー」を選ぶ。出なければ False。"""
    from pywinauto import Desktop, keyboard

    end = time.monotonic() + MENU_TIMEOUT
    menu = None
    while menu is None and time.monotonic() < end:
        menu = find_menu(pid)
        if menu is None:
            time.sleep(0.1)
    if menu is None:
        return False
    item = Desktop(backend="uia").window(handle=menu).child_window(
        title=MENU_ITEM, control_type="MenuItem")
    try:
        item.wait("exists", timeout=MENU_TIMEOUT)
    except Exception:  # noqa: BLE001
        keyboard.send_keys("{ESC}")  # 思っていたメニューと違う。閉じて止める
        return False
    try:
        item.invoke()
    except Exception:  # noqa: BLE001 - 部品として押せなければ、手と同じようにクリックする
        item.click_input()
    return True


def wait_finished(watcher, before):
    """開いた Excel を保存する（か、終わりと判定する）まで待つ。時間切れなら False。"""
    import pythoncom

    end = time.monotonic() + EXCEL_TIMEOUT
    while time.monotonic() < end:
        pythoncom.PumpWaitingMessages()
        watcher.poll()
        if watcher.finished > before:
            return True
        time.sleep(POLL_SECONDS)
    return False


def wait_closed(hwnd):
    import win32gui

    end = time.monotonic() + CLOSE_TIMEOUT
    while hwnd and win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd):
        if time.monotonic() >= end:
            return
        time.sleep(POLL_SECONDS)


def capture(rect):
    """画面の rect の範囲の色を [行][列] -> (r, g, b) で返す。メモリの中だけで使い、保存しない。"""
    import ctypes
    from ctypes import wintypes

    left, top, right, bottom = rect
    width, height = right - left, bottom - top
    user32, gdi32 = ctypes.windll.user32, ctypes.windll.gdi32
    user32.GetDC.restype = wintypes.HDC
    user32.GetDC.argtypes = [wintypes.HWND]
    user32.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
    gdi32.CreateCompatibleDC.restype = wintypes.HDC
    gdi32.CreateCompatibleDC.argtypes = [wintypes.HDC]
    gdi32.CreateCompatibleBitmap.restype = wintypes.HBITMAP
    gdi32.CreateCompatibleBitmap.argtypes = [wintypes.HDC, ctypes.c_int, ctypes.c_int]
    gdi32.SelectObject.restype = wintypes.HGDIOBJ
    gdi32.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
    gdi32.BitBlt.argtypes = [wintypes.HDC, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                             wintypes.HDC, ctypes.c_int, ctypes.c_int, wintypes.DWORD]
    gdi32.GetDIBits.argtypes = [wintypes.HDC, wintypes.HBITMAP, wintypes.UINT, wintypes.UINT,
                                ctypes.c_void_p, ctypes.c_void_p, wintypes.UINT]
    gdi32.DeleteObject.argtypes = [wintypes.HGDIOBJ]
    gdi32.DeleteDC.argtypes = [wintypes.HDC]

    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG), ("biHeight", wintypes.LONG),
                    ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD),
                    ("biCompression", wintypes.DWORD), ("biSizeImage", wintypes.DWORD),
                    ("biXPelsPerMeter", wintypes.LONG), ("biYPelsPerMeter", wintypes.LONG),
                    ("biClrUsed", wintypes.DWORD), ("biClrImportant", wintypes.DWORD)]

    screen = user32.GetDC(None)
    memory = gdi32.CreateCompatibleDC(screen)
    bitmap = gdi32.CreateCompatibleBitmap(screen, width, height)
    buf = ctypes.create_string_buffer(width * height * 4)
    try:
        old = gdi32.SelectObject(memory, bitmap)
        gdi32.BitBlt(memory, 0, 0, width, height, screen, left, top, 0x00CC0020)  # SRCCOPY
        gdi32.SelectObject(memory, old)  # GetDIBits はビットマップを DC から外して呼ぶ
        header = BITMAPINFOHEADER(ctypes.sizeof(BITMAPINFOHEADER), width, -height, 1, 32)  # 上の行から
        if gdi32.GetDIBits(memory, bitmap, 0, height, buf, ctypes.byref(header), 0) != height:
            raise ValueError("画面を読み取れませんでした")
    finally:
        gdi32.DeleteObject(bitmap)
        gdi32.DeleteDC(memory)
        user32.ReleaseDC(None, screen)
    data = buf.raw
    stride = width * 4
    return [[(data[i + 2], data[i + 1], data[i]) for i in range(y * stride, (y + 1) * stride, 4)]
            for y in range(height)]  # BGRA の並び


def read_rows(dialog, grid, settings):
    """1 行目を左クリックして選び、一覧の色から行の位置と人数を読み取る。(行の位置, 人数) を返す。

    行の位置は row_point に渡す形（1 行目の中央と行の高さ）。読み取れなければ ValueError。
    """
    import win32gui
    from pywinauto import mouse

    rect = win32gui.GetWindowRect(grid)
    left, top, right, bottom = rect
    first = (left + settings["x"], top + settings["y"])
    inset = BORDER_INSET + 1
    checks = [first, (left + inset, top + inset), (right - inset, top + inset),
              (left + inset, bottom - inset), (right - inset, bottom - inset)]
    if not all(grid_visible_at(grid, point) for point in checks):
        bring_to_front(dialog)
        time.sleep(0.5)
        if not all(grid_visible_at(grid, point) for point in checks):
            raise ValueError("患者一覧が別の画面に隠れています")
    mouse.click(coords=first)  # 手で左クリックしたときと同じく、1 行目が青くなる
    time.sleep(0.5)
    colors = row_colors(capture(rect))
    try:
        row_top, height, count = analyze_grid(colors, settings["y"])
    except ValueError:
        print(f"一覧の色の帯（上から。1 行目は y={settings['y']} のあたり）:")
        print("\n".join(describe_bands(colors)))
        raise
    layout = {"x": settings["x"], "y": row_top + height / 2, "row_height": height}
    return layout, count


def auto_run(folder, dialog, grid, settings):
    """患者一覧を 1 行目から順に右クリック →「発行プレビュー」して保存する。(保存, 除外) の件数を返す。

    一覧を読み取れなければ None を返す（半自動に切り替える）。
    """
    import win32gui
    import win32process
    from pywinauto import mouse

    bring_to_front(dialog)
    try:
        layout, total = read_rows(dialog, grid, settings)
    except ValueError as e:
        print(f"※ 患者一覧を読み取れませんでした（{e}）。")
        return None
    print(f"患者一覧: {total} 人分の行を見つけました。")
    watcher = Watcher(folder, end_on_repeat=True)
    pid = win32process.GetWindowThreadProcessId(dialog)[1]
    last_pos = None
    for index in range(total):
        row = index + 1
        point = row_point(layout, win32gui.GetWindowRect(grid), index)
        if point is None:
            print("患者一覧の下端まで処理しました。")
            break
        if last_pos is not None and cursor_pos() != last_pos:
            print("マウスが動いたので止めました。")
            break
        if not grid_visible_at(grid, point):
            bring_to_front(dialog)
            time.sleep(0.5)
            if not grid_visible_at(grid, point):
                print(f"※ {row} 行目が別の画面に隠れているので止めました。")
                break
        before, saved_before, excluded_before = watcher.finished, watcher.count, watcher.excluded
        mouse.right_click(coords=point)
        if not choose_preview(pid):
            print(f"※ {row} 行目で「{MENU_ITEM}」のメニューが出なかったので止めました。")
            break
        last_pos = cursor_pos()
        if not wait_finished(watcher, before):
            print(f"※ {row} 行目で、{EXCEL_TIMEOUT} 秒たっても Excel を保存できなかったので止めました。")
            break
        if watcher.repeated:  # 読み取った人数が多すぎた（念のための止め方）
            print(f"※ {row} 行目で直前と同じ患者が開いたので、保存せずに閉じて止めました。")
            break
        if watcher.count == saved_before and watcher.excluded == excluded_before:  # 保存に失敗した（理由は表示済み）
            print(f"※ {row} 行目で止めました。")
            break
        wait_closed(watcher.last_done)
    else:
        print(f"一覧の最後まで（{total} 人分）処理しました。")
    return watcher.count, watcher.excluded


def run(folder):
    """自動で全患者を処理する。患者一覧が見つからなければ半自動に切り替える。(保存, 除外) の件数を返す。"""
    dialog = wait_dialog()
    grid = find_grid(dialog)
    if grid is None:
        print("※ 手術予定表発行ダイアログの患者一覧が見つかりませんでした。半自動に切り替えます。")
        return watch(folder)
    settings = ensure_settings(grid)
    input("Enter で自動処理を始めます。終わるまでマウスとキーボードに触らないでください: ")
    result = auto_run(folder, dialog, grid, settings)
    if result is None:
        print("半自動に切り替えます。")
        return watch(folder)
    return result


def focus_console():
    """このウィンドウを前に戻す（終わったときの Enter が電子カルテのダイアログに入らないように）。"""
    try:
        import ctypes

        from pywinauto.controls.hwndwrapper import HwndWrapper

        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            HwndWrapper(hwnd).set_focus()
    except Exception:  # noqa: BLE001 - 戻せなければ、表示のとおり手でクリックしてもらう
        pass


def main():
    if sys.platform != "win32":
        print("このツールは Windows 専用です。")
        return 1
    # nuitka の exe では comtypes の型ライブラリ日時の比較を避ける（CLAUDE.md 参照）
    if "__compiled__" in globals():
        sys.frozen = True
    set_dpi_aware()
    # COM はこのスレッドで最初に初期化する（pythoncom は最初に読み込んだスレッドでだけ初期化する）。
    # 方式は第 2 段階で Excel につながっていた STA に合わせる。
    sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")  # pywinauto が coinit_flags について出す警告
        import pythoncom  # noqa: F401
        import pywinauto  # noqa: F401

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

    try:
        count, excluded = run(folder)
    except KeyboardInterrupt:
        print("\n中止しました。")
        return 1
    finally:
        focus_console()
    print(f"\n終了しました。保存 {count} 件、局所麻酔で除外 {excluded} 件"
          "（合計が一覧の人数と合っているか確かめてください）")
    return 0


if __name__ == "__main__":
    # exe をダブルクリックで起動したとき、結果を読む前にウィンドウが閉じないようにする
    try:
        exit_code = main()
    except Exception:  # noqa: BLE001
        traceback.print_exc()
        exit_code = 1
    input("この画面をクリックしてから Enter キーで閉じます。")
    sys.exit(exit_code)
