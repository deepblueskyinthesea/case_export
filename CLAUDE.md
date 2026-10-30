# CLAUDE.md

**case-export**：電子カルテの「手術予定表発行」から患者ごとの Excel を
`デスクトップ\症例ファイル\001.xls` から連番で保存する作業を自動化するツールの開発リポジトリ。
保存した Excel は麻酔記録システム AIMS-ν の患者情報タブ「Excel ファイルから読み込む」で使う。

もとの手作業（自動化の対象）：
1. 電子カルテの「手術依頼箋出力」ボタンを押す
2. 手術予定表発行ダイアログの患者一覧で、目的の患者を右クリック →「発行プレビュー」
3. その患者の Excel ファイルが Excel で開く
4. ファイル → 名前を付けて保存 → デスクトップ\症例ファイル に保存
5. これをその週の全患者分繰り返す

自動化の方式：手術予定表発行ダイアログを開くまで（1）は手で行う。その先の電子カルテは画面操作
（患者一覧は中の行を部品として取れないので座標で右クリックし、メニューは UI Automation で選ぶ）、
Excel は COM で電子カルテが開いたブックを SaveCopyAs する。院内端末で全患者分を自動で処理できる
ことを確かめた（第 3 段階）。

---

## Mandatory Rules for Claude Code
- Do NOT modify any files unless explicitly instructed.
- Do NOT refactor existing code unless clearly requested.
- Prefer minimal, localized changes over large improvements.
- Stability and existing behavior are more important than code cleanliness.

---

## File Modification Policy
- Only modify files explicitly listed by the user.
- If a file is not mentioned, do not touch it.
- Ask for confirmation before expanding the scope.

---

## Change Proposal Requirement

### 3-Step Approval Flow (MUST follow strictly)

**Step 1: Ask first — Do NOT jump to a detailed plan**
- When the user describes a task, first ask clarifying questions if needed.
- Do NOT immediately produce a detailed implementation plan or code changes.
- Your first response should be a brief summary of your understanding, NOT a full proposal.

**Step 2: Propose — Text only, NO tool calls**
- After the user confirms what they want, present your change proposal:
  - What will be changed
  - Why it is necessary
  - Potential risks or side effects
- This response MUST contain ONLY text. Do NOT call any tools (including AskUserQuestion) in the same response as the proposal.

**Step 3: Wait for explicit approval — Then implement**
- After presenting the proposal, wait for the user's next message.
- Only proceed with implementation after receiving a clear, non-empty approval (e.g., "OK", "進めて", "お願いします").
- If the user's response is empty, ambiguous, or unclear, ask again. Do NOT assume approval.

### Absolute Prohibitions
- NEVER combine a proposal and tool calls in the same response.
- NEVER begin implementation without explicit user approval in a separate message.
- NEVER treat an empty or auto-generated response as approval.
- NEVER produce a detailed implementation plan before understanding the user's request.

---

## 構成

| 場所 | 中身 |
|---|---|
| `main.py` | 本体（第 3 段階・自動）。起動時に前回の `NNN.xls` を削除し、手術予定表発行ダイアログが手で開かれるのを待つ（開いていればそのまま使う）。患者一覧を 1 行目から座標で右クリック →「発行プレビュー」し、開いた Excel の書き込み完了を待って `症例ファイル\NNN.xls` に SaveCopyAs し、ブックと Excel を閉じる。麻酔の欄（データシートの B14・C14）がちょうど「局所麻酔」の患者は保存せずに閉じ、番号を詰める。処理の始めに 1 行目を左クリックで選び、一覧の画面の色（横一列ごとに一番多い色）をメモリ上で読んで、区切り線から行の高さを、下の空き（ダークグレー）までの行数から人数を求め、その人数分だけ処理する（読み取れなければ半自動へ。直前に開いた患者と同じ内容が開いたら念のため止まる）。1 行目の位置は初回にマウスを置いて測り、exe と同じフォルダの `settings.json` に記録する（拡大率が変わったら測り直す）。ダイアログに患者一覧が見つからないときも第 2 段階の半自動（Excel を見張って保存するだけ）に切り替える。 |
| `inspect_ui.py` | 実機用の画面調査ツール（第 1 段階）。前面の画面の部品構成と Excel の状態を、患者情報を伏せたレポートに書き出す。読み取りのみ。 |
| `build.bat` | Windows 用のビルドスクリプト（`build.bat <スクリプト名>`。省略するとダブルクリックでも `main` をビルドする）。nuitka の standalone で `dist\<名前>.dist` を作る。 |
| `tests/` | Windows なしで確かめられる部分のテスト（`python -m unittest discover -s tests`）。 |
| `pyproject.toml` | uv プロジェクト定義。依存は pywin32・pywinauto（Windows のみ）。 |

構成が固まったらこの表を更新する。

---

## 実行環境

本番は電子カルテと Excel が入った院内の Windows 端末。

開発コンテナ（Linux）では電子カルテも Excel も操作できない。
電子カルテ・Excel の画面操作にかかわる部分は実機でしか確かめられないので、
コンテナで確認できた範囲と実機確認が必要な範囲を分けて報告する。

院内端末には exe（standalone のフォルダ）しか持ち込めない。onefile はウイルス対策ソフトに
引っかかるので使わない。ビルドはユーザーのデスクトップ PC（Windows、ネット接続あり、
uv あり。Python 3.10 は uv が用意する）で pull して `build.bat <スクリプト名>` を実行する。
コンテナでは Windows 用の exe を作れない。

nuitka の exe で pywinauto を使うスクリプトは、pywinauto を読み込む前に
`if "__compiled__" in globals(): sys.frozen = True` とする。しないと comtypes が補助ファイル
（ビルド PC で作った `comtypes.gen`）の型ライブラリの更新日時を実行 PC と比べ、
`ImportError: Typelib different than module` で止まる。

pywin32・pywinauto（comtypes）で COM を使うときは、メインの処理で最初に読み込んで初期化し、
別スレッドで COM を使わない。pythoncom は最初に読み込んだスレッドでだけ COM を初期化するので、
別スレッドが先に読み込むとメインの処理は初期化されないまま Excel への接続に失敗する
（スレッドが片付けた COM のつながりをメインで使うと exe ごと落ちることもある）。
`main.py` は `sys.coinit_flags = 2`（STA）にしてから pythoncom と pywinauto を読み込む。

---

## Python 環境

リポジトリルートが uv プロジェクト（Python 3.10 以上、`.python-version` は 3.10）。
`.venv` はまだ作っていない。

**`uv run` を使わない。** `uv run` は実行前に venv を `uv.lock` へ強制同期してしまう
（worktree 間で venv を共有すると、ほかの worktree の venv まで書き換わる）。
インタプリタを直接呼ぶこと。

```
.venv/bin/python <script>
```

---

## Git

- コミット前に `git branch --show-current` を確認する。コミット先は `master`。
  ただし Orca の worktree（`git rev-parse --git-dir` と `--git-common-dir` の結果が違う）では
  その作業ブランチにコミットしてよい。それ以外で `master` 以外にいたらコミットせず、
  ユーザーに報告して指示を待つ。
- worktree での作業が終わったら、次の順で master へ取り込んでプッシュする。
  1. 作業ブランチにコミットする
  2. `git -C /workspace/case_export pull --ff-only` でメインのチェックアウトの master を最新にする
  3. `git rebase master` で master の上に付け直す（他の作業が入っていたら検証をやり直す）
  4. `git -C /workspace/case_export merge --ff-only <作業ブランチ>` で取り込み、`git -C /workspace/case_export push` する
  5. 早送りできなければ（先に別の作業が入った）2 からやり直す

  付け直しで衝突したら、単純なものだけ直し、判断が要るものは止めて報告する。force push はしない。
  作業ブランチはリモートにプッシュしない。ローカルの作業ブランチは Orca でワークスペースを
  閉じたあとに削除する。メインのチェックアウトに同じファイルの未コミット変更があって
  取り込みに失敗したら、止めて報告する。
- 並列にする作業は、触るファイルが重ならないように分ける（同じファイルは取り込み時に衝突する）。
- worktree では、ユーザーから新しい依頼を受けたら、作業を始める前に `git rebase master` で
  最新の master の上に付け直す（ほかの並列作業が先に取り込まれていても古いコードで作業しないため）。
  未コミットの変更があるときは付け直さず、「master より○件遅れています」とだけ伝える。
  付け直しで衝突したら、止めて報告する。
- worktree での venv の扱い（メインの `.venv` へのシンボリックリンクで共有するか）は未定。
  共有する場合は、worktree で `uv add` などの venv 変更をしない。
- マージ済みで不要になったフィーチャーブランチは、削除（ローカルとリモート）を提案する。
  古いブランチに誤ってコミットするのを防ぐため。

---

## Security Rules
- Never request or output secrets, API keys, or credentials.
- Do not log or print personal data.
- Assume production-like constraints even in development.
- 実際の患者の Excel ファイルや電子カルテの画面キャプチャをリポジトリに入れない。
- テストにはダミーの患者データを使う。
- 患者名をログ・コメント・ファイル名の例に出さない。

---

## Cost Awareness
- Keep responses concise.
- Avoid repeating large code blocks unless necessary.
- Prefer explanation over full implementation when possible.
