# CLAUDE.md

麻酔記録システム **AIMS-ν（AIMS-nu）** の開発リポジトリ。

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
| `AIMS_nu/` | アプリのサブプロジェクト。**作業の大半はここ。この配下を触る前に [AIMS_nu/CLAUDE.md](AIMS_nu/CLAUDE.md) の地雷領域を読むこと。** |
| `AIMS_nu/AIMS_nu/` | アプリ本体のソース。`build.bat` 等の配布用スクリプトもここ。 |
| `scripts/` | コンテナ再作成時の環境復旧スクリプト。`/workspace` 直下のシンボリックリンクから呼ばれる。 |
| `AIファイル/` | 素材。 |

---

## Python 環境

仮想環境は **`AIMS_nu/.venv`（Python 3.10、PySide6 6.11.1）の1つだけ**。検証はこれで行う。

リポジトリルートは uv プロジェクトではないので、ルートで `uv sync` を打たないこと。
かつてルートにも Python 3.12 の venv があったが、中身が AIMS_nu 側の真部分集合で
取り違えの元にしかならなかったため削除した。

**`uv run` を使わない。** ユーザーの環境では `uv run` が venv を `uv.lock` へ強制同期してしまう。
インタプリタを直接呼ぶこと。

```
AIMS_nu/.venv/bin/python <script>
QT_QPA_PLATFORM=offscreen AIMS_nu/.venv/bin/python <script>   # GUI なしの検証
```

---

## Git

- コミット前に `git branch --show-current` を確認する。コミット先は `master`。
  ただし Orca の worktree（`git rev-parse --git-dir` と `--git-common-dir` の結果が違う）では
  その作業ブランチにコミットしてよい。それ以外で `master` 以外にいたらコミットせず、
  ユーザーに報告して指示を待つ。
- worktree での作業が終わったら、次の順で master へ取り込んでプッシュする。
  1. 作業ブランチにコミットする
  2. `git -C /workspace/python_dev_uv pull --ff-only` でメインのチェックアウトの master を最新にする
  3. `git rebase master` で master の上に付け直す（他の作業が入っていたら検証をやり直す）
  4. `git -C /workspace/python_dev_uv merge --ff-only <作業ブランチ>` で取り込み、`git -C /workspace/python_dev_uv push` する
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
- worktree の `AIMS_nu/.venv` はメインの venv へのシンボリックリンク（Orca のセットアップスクリプトで作る）。
  全 worktree で共有なので、worktree で `uv add` などの venv 変更をしない。
- 改行は `.gitattributes` でリポジトリ内 LF に統一済み（2026-10-08、`.bat`/`.cmd` は CRLF）。
  例外は共有マスタ（`product_registry_shared/*.json`）とコスト表（`cost/*.json`）で、
  アプリが中身のバイト列をハッシュして判定するため CRLF のまま固定している。
  この 2 か所を Python で書き換えるときだけバイナリモード（`rb`/`wb`）で CRLF を保つ。
- マージ済みで不要になったフィーチャーブランチは、削除（ローカルとリモート）を提案する。
  古いブランチに誤ってコミットするのを防ぐため。

---

## Security Rules
- Never request or output secrets, API keys, or credentials.
- Do not log or print personal data.
- Assume production-like constraints even in development.

---

## Cost Awareness
- Keep responses concise.
- Avoid repeating large code blocks unless necessary.
- Prefer explanation over full implementation when possible.
