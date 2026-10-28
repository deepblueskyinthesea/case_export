# CLAUDE.md

**case-export**：電子カルテの「主術予定表発行」から患者ごとの Excel を
`デスクトップ\症例ファイル\<患者名>` に保存する作業を自動化するツールの開発リポジトリ。
保存した Excel は麻酔記録システム AIMS-ν の患者情報タブ「Excel ファイルから読み込む」で使う。

現在の手作業（自動化の対象）：
1. 電子カルテの「主術依頼箋出力」ボタンを押す
2. 主術予定表発行ダイアログの患者一覧で、目的の患者を右クリック →「発行プレビュー」
3. その患者の Excel ファイルが Excel で開く
4. ファイル → 名前を付けて保存 → デスクトップ\症例ファイル に患者名で保存
5. これをその週の全患者分繰り返す

自動化の方式（画面操作か別の方法か）は未定。

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
| `main.py` | エントリポイント（現状は `uv init` の雛形のみ）。 |
| `pyproject.toml` | uv プロジェクト定義。依存はまだない。 |

構成が固まったらこの表を更新する。

---

## 実行環境

本番は電子カルテと Excel が入った院内の Windows 端末。

開発コンテナ（Linux）では電子カルテも Excel も操作できない。
電子カルテ・Excel の画面操作にかかわる部分は実機でしか確かめられないので、
コンテナで確認できた範囲と実機確認が必要な範囲を分けて報告する。

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
