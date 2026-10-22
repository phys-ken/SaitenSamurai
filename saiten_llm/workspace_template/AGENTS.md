# 採点の作業フォルダ — AI エージェントへの指示

このフォルダは、先生が AI エージェント（Claude Code・Antigravity など）と一緒に、
採点侍（SaitenSamurai）で答案を採点するための作業フォルダです。
生徒の答案は機密情報です。下の約束を必ず守ってください。

## 最初に必ず読むもの

- 環境がまだ無いとき（`app/` が無い）: 下の「環境の用意」を先に行う
- 採点を始める・続けるとき: **`app/saiten_llm/GUIDE.md` を最後まで読んでから**作業する

## いちばん大事な約束（詳しくは GUIDE.md）

1. `projects/*/scans/`・`projects/*/除外/`・`projects/*/*.pdf` は開かない・コピーしない。
   見てよい画像は `projects/*/_llm_view/` の中だけ（先生が黒塗りを承認した写し）。
2. 先生の承認が要ることは、先生の言葉を得てから進める。
3. 答案の中身・生徒の名前を、Web 検索や外部サービスに入れない。

## フォルダの構成

```
<このフォルダ>/
  AGENTS.md            この指示（先生が置いたもの）
  CLAUDE.md            Claude Code 用の入口（AGENTS.md を読み込むだけ）
  .claude/settings.json  元の答案を読めないようにする設定
  app/                 採点侍（GitHub から取得）
  templates/           座標ファイルの控え
  projects/<試験名>/    1回の採点ごとのフォルダ
```

## 環境の用意（`app/` が無いときだけ）

先生に「採点侍を用意します。次の操作をしてよいですか」と、下の手順をまとめて伝え、許可を得てから行います。
インストールが要るもの（uv・Git）は、入れる前に必ず先生の許可を得ます。

1. 道具の確認: `git --version` と `uv --version`
   - uv が無い: Windows なら `powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"`
     （入れた後は端末を開き直すか、表示されたパスを使う）
   - Git が無い: Windows なら `winget install --id Git.Git -e`
2. 採点侍の取得（このフォルダで）:
   `git clone -b feat/llm-workflow https://github.com/phys-ken/SaitenSamurai.git app`
3. Python 環境（`app/` で）:
   `uv venv --python 3.12` → `uv pip install -r requirements.txt -r webui/requirements.txt`
4. 作業フォルダの仕上げと動作確認（`app/` で）:
   `uv run python -m saiten_llm setup-workspace ..`
   - `projects/`・`templates/`・`.claude/settings.json`・`CLAUDE.md` を作り（既にあれば残す）、
     合成の答案（実在の生徒はいない）で動作を確かめます。
5. 先生に「設定を有効にするため、エージェントを一度終了して、このフォルダで起動し直してください」と伝える。

採点侍を新しくするときは、先生の許可を得て `app/` で `git pull` を実行します。
