# 採点侍 (SaitenSamurai) — AIエージェント向けの前提と制約

> **環境全体（他のアプリ・ポート・機械の状態）のことは、このファイルに書かない。**
> → 末尾の「環境全体のことは、ここに書かない」を読むこと


このファイルは Claude Code などのエージェントが毎回読む「変わらない制約」だけを置く場所です。
アーキテクチャ・モジュール構成・詳細な設計判断は [DEVELOPMENT.md](DEVELOPMENT.md) にあります。
新しい実装ルールを足すときは、恒久的な制約なら**ここ**、設計の説明なら **DEVELOPMENT.md** に書いてください。

## このソフトが何か

学校の先生が、スキャンした答案を採点するための Windows デスクトップアプリです。
Python + tkinter の GUI アプリで、単体 exe（PyInstaller）として配布しています。
利用者は非エンジニアの教員で、**採点結果が間違うと生徒の成績が壊れます**。
正しさと、壊れたときに気づけることを、速さより優先してください。

## 絶対にやらないこと

- **個人データをコミットしない。** 生徒名・答案画像・実データ由来の answer_key・座標ファイルは
  すべて `.gitignore` 済みですが、新しいサンプルを足すときは必ず匿名の合成データにしてください。
- **`main` へ直接 push しない。** 作業ブランチを切って、CI（`.github/workflows/test.yml`,
  windows-latest）が緑になってから main へ。
- **`stats-data` ブランチを触らない。** GitHub Actions がダウンロード集計を自動 push する専用ブランチです。
- **リリースを勝手に作らない。** タグ push で exe がビルドされて公開されます。人間の判断が必要です。

## 依存を足したときの必須手順

`requirements.txt` に依存を追加したら、**`.github/workflows/release.yml` の
`Install dependencies` ステップにも手で同じものを足す**こと。

ここが同期していないと、PyInstaller は「パッケージが見つからない」warning を出しつつ
**ビルド自体は成功してしまい**、機能が欠けた壊れた exe がそのまま Release として公開されます。
実際に v4.5.1 で scikit-learn が抜けて K-means 機能が死んだ exe を出しています（v4.5.2 で修正）。
詳細と検証用スモークテストは DEVELOPMENT.md の「リリースビルドの依存関係に関する注意」を参照。

## テストの回し方

```bash
# ロジック系（この Linux 環境で回せるのはここまで）
.venv/bin/python -m pytest tests/ -q --timeout=60 -p no:warnings

# GUI テストも含めるなら仮想ディスプレイが要る
xvfb-run -a .venv/bin/python -m pytest tests/ -q --timeout=60 -p no:warnings
```

- `--timeout` は必ず付けてください。GUI テストはハングするとタイムアウトしないと終わりません。
- `visual` / `legacy_mock` マーカーは `pyproject.toml` で既定除外されています。
  UI を変えたときは `-m visual` を明示実行してください。
- **この開発機（Linux/kenmini）には X ディスプレイがありません。**
  GUI を含む本当の確認は GitHub Actions（windows-latest）か Windows 実機でやること。
  ローカルで緑でも「GUI は未検証」であることを、報告のときに正直に書いてください。

## コードの決まりごと

- `constants.py` は他の `main_src/` モジュールを import してはいけません（循環 import の防止線）。
- `scoring_engine.py` は純粋ロジックのみ。ファイル I/O・画像処理・GUI を持ち込まないこと。
- 型チェックは pyright の `basic` モード。**error は 0 件を維持**（warning は段階的解消中）。
- ログは標準 `logging`。各モジュールで `logger = logging.getLogger(__name__)`。
- パスは `pathlib.Path`。exe 同梱リソースは `resource_path()` 経由（PyInstaller 互換のため）。
- PyMuPDF / matplotlib / reportlab はオプショナル依存。`HAS_*` フラグで分岐し、
  未インストールでも起動できる状態を保つこと。

## 言語

コミットメッセージ・ドキュメント・UI 文言・コメントはすべて**日本語**です。
コミットは Conventional Commits 形式（`feat:` `fix:` `docs:` `chore(release):` など）。

## いま何をしているか

`.status` に現在の状態と直近のゴールが書いてあります。作業を始める前に読んでください。
過去の作業から得た教訓は `_dev_notes/lessons/` にあります（ローカル専用、git 管理外）。

---

## 環境全体のことは、ここに書かない

他のアプリ・ポート番号・機械の状態・定時ジョブの時刻は、このファイルにも `README.md` にも
書き写さない。

- **「今どうなっているか」**（ポート・稼働中か・最後に走った時刻）
  → hub **http://kenmini.tail6d3d82.ts.net:8000/** が実機から生成して出す
- **「なぜそうしたか」「環境全体の方針」**
  → Obsidian `環境メモ.md`（Claude Code なら `/ob-env-read`）

書き写した瞬間に、同じ事実の3つ目のコピーになる。コピーは必ずズレる。
