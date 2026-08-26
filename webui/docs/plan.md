# webui 移植計画（正本）

grill-me 合意（2026-08-06）:
土台=同居・凍結 / UI=ビルドなし素のJS / 順序=モード縦切り / Winゲート=後回し。
制約: 後方互換ベストエフォート（既存 session_state.json を読める）、
既存機能を確実に移植（手書きは最後）、入出力は現行と同一、テストは多角的に。
コメントモード（1枚採点で答案に書き込み）の余白を残す。

## アーキテクチャ

```
webui/
  app.py                 # エントリ。ここだけが import webview する
  api/
    bridge.py            # js_api。薄い層 — main_src のドライバを呼ぶだけ
    jobs.py              # 長時間処理のスレッド化・進捗push・cancel_event
    session_compat.py    # 既存 session_state.json の読み込みアダプタ
  static/
    index.html
    css/app.css
    js/                  # ES modules。ビルドなし
      bridge.js          # pywebview.api ラッパ（テスト時はモック注入点）
      main.js ...
  requirements.txt       # 実行時依存 (pywebview)
  requirements-dev.txt   # 開発依存 (PySide6=Linuxプレビュー, playwright, pytest)
```

設計原則:
- **bridge.py は pywebview を import しない**（純粋な呼び出し層）。
  → CI(Windows) でも pywebview 無しで pytest が回る
- UI は標準 Web API のみ（WebView2 固有機能を使わない）→ エンジン差を最小化
- 答案画像の表示は「画像レイヤ＋注釈レイヤ」の重ね構造で作る
  → 将来のコメントモード（PointerEvent手書き）の差し込み口
- 進捗: Python スレッド → window.evaluate_js("app.onProgress(...)")
- ファイル選択は webview.create_file_dialog（ネイティブ）

## テスト体系（4層）

| 層 | 何を | どこで | 実行 |
|---|---|---|---|
| L1 | bridge/API（pywebviewなしで直接呼ぶ） | tests/webui/ | 全環境・CI込み |
| L2 | UI単体（bridgeモックで画面動作） | webui/tests_browser/ | Playwright+Chrome(この機械) |
| L3 | 統合スモーク＋目視 | pywebview(Qt)+Xvfb キャプチャ | この機械・随時 |
| L4 | ゴールデン（同一入力→tk版と同一出力） | tests/webui/ | 全環境 |
| 既存 | ロジック1028件 | tests/ | 触らず緑維持 |

## マイルストーン

- **M1: マーク採点のみモード完動**（フォルダ/座標/正答選択→認識→マークチェック
  →採点→集計。ゴールデンで tk 版と同一出力）
- M2: 数学（複数桁）モード
- M3: 記述採点（最大の山: 領域設定ドラッグ・グリッド採点・1枚採点）
- M4: セッション復元（既存 session_state.json 互換）
- M5: Windows ゲート（IME/exe/高DPI）→ その後にコメントモード
