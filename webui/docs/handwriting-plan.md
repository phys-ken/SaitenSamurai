# 手書きコメント機能 計画（grill-me 確定 2026-08-26）

当初要件「一枚採点モードにコメントモード（マウス/ペンタブで答案に書き込み）」の実装。
一枚採点ビューの annotation-layer は最初からこの差し込み口として設計してある。

## データ設計

- `01_Results/handwriting.json`（tk 互換群と同じ置き場・アトミック保存）
  ```json
  {"version": 1,
   "sheets": {"s001.png": {"w": 595, "h": 842,
     "strokes": [{"color": "#C73E2E", "width": 2.0,
                  "points": [[x, y, pressure], ...]}]}}}
  ```
- 座標は 00_Processing 画像の実ピクセル（記述領域と同じ座標系）。
  書き込み時にそのシートの natural サイズも保存し、焼き込み時の倍率に使う
- 元画像・補正画像は不可侵。焼き込みは「採点済み答案を生成」の最終段で、
  ドライバ出力（02_Graded_Detail）に webui 側が PIL で合成する
  → main_src 無改修・何度でも作り直せる原則を維持

## 操作設計

- ペン（pointerType=pen）は常時描画。筆圧で線幅が変わる
- マウスは「✎コメント」トグル（キー C）で描画モードに。オフ時は従来のパン
- ツール: 朱/青/墨・太さ2段階・消しゴム（ストローク単位）・戻す/やり直す・全消去
- Ctrl+ホイールズームは描画モード中も有効。ズームは再描画（ベクターなので劣化なし）
- マークのみモードにも Step2 に「答案に書き込む」を追加し、同じ一枚ビューを
  コメント専用（採点パネルなし）で開く

## 実装

1. bridge: HandwritingMixin（get/set_handwriting、_bake_handwriting）
2. jobs: _scoring_worker の全モード分岐後に焼き込みフック
3. JS: handwriting.js（canvas 描画・ツールバー・undo/redo・消しゴム判定）
   一枚採点 layoutSheet と連動してズーム追従
4. テスト: L1（保存往復・焼き込み画素・倍率）/ L2（トグル描画・undo・保存呼出）
