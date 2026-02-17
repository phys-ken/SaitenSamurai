# Abandoned Improvements (v4.3.2) - 2026-02-17

## 概要
このドキュメントは、2026年2月に試みられたが、最終的に採用されなかった機能改善（撤退案件）に関する記録です。
実装コードと検証テストは `feat/preview-and-check-improvements` ブランチに保存されています。

## 対象ブランチ
`feat/preview-and-check-improvements`

## 試みられた改善内容

### 1. 記述式採点プレビューの改善 (Descriptive Preview Improvement)
- **目的**: 記述式採点のプレビュー画面において、マーク認識枠（BOXED）が描画されていない「クリーンな画像」を表示し、文字の視認性を向上させる。
- **実装アプローチ**:
    - `omr_engine.py`: 補正済みでマーク描画前の画像を `00_Processing_Clean` フォルダに別途保存する処理を追加。
    - `main_gui.py`: `DescriptiveScorerGUI` 起動時に `image_folder` としてクリーンな画像フォルダを渡すように変更。
    - `descriptive_gui.py`: `_SingleQuestionScorer` が `image_folder` (clean) を優先して読み込むように変更。

### 2. マークチェック画面の正答表示 (Mark Check Answer Key Display)
- **目的**: マークチェック画面において、正答（Answer Key）の箇所に赤色の点線枠を表示し、修正作業の効率化を図る。
- **実装アプローチ**:
    - `gui_components.py`: `MarkCheckerGUI` に `template_path` 引数を追加し、正答データを読み込む。
    - `mark_checker.py`: `load_coordinates_csv_checker` を修正し、`mark_coords` カラムを展開して各選択肢の座標情報（`choice`, `x`, `y`, `width`, `height`）を取得可能にする。
    - `gui_components.py`: `_draw_overlay` メソッド内で、現在の問題の正答と一致する選択肢の座標に赤色点線枠（`outline="red", dash=(4, 4)`）を描画する処理を追加。

## 検証結果 (Verification)
- **自動テスト**: `tests/test_scoring_e2e.py` はパス。
- **ビジュアル検証**: `tests/visual_verification.py` および `tests/capture_improvements.py` により、機能自体は動作していることを確認（`tests/visual_report.html` 参照）。
- **撤退理由**: ユーザー判断により、現状の完成度または方向性が期待と合致しなかったため（"うまくいきませんでした"）。

## コードの参照方法
実装詳細やテストコードを確認したい場合は、当該ブランチをチェックアウトしてください。

```bash
git checkout feat/preview-and-check-improvements
```
