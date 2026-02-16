---
hide:
  - navigation
---

<div class="hero" markdown>

![採点侍ロゴ](images/samurai.png){ .hero-logo }

# 採点侍 — SaitenSamurai

<p class="tagline">
普通紙マークシート採点 ＆ 記述式採点を、<strong>これ1本で。</strong><br>
教員による教員のための、Windows向け無料採点支援ソフトウェアです。
</p>

[ダウンロード (Windows)](download.md){ .btn-download }

</div>

---

## 3つの採点モードで、あらゆる試験に対応

<div class="feature-grid" markdown>

<div class="feature-card" markdown>

### :material-checkbox-marked-outline: マーク採点モード

Mark2 形式のマークシートを自動で読み取り・採点。閾値の自動キャリブレーション、未マーク・ダブルマーク検出と GUI での修正に対応しています。

</div>

<div class="feature-card" markdown>

### :material-text-box-edit-outline: 記述式採点モード

スキャン画像から問題領域をマウスで指定し、1枚ずつ or 一覧グリッドで効率的に採点。○×△の判定と部分点にも対応しています。

</div>

<div class="feature-card" markdown>

### :material-clipboard-check-outline: マーク＋記述 混合モード

マーク式と記述式が混在する試験もワンストップで処理。1つのワークフローで、すべての採点が完結します。

</div>

</div>

起動画面でモードを選択するだけで、すぐに使い始められます。

![モード選択画面](images/01_startup_mode_dialog.png){ .screenshot-small }
<span class="caption">起動時のモード選択ダイアログ</span>

---

## 主な特徴

<div class="feature-grid" markdown>

<div class="feature-card" markdown>

### :material-auto-fix: 自動傾き補正

コーナーマーカー検出と射影変換で、スキャン時の傾きを自動補正。普通のコピー機・スキャナで十分です。

</div>

<div class="feature-card" markdown>

### :material-chart-bar: CTT 分析

古典的テスト理論（α係数, P値, D値, I-T相関）を自動算出し、PDF レポートとして出力。試験の質を可視化できます。

</div>

<div class="feature-card" markdown>

### :material-file-excel-box: Excel 一括出力

生徒別の成績サマリー Excel、試験統計 Excel を自動生成。採点済み答案画像（○×マーク付き）も一括で出力します。

</div>

<div class="feature-card" markdown>

### :material-content-save: セッション保存

作業途中の状態を保存し、後から再開可能。大量の答案も安心してペースに合わせて処理できます。

</div>

</div>

---

## 動作環境

| 項目 | 要件 |
|---|---|
| **OS** | **Windows 11**（動作確認済み） |
| **必要なもの** | SaitenSamurai.exe のみ（インストール不要） |
| **スキャン画像** | JPEG / PNG / PDF に対応 |

!!! info "テスト用サンプル付き"
    `sample_basefile/` フォルダに、座標ファイル・正答テンプレート・スキャン画像のサンプルが同梱されています。まずはサンプルで動作を確認してみてください。

---

## はじめる

1. **[ダウンロード](download.md)** — GitHub Releases から最新の exe をダウンロード
2. **[クイックスタート](quickstart.md)** — 最初の採点を 5 分で体験
3. **[使い方ガイド](usage/mark.md)** — 各モードの詳しい操作方法

---

<div style="text-align: center; margin-top: 2em; color: #999; font-size: 0.85em;">
採点侍は <a href="https://github.com/phys-ken">phys-ken</a> が開発する GPL-3.0 ライセンスのオープンソースソフトウェアです。
</div>
