# クレジット・謝辞

採点侍は、先人の素晴らしいソフトウェアやプロジェクトを参考にして開発されました。
ここに感謝と敬意を表します。

---

## 開発者

<div class="feature-card" markdown>

### phys-ken

採点侍（SaitenSamurai）および採点斬り 2021 の開発者。

:material-web: [https://phys-ken.github.io/phys-ken/](https://phys-ken.github.io/phys-ken/)  
:material-github: [https://github.com/phys-ken](https://github.com/phys-ken)

</div>

---

## 参考にしたプロジェクト

### 採点斬り — 島守睦美 氏

採点侍の前身である **採点斬り 2021** は、島守睦美氏が開発された **「採点斬り！！」** を参考にして生まれました。

採点斬りは、答案をスキャナで読み込み、問題ごとに画像を切り出して効率的に採点するという「デジタル採点」の手法を Visual Basic で実現した先駆的なフリーソフトです。

- **開発者**: 島守睦美 氏
- **公開サイト**: 私立島守学園（現在はアクセス不可）
- **紹介ページ**: [アーカイブ（Wayback Machine）](https://web.archive.org/web/20160625063811/http://www.nurs.or.jp/~lionfan/freesoft_49.html)

### 採点革命 — 竹内俊彦 氏

採点斬りの元となった **「採点革命」** は、竹内俊彦氏（青山学院大学）が開発されたデジタル採点の草分け的ソフトウェアです。

- **開発者**: 竹内俊彦 氏（青山学院大学理工学部）
- **紹介ページ**: [アーカイブ（Wayback Machine）](https://web.archive.org/web/20161024200711/http://www.nurs.or.jp/~lionfan/freesoft_45.html)

### 採点斬り 2021 — phys-ken

「採点革命」や「採点斬り」のコンセプトを受け継ぎ、現在の環境でも動作する Python 版として開発されたプロジェクトです。採点侍の記述式採点機能は、このプロジェクトの設計を基盤としています。

- **ライセンス**: GPL-3.0
- :material-web: [https://phys-ken.github.io/saitenGiri2021/](https://phys-ken.github.io/saitenGiri2021/)
- :material-github: [https://github.com/phys-ken/saitenGiri2021](https://github.com/phys-ken/saitenGiri2021)

---

## 技術基盤

### Mark2 — 慶應義塾大学 SFC 研究所

マークシートの座標系（595×842 pt, A4）と OMR 認識ロジックの基盤として利用しています。

- **ライセンス**: MIT
- :material-github: [https://github.com/Mark2OSS/Mark2](https://github.com/Mark2OSS/Mark2)

---

## 参考サイト

### デジタル採点 All in One — Object Pascalと僕と

模範解答の表示方法など、デジタル採点の手法について参考にさせていただきました。マークシートリーダーや手書き答案採点を含む統合パッケージを公開されています。

- :material-web: [https://coding-tips-memoranda.com/デジタル採点-all-in-one/](https://coding-tips-memoranda.com/%E3%83%87%E3%82%B8%E3%82%BF%E3%83%AB%E6%8E%A1%E7%82%B9-all-in-one/)

---

## デジタル採点の系譜

| 世代 | ソフトウェア | 開発者 | 技術 |
|---|---|---|---|
| 第 1 世代 | **採点革命** | 竹内俊彦 氏 | HSP |
| 第 2 世代 | **採点斬り！！** | 島守睦美 氏 | Visual Basic |
| 第 3 世代 | **採点斬り 2021** | phys-ken | Python |
| 第 4 世代 | **採点侍 (SaitenSamurai)** | phys-ken | Python + tkinter |

---

## 使用ライブラリ

| ライブラリ | ライセンス | 用途 |
|---|---|---|
| OpenCV (headless) | Apache-2.0 / MIT | 画像処理・OMR |
| NumPy | BSD 3-Clause | 数値計算 |
| pandas | BSD 3-Clause | データフレーム処理 |
| Pillow | HPND | 画像描画 |
| openpyxl | MIT | Excel 入出力 |
| PyMuPDF | AGPL-3.0 | PDF→画像変換（オプション） |
| matplotlib | PSF/BSD 互換 | CTT グラフ描画 |
| ReportLab | BSD 3-Clause | CTT PDF レポート |
| PyInstaller | GPL-2.0+（Bootloader Exception 付） | exe ビルド |

詳細は [THIRDPARTYLICENSES.md](https://github.com/phys-ken/SaitenSamurai/blob/main/THIRDPARTYLICENSES.md) をご確認ください。
