# 採点侍 ユーザーマニュアル


**このマニュアルでは、採点侍の基本的な使い方から応用的な機能まで説明します。**

## はじめに

採点侍は、学校の定期試験や小テストの採点作業を効率化するためのソフトウェアです。スキャンした解答用紙から特定の領域を切り出し、様々な採点モードで素早く採点を行い、結果を管理することができます。

## インストール方法

インストールは不要です。以下の手順で実行できます：

1. [公式リポジトリ](https://github.com/phys-ken/SaitenSamurai)または[リリースページ](https://github.com/phys-ken/SaitenSamurai/releases)から最新版をダウンロード
2. ダウンロードしたZIPファイルを解凍
3. 解凍したフォルダ内の**採点侍.exe**をダブルクリック

**注意**: 本ソフトウェアはWindows環境専用です。Macには対応していません。

## 基本的な使い方

### 1. 初期設定

1. アプリを起動後、「初期設定をする」ボタンを押します
2. 同じ場所に`setting`フォルダが作成されます
   
   ![初期設定画面](../resources/setup_screen.png)

### 2. 解答用紙の準備

1. スキャンした解答用紙画像を`setting/input`フォルダに保存します
   - テスト用の画像は`test_figs`フォルダにあります
   - JPEGまたはPNG形式を推奨します

### 3. 採点領域の設定

1. 「どこを斬るか決める」ボタンをクリック
2. 表示された解答用紙の上で:
   - 最初の選択領域は「氏名欄」となります
   - 2つ目以降の選択領域は「回答部分」となります
   - マウスでドラッグして範囲を指定します
   - 実寸で0.5cm程度余白を取るようにすると、スキャン時のズレに対応できます
3. 範囲設定が完了したら「入力終了」をクリック

### 4. 解答用紙の切り取り処理

1. 「全員の解答用紙を斬る」ボタンをクリック
2. 処理の進捗状況は画面右側の「採点状況」パネルで確認できます
3. 処理が完了すると「設定」フォルダ内に必要なファイルが生成されます

### 5. 採点作業

#### 5-1. 採点問題の選択

1. 「斬った画像を採点する」ボタンをクリック
2. 問題選択画面が表示されます
3. 採点したい問題と採点モードを選択し「採点開始」をクリック
   - 「1枚ずつ採点」モード: 従来型の1画像ずつを評価
   - 「一覧採点」モード: グリッド表示で複数画像を効率的に採点

#### 5-2. 一覧採点モード（グリッド表示）

![グリッド採点](../resources/grid_grading.png)

一覧採点では以下の3つの採点スタイルから選択できます：

1. **一つずつクリック採点**:
   - まず画像を選択し、次に点数ボタンをクリック
   - Ctrlキーで複数選択、Shiftキーで範囲選択が可能

2. **連続クリック採点**:
   - まず点数ボタンを選択してアクティブにする
   - その後クリックした画像すべてに同じ点数が設定される

3. **数字キーで連続採点**:
   - 画像を選んで数字キーを押すと採点と同時に次の画像へ移動
   - 効率的に連続採点が可能

画面上部の機能：
- 表示サイズ調整: スライダーでサムネイルサイズを変更
- 並び替え: 「ファイル名順」「点数順」などで並べ替え
- 採点実行: 採点結果を保存してExcelファイルを更新

#### 5-3. 1枚ずつ採点モード

従来型の採点方法で、1枚ずつ順に採点していきます：

- 数字キーで得点を入力
- 矢印キーで「次へ進む・前に戻る」
- スペースキーでスキップ（後で再度表示）

### 6. 結果の出力

#### 6-1. Excel出力

1. 「Excelに出力」ボタンを押して`setting/saiten.xlsx`を作成
2. 学生ごと、問題ごとの採点結果が一覧表示されます
   
   ※ Excel ファイルを開いた状態で採点実行するとエラーになる場合があります

#### 6-2. 採点済み解答用紙の出力

1. 「採点済み答案を出力」ボタンをクリック
2. 以下のオプションを選択できます：
   - 設問ごとの得点表示
   - 合計点の表示
   - 〇×△マークの表示
   - 透過度、表示位置、色などの細かい設定
3. 採点結果を反映した解答用紙が`setting/export`フォルダに保存されます

## その他の機能

### 模範解答の切り取り

「その他の機能」メニューから「模範解答のみ斬る」機能を使用すると、`setting/answerdata`フォルダ内の解答用紙のみを斬り取ることができます。これにより模範解答を効率的に管理できます。

## 今後追加予定の機能

- **PDF読み取り機能**: スキャンしたPDFから直接解答用紙を読み込む機能
- **PDF書き出し機能**: 採点結果をPDF形式で出力する機能
- **AI連携機能**: 解答内容を自動認識して採点を支援する機能

## よくある質問

* **Q: 画面が「応答なし」になる、処理に時間がかかる場合は？**
  * A: 画像処理は時間がかかることがあります。処理中のログはコンソール画面で確認できます。画像ファイルのサイズを小さくする、グレースケールで保存するなどで改善できる場合があります。

* **Q: 採点済みの項目を再確認・再採点するには？**
  * A: 一覧採点モードを使用すると簡単に確認・再採点が可能です。1枚ずつ採点モードでは、`setting/output/Q_000X`内の採点済みフォルダから対象ファイルを直下に移動して再採点します。

* **Q: 採点結果の画像がグレースケールになってしまう**
  * A: 画像の最適化処理により、元画像が高画質・大容量の場合にグレースケール化されることがあります。元画像のファイルサイズを小さくすると解決する場合があります。

* **Q: 採点結果の得点表示が重なる**
  * A: 採点時の文字サイズは氏名欄の高さの半分に設定されています。氏名欄が高すぎると文字が重なることがあります。氏名欄の高さを解答欄と同程度にすると解決します。

## サポートとフィードバック

* バグの報告や機能リクエストは[GitHub Issues](https://github.com/phys-ken/SaitenSamurai/issues)または[note](https://note.com/phys_ken)でお知らせください。
* 今後のアップデート情報は公式リポジトリで随時公開します。

## 関連ソフト

* [ウラオモテヤマネコ](https://phys-ken.github.io/uraomoteYamaneko/) - 両面スキャンデータの整理用ソフト