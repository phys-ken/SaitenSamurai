# マ�Eク採点の使ぁE��

Mark2 形式�Eマ�Eクシートを自動で読み取り、採点するモードです、E

---

## ワークフロー概要E

```
座標ファイル読込 ↁEスキャン画像読込 ↁEOMR認譁EↁEマ�EクチェチE�� ↁE採点 ↁE結果出劁E
```

---

## 1. ファイルの準備

マ�Eク採点に忁E��なファイルは 3 つです、E

| ファイル | 説昁E|
|---|---|
| **Mark2 座標ファイル** | Mark2 で作�Eした座樁EExcel�E�E.xlsx`�E�E|
| **正答ファイル** | 吁E��の正答を記�Eした Excel�E�E.xlsx`�E�E|
| **スキャン画僁E* | 答案をスキャンした画像！EPEG / PNG / PDF�E�E|

!!! info "Mark2 座標ファイルとは"
    [Mark2�E��E應義塾大学 SFC 研究所�E�](https://github.com/Mark2OSS/Mark2) で作�Eする、�Eークシート上�Eマ�Eク位置を定義した Excel ファイルです、E95ÁE42 pt の A4 座標系を使用してぁE��す、E

!!! tip "模篁E��答�E表示方況E
    模篁E��答�E表示方式につぁE��は、以下�Eサイトも参老E��なります、E
    
    :material-link: [チE��タル採点 All in One  EObject Pascalと僕と](https://coding-tips-memoranda.com/%E3%83%87%E3%82%B8%E3%82%BF%E3%83%AB%E6%8E%A1%E7%82%B9-all-in-one/)

---

## 2. メイン画面の操佁E

モード選択で **「�Eーク採点、E* を選ぶと、�Eーク採点専用のメイン画面が開きます、E

画面の持E��に従って、座標ファイル→正答ファイル→スキャン画像フォルダの頁E��読み込んでください、E

![マ�Eク採点メイン画面](../images/02a_main_mark_only.png){ .screenshot }
<span class="caption">マ�Eク採点モード�Eメイン画面</span>

---

## 3. 閾値キャリブレーション

OMR�E��E学マ�Eク認識）では、�Eークの塗りつぶし度合いを「黒い面積�E比率」で判定します、E

採点侍�E **K-means クラスタリング** と **大津の二値化況E* を絁E��合わせて、E��値を�E動で最適化します、E

キャリブレーション画面では、�E動推定された閾値を確認し、忁E��に応じて手動で微調整できます、E

![閾値キャリブレーション](../images/17_threshold_calibrator.png){ .screenshot }
<span class="caption">閾値キャリブレーション画面</span>

---

## 4. マ�EクチェチE��

OMR の認識結果を�Eに、以下�Eエラーを�E動検�Eします、E

- **未マ�Eク**: どの選択肢も�EークされてぁE��ぁE
- **ダブルマ�Eク**: 褁E��の選択肢が�EークされてぁE��

エラーがある場合�E、�EークチェチE��画面で 1 件ずつ確認し、GUI 上で修正できます、E

![マ�EクチェチE��画面](../images/18_mark_checker.png){ .screenshot }
<span class="caption">マ�EクチェチE��画面  Eエラーの確認と修正</span>

---

## 5. 描画設宁E

採点結果を答案画像に描画する際�E設定をカスタマイズできます、E

○×�Eークの大きさ、色、位置、得点チE��スト�E表示位置などを調整できます、E

![描画設定](../images/03_rendering_settings_mark_only.png){ .screenshot }
<span class="caption">マ�Eク採点モード�E描画設宁E/span>

---

## 6. 合計点の表示位置

合計点を答案画像�Eどこに描画するかを持E��できます、E

![合計点の位置選択](../images/13_select_total_position.png){ .screenshot }
<span class="caption">合計点の表示位置を選抁E/span>

---

## 7. 結果の出劁E

採点が完亁E��ると、以下�Eファイルが�E動生成されます、E

| 出力�E | 冁E�� |
|---|---|
| `01_Results/` | 生徒別成績サマリー Excel |
| `02_Graded_Detail/` | 採点済み答案画像（○×�Eーク・得点付き�E�E|
| `03_Final_Report/` | 試験統訁EExcel、CTT 刁E�� PDF |

---

## 8. 生徒ビューア

個別の生徒�E採点結果を確認するビューアも搭載してぁE��す、E

![生徒ビューア](../images/11_student_viewer.png){ .screenshot }
<span class="caption">生徒別の採点結果ビューア</span>
