# 記述式採点の使ぁE��

スキャン画像から採点領域を指定し、記述式�E答案を効玁E��に採点するモードです、E

---

## ワークフロー概要E

```
スキャン画像読込 ↁE採点領域の設宁EↁE問題ごとに採点 ↁE結果出劁E
```

---

## 1. スキャン画像�E準備

記述式採点モードでは、Mark2 座標ファイルは不要です。忁E��なのはスキャン画像�Eみです、E

| 忁E��なも�E | 説昁E|
|---|---|
| **スキャン画僁E* | 答案をスキャンした JPEG / PNG / PDF |

!!! tip "スキャン時�EポインチE
    - 解像度は **200、E00 dpi** がおすすめでぁE
    - コーナ�Eマ�Eカーがある場合�E、傾き補正が�E動で行われまぁE
    - 普通�Eコピ�E機�E「スキャン→フォルダ保存」機�Eで十�EでぁE

---

## 2. メイン画面

モード選択で **「記述式採点、E* を選ぶと、記述式専用のメイン画面が開きます、E

![記述式メイン画面](../images/02c_main_descriptive_only.png){ .screenshot }
<span class="caption">記述式採点モード�Eメイン画面</span>

---

## 3. 採点領域の設宁E

答案画像上で、各問題�E採点領域を�EウスドラチE��で持E��します、E

問題�E数だけ領域を設定します。追加の領域が忁E��かどぁE��を、ダイアログで確認されます、E

![領域選択](../images/15_select_region.png){ .screenshot }
<span class="caption">マウスドラチE��で採点領域を指宁E/span>

問題情報�E��E点、採点方式）を設定するダイアログが表示されます、E

![問題情報の設定](../images/04_ask_question_info.png){ .screenshot-small }
<span class="caption">問題情報の設定ダイアログ</span>

追加の問題がある場合�E、続けて設定します、E

![追加確認](../images/05_ask_add_more.png){ .screenshot-small }
<span class="caption">問題�E追加確誁E/span>

---

## 4. 採点方況E

記述式採点には **2 つの表示モーチE* があります、E

### 1枚ずつ採点モーチE

1 人の答案を 1 枚ずつ表示し、拡大画像を見ながら採点します、E
○（正解�E��E×（不正解�E��E△�E�部刁E���E��Eタンで判定します、E

![1枚ずつ採点](../images/desc_02_scorer_buttons.png){ .screenshot }
<span class="caption">1枚ずつの採点画面  E○×△ボタンで判宁E/span>

採点結果は即座にフィードバチE��表示されます、E

=== "○（正解�E�E

    ![正解の表示](../images/desc_03_scored_maru.png){ .screenshot }
    <span class="caption">○判定：背景が緑に変化</span>

=== "×（不正解�E�E

    ![不正解の表示](../images/desc_04_scored_batsu.png){ .screenshot }
    <span class="caption">×判定：背景が赤に変化</span>

=== "△�E�部刁E���E�E

    ![部刁E��の表示](../images/desc_05_scored_middle.png){ .screenshot }
    <span class="caption">△判定：部刁E��を�E劁E/span>

### グリチE��一覧モーチE

全生徒�E同じ問題を一覧グリチE��で表示し、素早く採点できます、E
大量�E答案を効玁E��に処琁E��たいときに便利です、E

![グリチE��モード](../images/desc_08_grid_mode.png){ .screenshot-wide }
<span class="caption">グリチE��一覧モーチE E全生徒�E解答を一覧表示</span>

---

## 5. 未採点フィルタ

採点漏れを防ぐために、未採点の答案だけを表示するフィルタ機�Eがあります、E

![未採点フィルタ](../images/desc_06_filter_active.png){ .screenshot }
<span class="caption">未採点フィルタを有効にした状慁E/span>

すべての採点が完亁E��ると表示が�Eり替わります、E

![全採点完亁E(../images/desc_07_all_scored.png){ .screenshot }
<span class="caption">すべての採点が完亁E��た状慁E/span>

---

## 6. 描画設宁E

記述式採点用の描画設定では、○×△マ�Eクの表示めE��点チE��スト�Eスタイルをカスタマイズできます、E

![描画設定（記述式）](../images/03_rendering_settings_desc_only.png){ .screenshot }
<span class="caption">記述式採点の描画設宁E/span>

---

## 7. 結果の出劁E

採点が完亁E��ると、以下�EファイルぁE`_saiten_grading_results/` フォルダに自動生成されます、E

| 出力�E | 冁E�� |
|---|---|
| `01_Results/` | 生徒別成績サマリー Excel |
| `02_Graded_Detail/` | 採点済み答案画像（○×△マ�Eク・得点付き�E�E|
| `03_Final_Report/` | 試験統訁EExcel |
