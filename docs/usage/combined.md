# マ�Eク�E�記述の使ぁE��

マ�Eク式と記述式が混在する試験を、Eつのワークフローで採点するモードです、E

---

## ワークフロー概要E

```
座標ファイル読込 ↁEスキャン画像読込 ↁEOMR認譁EↁEマ�EクチェチE��
  ↁE記述式�E領域設宁EↁE記述式�E採点 ↁE統合して結果出劁E
```

マ�Eク部刁E�Eマ�Eク採点モードと同じ自動認識、記述部刁E�E手動採点で処琁E��、最後に合算して結果を�E力します、E

---

## 1. ファイルの準備

マ�Eク�E�記述モードでは、�Eーク採点と同じ 3 つのファイルを用意します、E

| ファイル | 説昁E|
|---|---|
| **Mark2 座標ファイル** | Mark2 で作�Eした座樁EExcel |
| **正答ファイル** | マ�Eク問題�E正答を記�Eした Excel |
| **スキャン画僁E* | 答案をスキャンした JPEG / PNG / PDF |

---

## 2. メイン画面

モード選択で **「�Eーク�E�記述、E* を選ぶと、混合モード�Eメイン画面が開きます、E
マ�Eク採点のスチE��プに加え、記述式�E設定スチE��プが追加されてぁE��す、E

![マ�Eク�E�記述メイン画面](../images/02b_main_mark_descriptive.png){ .screenshot }
<span class="caption">マ�Eク�E�記述モード�Eメイン画面</span>

---

## 3. マ�Eク部刁E�E処琁E

マ�Eク部刁E�E処琁E�E、[マ�Eク採点モード](mark.md) と同じです、E

1. **座標ファイル・正答ファイルの読み込み**
2. **OMR�E��E学マ�Eク認識）による自動読み取り**
3. **閾値キャリブレーション**�E�忁E��に応じて�E�E
4. **マ�EクチェチE��**�E�未マ�Eク・ダブルマ�Eクの確認と修正�E�E

---

## 4. 記述式�E領域設宁E

マ�Eク部刁E�E処琁E��完亁E��たら、記述式�E領域を追加で設定します、E

統合セチE��アチE�E画面で、記述式�E問題領域を�Eウスで持E��します、E

![統合セチE��アチE�E](../images/19_integrated_descriptive_setup.png){ .screenshot }
<span class="caption">記述式�E領域設定（統合セチE��アチE�E画面�E�E/span>

セチE��アチE�E時�Eアクション選択では、E��域の追加・変更が可能です、E

![セチE��アチE�Eアクション](../images/12_descriptive_setup_action.png){ .screenshot-small }
<span class="caption">記述式セチE��アチE�Eのアクション選抁E/span>

---

## 5. 記述式�E採点

記述式領域の設定が完亁E��たら、[記述式採点モード](descriptive.md) と同じ方法で手動採点を行います、E

- **1枚ずつ表示**: ○×△ボタンで 1 人ずつ採点
- **グリチE��一覧**: 全生徒�E解答を一覧で確認しながら採点

![記述式スコアラー一覧](../images/10_descriptive_scorer_list.png){ .screenshot }
<span class="caption">記述式問題�E一覧画面</span>

採点レビュー画面で、記述式�E採点結果を確認できます、E

![記述式レビュー](../images/09_descriptive_review_gui.png){ .screenshot }
<span class="caption">記述式�E採点レビュー画面</span>

---

## 6. 描画設宁E

マ�Eク�E�記述モード�E描画設定では、�Eーク部刁E��記述部刁E�E両方の表示スタイルを設定できます、E

![描画設定](../images/03_rendering_settings.png){ .screenshot }
<span class="caption">マ�Eク�E�記述モード�E描画設宁E/span>

---

## 7. 結果の出劁E

マ�Eク部刁E��記述部刁E�E点数が合算され、統合結果として出力されます、E

| 出力�E | 冁E�� |
|---|---|
| `01_Results/` | 生徒別成績サマリー Excel�E��Eーク�E�記述の合計！E|
| `02_Graded_Detail/` | 採点済み答案画像（○×�Eーク・得点すべて描画�E�E|
| `03_Final_Report/` | 試験統訁EExcel、CTT 刁E�� PDF |

!!! info "セチE��ョン保存�E活用"
    混合モードでは作業量が多くなるため、E*セチE��ョン保孁E* を活用して作業を�E割することをおすすめします。作業途中の状態を保存し、後から�E開できます、E
