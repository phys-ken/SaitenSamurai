"""
test_zero_ten_equivalence.py — 0⇔10等価判定の一貫性テスト

マークシートは最大10択(1,2,...,9,0)で、10番目のマーク位置は選択肢"0"。
旧データ形式では同じ位置が"10"と記録されていたため、採点・CTT分析の
両方で "0" と "10" を同一視する必要がある。

このルールは従来 scoring_engine.py / ctt_analyzer.py に個別実装されており、
複数正答(';'区切り)の採点では scoring_engine 側だけ正規化が抜けていた
(採点結果とCTT分析結果がズレる潜在バグ)。本テストは:
  1. 単一正答の0⇔10等価判定(既存挙動の回帰ガード)
  2. 複数正答の0⇔10等価判定(バグ修正の検証)
  3. score_answers と CTTAnalyzer の判定一致(実装統一の検証)
を固定する。
"""
import sys
from pathlib import Path

import pandas as pd
import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "main_src"))

from scoring_engine import score_answers, choice_to_position_index


def _template(q_no, correct_answer, points=2, aspect=1):
    """score_answers 用の最小テンプレート辞書を作る"""
    return {q_no: {'正答': correct_answer, '配点': points, '観点': aspect}}


def _is_correct(correct_answer, student_answer):
    """1問だけ採点して正誤を返すヘルパー"""
    result = score_answers({1: student_answer}, _template(1, correct_answer))
    return result['results'][1]['correct']


class TestSingleAnswerZeroTen:
    """単一正答の0⇔10等価判定(既存挙動の回帰ガード)"""

    def test_key_zero_answer_ten(self):
        assert _is_correct('0', '10') is True

    def test_key_ten_answer_zero(self):
        assert _is_correct('10', '0') is True

    def test_exact_match_still_works(self):
        assert _is_correct('3', '3') is True

    def test_mismatch_still_fails(self):
        assert _is_correct('3', '4') is False

    def test_one_is_not_ten(self):
        """"1" と "10" は別の選択肢"""
        assert _is_correct('1', '10') is False


class TestMultiAnswerZeroTen:
    """複数正答(';'区切り)の0⇔10等価判定(バグ修正の検証)

    従来は複数正答の集合比較で0⇔10正規化が行われず、
    正答キー"9;10"に対し解答"9;0"が誤答扱いになっていた。
    """

    def test_key_with_ten_answer_with_zero(self):
        assert _is_correct('9;10', '9;0') is True

    def test_key_with_zero_answer_with_ten(self):
        assert _is_correct('0;3', '10;3') is True

    def test_pipe_separator(self):
        assert _is_correct('9|10', '9;0') is True

    def test_wrong_multi_answer_still_fails(self):
        assert _is_correct('9;10', '9;1') is False

    def test_partial_answer_still_fails(self):
        """部分解答(片方だけ)は誤答"""
        assert _is_correct('9;10', '0') is False

    def test_empty_answer_still_fails(self):
        assert _is_correct('9;10', '') is False

    def test_multi_exact_match_regression(self):
        """0⇔10が絡まない複数正答の既存挙動は不変"""
        assert _is_correct('2;5', '5;2') is True
        assert _is_correct('2;5', '2;4') is False


class TestChoiceToPositionIndex:
    """選択肢の値→マーク位置インデックス変換の共通ヘルパー

    正答位置の赤字表示(image_renderer)と正答枠オーバーレイ(gui_components)
    が同じ変換ルールを共有することを保証する。
    """

    @pytest.mark.parametrize("choice,num_choices,expected", [
        ('3', 6, 2),       # 通常の選択肢
        ('1', 6, 0),       # 先頭
        ('6', 6, 5),       # 末尾ちょうど
        ('7', 6, None),    # 範囲外
        ('0', 10, 9),      # 選択肢"0" = 10番目の位置
        ('10', 10, 9),     # レガシー"10"表記も同じ位置
        ('0', 6, None),    # 10番目の位置が存在しない
        ('10', 6, None),   # 同上(レガシー表記)
        (3, 6, 2),         # int型も受け付ける
        (3.0, 6, 2),       # float型(Excel由来)も受け付ける
        ('3.5', 6, None),  # 非整数は無効
        ('abc', 6, None),  # 非数値は無効
        ('', 6, None),     # 空文字は無効
        ('-1', 6, None),   # 負数は無効
    ])
    def test_mapping(self, choice, num_choices, expected):
        assert choice_to_position_index(choice, num_choices) == expected


class TestDistractorLegacyTenAggregation:
    """選択肢分析: レガシー"10"表記の解答が"0"行に集約されること

    旧バージョンの読み取り結果("10"表記)を再分析した場合でも、
    該当生徒が選択肢分析表から消えない(正答率とレポートが一致する)。
    """

    def test_legacy_ten_counted_in_zero_row(self):
        from ctt_analyzer import CTTAnalyzer

        # 4人中2人がレガシー表記"10"(=選択肢"0")、1人が現行表記"0"、1人が誤答"3"
        ans_df = pd.DataFrame({
            'StudentID': ['a.jpg', 'b.jpg', 'c.jpg', 'd.jpg'],
            '1': ['10', '10', '0', '3'],
        })
        key_df = pd.DataFrame({'QuestionID': ['1'], 'Key': ['0']})
        az = CTTAnalyzer(ans_df, key_df)
        ds = az.calculate_distractor_analysis()

        zero_rows = ds[(ds['QuestionID'] == '1') & (ds['Choice'] == '0')]
        # "0"行は1行だけ(二重生成されない)
        assert len(zero_rows) == 1
        # "10"の2人と"0"の1人が合算されて3人
        assert zero_rows.iloc[0]['Count_全体'] == 3
        assert zero_rows.iloc[0]['IsKey']
        # 正答率(score_matrix)と選択肢分析が一致する
        assert az.score_matrix['1'].sum() == 3


class TestScoringCttConsistency:
    """score_answers と CTTAnalyzer._calculate_score_matrix の判定一致

    採点結果(答案に印字される点数)とCTT分析(成績レポートの正答率)が
    同じ解答に対して食い違わないことを保証する。
    """

    @pytest.mark.parametrize("correct_key,student_answer", [
        ('0', '10'),      # 単一・0⇔10
        ('10', '0'),      # 単一・逆向き
        ('9;10', '9;0'),  # 複数・0⇔10(旧バグ経路)
        ('0;3', '10;3'),  # 複数・逆向き
        ('9;10', '9;1'),  # 複数・誤答
        ('3', '4'),       # 単一・誤答
    ])
    def test_agreement(self, correct_key, student_answer):
        from ctt_analyzer import CTTAnalyzer

        # score_answers 側
        scoring_correct = _is_correct(correct_key, student_answer)

        # CTTAnalyzer 側(最小構成: 2受験者×1問)
        ans_df = pd.DataFrame({
            'StudentID': ['a.jpg', 'b.jpg'],
            '1': [student_answer, student_answer],
        })
        key_df = pd.DataFrame({'QuestionID': ['1'], 'Key': [correct_key]})
        az = CTTAnalyzer(ans_df, key_df)
        ctt_correct = bool(az.score_matrix['1'].iloc[0] == 1)

        assert scoring_correct == ctt_correct, (
            f"採点とCTT分析で判定が食い違う: 正答={correct_key} 解答={student_answer} "
            f"score_answers={scoring_correct} CTT={ctt_correct}"
        )
