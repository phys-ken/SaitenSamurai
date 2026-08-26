"""
test_scoring_e2e.py — サンプルマークシート画像の End-to-End 採点テスト

sample_basefile/ に同梱されたサンプルファイルを使い、
OMR認識 → 自動閾値推定 → 採点 の一連パイプラインを通して
期待される満点（90点）が得られることを検証する。

テスト対象ファイル:
  - sample_basefile/sample_marksheet.jpg  (マークシート画像)
  - sample_basefile/M2-03-002_座標ファイル.xlsx     (座標定義)
  - sample_basefile/answer_key_sample.xlsx          (正答テンプレート)
"""

import sys
from pathlib import Path

import cv2
import numpy as np
import pytest

# conftest.py で main_src がパスに追加済み
from saitensamurai import (
    parse_excel_coordinates,
    detect_corner_markers,
    apply_perspective_transform,
    recognize_marks,
    load_template,
    score_answers,
    estimate_color_threshold_from_pixels,
    collect_mark_fill_ratios,
    analyze_fill_ratio_distribution,
)

# ============================================================
# 定数
# ============================================================

PROJECT_ROOT = Path(__file__).parent.parent
SAMPLE_DIR = PROJECT_ROOT / "sample_basefile"

IMAGE_PATH = SAMPLE_DIR / "sample_marksheet.jpg"
COORD_EXCEL = SAMPLE_DIR / "M2-03-002_座標ファイル.xlsx"
ANSWER_KEY = SAMPLE_DIR / "answer_key_sample.xlsx"

SKIP_QUESTIONS = 4       # Q1-Q4 は学籍番号領域
EXPECTED_SCORE = 90      # 期待される満点
EXPECTED_MAX_SCORE = 90  # テンプレートの総配点


# ============================================================
# フィクスチャ
# ============================================================

@pytest.fixture(scope="module")
def sample_files_exist():
    """サンプルファイルの存在を確認"""
    for p in [IMAGE_PATH, COORD_EXCEL, ANSWER_KEY]:
        if not p.exists():
            pytest.skip(f"サンプルファイルが見つかりません: {p}")


@pytest.fixture(scope="module")
def coordinates():
    """座標データ"""
    coords, q_groups = parse_excel_coordinates(str(COORD_EXCEL), SKIP_QUESTIONS)
    return coords, q_groups


@pytest.fixture(scope="module")
def corrected_image():
    """射影変換済み画像"""
    with open(str(IMAGE_PATH), 'rb') as f:
        img_bytes = f.read()
    image = cv2.imdecode(np.frombuffer(img_bytes, np.uint8), cv2.IMREAD_COLOR)
    assert image is not None, "画像の読み込みに失敗"

    markers = detect_corner_markers(image, debug=False)
    corrected, _ = apply_perspective_transform(image, markers)
    return corrected


@pytest.fixture(scope="module")
def auto_thresholds(corrected_image, coordinates):
    """自動閾値推定"""
    coords, _ = coordinates
    gray = cv2.cvtColor(corrected_image, cv2.COLOR_BGR2GRAY)

    color_result = estimate_color_threshold_from_pixels([gray], coords)
    recommended_color = color_result['recommended_color_threshold']

    fill_ratios = collect_mark_fill_ratios(gray, coords, recommended_color)
    area_result = analyze_fill_ratio_distribution(fill_ratios)
    recommended_area = area_result['recommended_area_threshold']

    return recommended_color, recommended_area


@pytest.fixture(scope="module")
def template_dict():
    """正答テンプレート"""
    return load_template(str(ANSWER_KEY))


# ============================================================
# テスト
# ============================================================

class TestSampleFilesIntegrity:
    """サンプルファイルの整合性テスト"""

    def test_files_exist(self, sample_files_exist):
        """3つのサンプルファイルがすべて存在する"""
        assert IMAGE_PATH.exists()
        assert COORD_EXCEL.exists()
        assert ANSWER_KEY.exists()

    def test_coordinate_parsing(self, sample_files_exist, coordinates):
        """座標Excelが正しくパースされる"""
        coords, q_groups = coordinates
        assert len(coords) > 0, "座標が0件"
        # 49問 × 10選択肢 = 490 座標が期待される
        assert len(coords) == 490, f"座標数が想定と異なる: {len(coords)} (期待: 490)"

    def test_template_loading(self, sample_files_exist, template_dict):
        """テンプレートが正しく読み込まれる"""
        assert len(template_dict) == 45, f"テンプレート問題数: {len(template_dict)} (期待: 45)"
        total_max = sum(d['配点'] for d in template_dict.values())
        assert total_max == EXPECTED_MAX_SCORE, f"総配点: {total_max} (期待: {EXPECTED_MAX_SCORE})"


class TestCornerDetection:
    """マーカー検出と射影変換のテスト"""

    def test_corner_markers_detected(self, sample_files_exist):
        """4隅のマーカーが検出できる"""
        with open(str(IMAGE_PATH), 'rb') as f:
            img_bytes = f.read()
        image = cv2.imdecode(np.frombuffer(img_bytes, np.uint8), cv2.IMREAD_COLOR)
        markers = detect_corner_markers(image, debug=False)
        assert len(markers) == 4, f"マーカー数: {len(markers)} (期待: 4)"

    def test_corrected_image_size(self, sample_files_exist, corrected_image):
        """補正後の画像サイズが 595x842"""
        h, w = corrected_image.shape[:2]
        assert (w, h) == (595, 842), f"画像サイズ: {w}x{h} (期待: 595x842)"


class TestAutoThreshold:
    """自動閾値推定のテスト"""

    def test_thresholds_in_valid_range(self, sample_files_exist, auto_thresholds):
        """推定された閾値が妥当な範囲に収まる"""
        color_th, area_th = auto_thresholds
        assert 0.03 <= color_th <= 0.35, f"color_threshold: {color_th} (範囲外)"
        assert 0.05 <= area_th <= 0.80, f"area_threshold: {area_th} (範囲外)"


class TestEndToEndScoring:
    """End-to-End 採点テスト（メイン: 満点検証）"""

    def test_full_score_90(self, sample_files_exist, corrected_image,
                           coordinates, auto_thresholds, template_dict):
        """
        自動閾値で OMR 認識 → 採点した結果が 90/90 点（満点）になる。

        パイプライン:
          1. 座標パース
          2. 画像読込 → マーカー検出 → 射影変換
          3. 自動二値化閾値推定
          4. recognize_marks() で OMR 認識
          5. 選択肢番号マッピング (choice_idx → display_value)
          6. score_answers() で採点
          7. assert total_score == 90
        """
        coords, q_groups = coordinates
        color_th, area_th = auto_thresholds

        # OMR 認識
        raw_marks = recognize_marks(
            corrected_image, coords,
            color_threshold=color_th,
            area_threshold=area_th,
        )

        # 各設問の選択肢数を算出
        choice_counts = {}
        for coord in coords:
            qno = coord['question_no']
            choice_counts[qno] = choice_counts.get(qno, 0) + 1

        # raw_choiceルックアップ: {q_no: {sorted_choice_idx: raw_choice_value}}
        raw_choice_map = {}
        for coord in coords:
            q = coord['question_no']
            if q not in raw_choice_map:
                raw_choice_map[q] = {}
            raw_choice_map[q][coord['choice']] = coord['raw_choice']

        # raw_marks → student_answers 変換
        # raw_choice（Excel列ヘッダ値）を表示値として使用
        student_answers = {}
        for q_no, choices in raw_marks.items():
            if q_no <= SKIP_QUESTIONS:
                continue
            scored_q = q_no - SKIP_QUESTIONS
            vals = [str(raw_choice_map[q_no][c]) for c in sorted(choices)]
            student_answers[scored_q] = ';'.join(vals) if vals else ''

        # 採点
        result = score_answers(student_answers, template_dict)

        # アサーション
        assert result['max_score'] == EXPECTED_MAX_SCORE, (
            f"総配点が想定と異なる: {result['max_score']} (期待: {EXPECTED_MAX_SCORE})"
        )
        assert result['total_score'] == EXPECTED_SCORE, (
            f"合計得点が想定と異なる: {result['total_score']} (期待: {EXPECTED_SCORE})\n"
            f"不正解の問題:\n"
            + _format_wrong_answers(result['results'])
        )

    def test_all_questions_answered(self, sample_files_exist, corrected_image,
                                    coordinates, auto_thresholds, template_dict):
        """全45問に対して解答が認識されている（ノーマーク問題がない）"""
        coords, _ = coordinates
        color_th, area_th = auto_thresholds

        raw_marks = recognize_marks(
            corrected_image, coords,
            color_threshold=color_th,
            area_threshold=area_th,
        )

        choice_counts = {}
        for coord in coords:
            qno = coord['question_no']
            choice_counts[qno] = choice_counts.get(qno, 0) + 1

        # raw_choiceルックアップ
        raw_choice_map = {}
        for coord in coords:
            q = coord['question_no']
            if q not in raw_choice_map:
                raw_choice_map[q] = {}
            raw_choice_map[q][coord['choice']] = coord['raw_choice']

        student_answers = {}
        for q_no, choices in raw_marks.items():
            if q_no <= SKIP_QUESTIONS:
                continue
            scored_q = q_no - SKIP_QUESTIONS
            vals = [str(raw_choice_map[q_no][c]) for c in sorted(choices)]
            student_answers[scored_q] = ';'.join(vals) if vals else ''

        # テンプレートの全問に解答があるか
        missing = [q for q in template_dict if q not in student_answers or student_answers[q] == '']
        assert len(missing) == 0, f"未回答の問題: {missing}"

    def test_no_double_marks(self, sample_files_exist, corrected_image,
                             coordinates, auto_thresholds):
        """ダブルマーク（1問に複数選択）が検出されない"""
        coords, _ = coordinates
        color_th, area_th = auto_thresholds

        raw_marks = recognize_marks(
            corrected_image, coords,
            color_threshold=color_th,
            area_threshold=area_th,
        )

        double_marks = {
            q: choices for q, choices in raw_marks.items()
            if q > SKIP_QUESTIONS and len(choices) > 1
        }
        assert len(double_marks) == 0, (
            f"ダブルマーク検出: {double_marks}"
        )


# ============================================================
# ヘルパー
# ============================================================

def _format_wrong_answers(results: dict) -> str:
    """不正解の問題をフォーマットして返す"""
    lines = []
    for q_no in sorted(results.keys()):
        r = results[q_no]
        if not r['correct']:
            lines.append(
                f"  Q{q_no}: 解答={r['student_answer']} "
                f"正答={r['correct_answer']} "
                f"(配点: {r['max_points']}点)"
            )
    return '\n'.join(lines) if lines else '  (なし)'


# ============================================================
# 誤答系 E2E（同じ実画像・わざと違う answer_key で採点する）
# ============================================================
# 従来の実画像 e2e は満点ケースのみで、「間違いが間違いとして採点される」
# ことは画像レベルで検証されていなかった（fable-review/test-plan.md 優先2）。
# 認識結果は満点用と同一なので、answer_key 側を書き換えて誤答を作る。


@pytest.fixture(scope="module")
def student_answers(sample_files_exist, corrected_image, coordinates,
                    auto_thresholds):
    """実画像から認識した解答辞書（採点用番号ベース）"""
    coords, _ = coordinates
    color_th, area_th = auto_thresholds
    raw_marks = recognize_marks(
        corrected_image, coords,
        color_threshold=color_th, area_threshold=area_th,
    )
    raw_choice_map = {}
    for coord in coords:
        raw_choice_map.setdefault(coord['question_no'], {})[coord['choice']] = coord['raw_choice']
    answers = {}
    for q_no, choices in raw_marks.items():
        if q_no <= SKIP_QUESTIONS:
            continue
        vals = [str(raw_choice_map[q_no][c]) for c in sorted(choices)]
        answers[q_no - SKIP_QUESTIONS] = ';'.join(vals) if vals else ''
    return answers


class TestEndToEndWrongAnswers:

    def _score_with_modified_key(self, tmp_path, student_answers, mutate):
        """answer_key をコピーして mutate(df) を適用し、採点結果を返す"""
        import pandas as pd
        df = pd.read_excel(ANSWER_KEY)
        # int64 列に '5;6' 等の文字列を代入すると pandas 2.x は TypeError に
        # なるため、書き換え前に object 型へ寄せる
        df['正答'] = df['正答'].astype(object)
        mutate(df)
        key_path = tmp_path / "modified_key.xlsx"
        df.to_excel(key_path, index=False)
        return score_answers(student_answers, load_template(str(key_path)))

    def test_wrong_keys_reduce_score(self, tmp_path, student_answers):
        """正答を3問書き換えると、その配点ぶんだけ減点される（誤答が誤答になる）"""
        def mutate(df):
            # Q1(正答1)→2, Q2(正答2)→3, Q3(正答3)→4 （各2点）
            for i, wrong in [(0, 2), (1, 3), (2, 4)]:
                df.loc[i, '正答'] = wrong
        result = self._score_with_modified_key(tmp_path, student_answers, mutate)
        assert result['total_score'] == EXPECTED_SCORE - 6
        for q in (1, 2, 3):
            assert result['results'][q]['correct'] is False

    def test_multiple_answer_requires_all_marks(self, tmp_path, student_answers):
        """複数正答キー（;区切り）は集合一致。片方しか塗っていなければ誤答"""
        def mutate(df):
            df.loc[4, '正答'] = '5;6'   # Q5: 生徒は5のみマーク
        result = self._score_with_modified_key(tmp_path, student_answers, mutate)
        assert result['results'][5]['correct'] is False
        assert result['total_score'] == EXPECTED_SCORE - 2

    def test_zero_ten_equivalence_from_image(self, tmp_path, student_answers):
        """キー表記を10→0に変えても実画像の解答は正解のまま（0⇔10等価）"""
        def mutate(df):
            assert str(df.loc[9, '正答']) in ('10', '0')  # Q10 は10番目の位置が正答
            df.loc[9, '正答'] = '0'
        result = self._score_with_modified_key(tmp_path, student_answers, mutate)
        assert result['results'][10]['correct'] is True
        assert result['total_score'] == EXPECTED_SCORE
