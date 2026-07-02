"""
test_question_summary.py — 「問題概要」フィールドの伝播テスト

answer_key.xlsx の任意列「問題概要」が、テンプレート読込 → CTT分析
(key_df/item_stats) → R連携エクスポート(item_info.csv) まで欠落なく
伝播すること、および列が存在しない旧形式テンプレートでも従来どおり
動作することを固定する。
"""
import sys
from pathlib import Path

import pandas as pd
import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "main_src"))

from scoring_engine import load_template
from ctt_analyzer import convert_mark2_to_ctt_data, CTTAnalyzer


SUMMARY_TEXT = "二次関数の頂点を求める"


def _make_template_xlsx(path, with_summary):
    rows = [
        {'問題番号': 1, '正答': '3', '配点': 2, '観点': 1},
        {'問題番号': 2, '正答': '1;2', '配点': 3, '観点': 2},
    ]
    if with_summary:
        rows[0]['問題概要'] = SUMMARY_TEXT
        rows[1]['問題概要'] = ''  # 未入力の行が混ざっても動くこと
    pd.DataFrame(rows).to_excel(path, index=False)


class TestLoadTemplateSummary:
    """load_template の問題概要読込(新旧両形式)"""

    def test_new_format_reads_summary(self, tmp_path):
        p = tmp_path / "key.xlsx"
        _make_template_xlsx(p, with_summary=True)
        td = load_template(p)
        assert td[1]['問題概要'] == SUMMARY_TEXT
        assert td[2]['問題概要'] == ''

    def test_old_format_defaults_to_empty(self, tmp_path):
        """問題概要列が無い旧テンプレートでも読み込める"""
        p = tmp_path / "key.xlsx"
        _make_template_xlsx(p, with_summary=False)
        td = load_template(p)
        assert td[1]['問題概要'] == ''
        assert td[1]['正答'] == '3'
        assert td[1]['配点'] == 2


class TestGenerateTemplateHasSummaryColumn:
    """generate_template が問題概要列を含むテンプレートを出力する"""

    def test_column_exists(self, tmp_path):
        import openpyxl
        from omr_engine import generate_template

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['header1'] * 8)
        ws.append([''] * 8)
        ws.append([''] * 8)
        ws.append([1, 'Q1', '', '', 100, 100, 20, 20])
        coord_path = tmp_path / "coord.xlsx"
        wb.save(str(coord_path))

        result = generate_template(coord_path, tmp_path, skip_questions=0)
        df = pd.read_excel(result)
        assert '問題概要' in df.columns


class TestCttPropagation:
    """CTT分析への伝播(key_df → item_stats)"""

    def _template_dict(self, with_summary):
        td = {
            1: {'正答': '3', '配点': 2, '観点': 1},
            2: {'正答': '1', '配点': 2, '観点': 1},
        }
        if with_summary:
            td[1]['問題概要'] = SUMMARY_TEXT
            td[2]['問題概要'] = ''
        return td

    def _mark2_results(self):
        return [
            {'image': 'a.jpg', 'answers': {1: '3', 2: '1'}},
            {'image': 'b.jpg', 'answers': {1: '2', 2: '1'}},
        ]

    def test_key_df_has_summary(self):
        ans_df, key_df = convert_mark2_to_ctt_data(
            None, None, template_dict=self._template_dict(True),
            mark2_results=self._mark2_results())
        assert 'Summary' in key_df.columns
        assert key_df.loc[key_df['QuestionID'] == '1', 'Summary'].iloc[0] == SUMMARY_TEXT

    def test_key_df_without_summary_field(self):
        """旧形式template_dict(問題概要キー無し)でも空文字で動く"""
        ans_df, key_df = convert_mark2_to_ctt_data(
            None, None, template_dict=self._template_dict(False),
            mark2_results=self._mark2_results())
        assert 'Summary' in key_df.columns
        assert (key_df['Summary'] == '').all()

    def test_item_stats_has_summary(self):
        ans_df, key_df = convert_mark2_to_ctt_data(
            None, None, template_dict=self._template_dict(True),
            mark2_results=self._mark2_results())
        az = CTTAnalyzer(ans_df, key_df)
        stats = az.calculate_item_stats()
        assert 'Summary' in stats.columns
        assert stats.loc[stats['QuestionID'] == '1', 'Summary'].iloc[0] == SUMMARY_TEXT

    def test_old_key_df_without_summary_column(self):
        """Summary列そのものが無いkey_dfでもCTTAnalyzerが動く(後方互換)"""
        ans_df = pd.DataFrame({'StudentID': ['a.jpg', 'b.jpg'], '1': ['3', '2']})
        key_df = pd.DataFrame({'QuestionID': ['1'], 'Key': ['3']})
        az = CTTAnalyzer(ans_df, key_df)
        stats = az.calculate_item_stats()
        assert stats.loc[0, 'Summary'] == ''


class TestPdfWithHostileSummary:
    """数式風の自由入力("x<5"等)がPDFのマークアップを壊さないこと

    reportlabのParagraphはミニHTMLをパースするため、未エスケープの
    '<' や '&' が混入するとPDF生成全体が例外で失敗する。
    """

    def test_pdf_generation_with_markup_chars(self, tmp_path):
        pytest.importorskip("reportlab")
        from ctt_analyzer import CTTPDFReporter

        ans_df = pd.DataFrame({
            'StudentID': [f's{i}.jpg' for i in range(10)],
            '1': ['3', '3', '2', '3', '1', '3', '3', '2', '3', '3'],
            '2': ['1', '2', '1', '1', '1', '2', '1', '1', '2', '1'],
        })
        key_df = pd.DataFrame({
            'QuestionID': ['1', '2'],
            'Key': ['3', '1'],
            'Summary': ['x<5 かつ a&b の範囲', '0<t<2π の <b>タグ風</b>'],
        })
        az = CTTAnalyzer(ans_df, key_df)
        pdf_path = tmp_path / "report.pdf"
        reporter = CTTPDFReporter(str(pdf_path))
        reporter.generate_report(
            az.calculate_test_stats(), az.calculate_item_stats(),
            az.calculate_distractor_analysis(),
            az.total_scores, az.questions, az.score_matrix)
        assert pdf_path.exists()
        assert pdf_path.stat().st_size > 1000


class TestRExportSummary:
    """R連携 item_info.csv への伝播"""

    def test_item_info_has_summary(self, tmp_path):
        from r_export import _export_item_info

        score_matrix = pd.DataFrame({'Q001': [1, 0], 'Q002': [1, 1]})
        key_df = pd.DataFrame({
            'QuestionID': ['1', '2'],
            'Key': ['3', '1'],
            'Summary': [SUMMARY_TEXT, ''],
        })
        col_map = {'1': 'Q001', '2': 'Q002'}
        out = tmp_path / "item_info.csv"
        _export_item_info(score_matrix, key_df, out, col_map=col_map)

        df = pd.read_csv(out)
        assert 'Summary' in df.columns
        assert df.loc[df['ItemID'] == 'Q001', 'Summary'].iloc[0] == SUMMARY_TEXT

    def test_item_info_escapes_formula_injection(self, tmp_path):
        """'='始まりの概要はエスケープされてCSVに書かれる(CWE-1236対策)"""
        from r_export import _export_item_info

        score_matrix = pd.DataFrame({'Q001': [1, 0]})
        key_df = pd.DataFrame({
            'QuestionID': ['1'], 'Key': ['3'], 'Summary': ['=SUM()の使い方'],
        })
        out = tmp_path / "item_info.csv"
        _export_item_info(score_matrix, key_df, out, col_map={'1': 'Q001'})

        df = pd.read_csv(out)
        assert df.loc[0, 'Summary'] == "'=SUM()の使い方"

    def test_item_info_without_summary_column(self, tmp_path):
        """Summary列が無いkey_dfでは従来どおり2列のCSV"""
        from r_export import _export_item_info

        score_matrix = pd.DataFrame({'Q001': [1, 0]})
        key_df = pd.DataFrame({'QuestionID': ['1'], 'Key': ['3']})
        out = tmp_path / "item_info.csv"
        _export_item_info(score_matrix, key_df, out, col_map={'1': 'Q001'})

        df = pd.read_csv(out)
        assert list(df.columns) == ['ItemID', 'MeanScore']
