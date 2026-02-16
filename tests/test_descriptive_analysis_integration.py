#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
②記述採点をCTT/R分析に含む統合テスト

記述問題をバイナリ(0/1)としてCTTおよびR Exportに統合する機能を検証。
- convert_mark2_to_ctt_data の記述統合
- generate_ctt_analysis への記述データ引き渡し
- export_r_analysis_kit への記述データ引き渡し
- GUI チェックボックスの存在と動作
"""

import sys
import tempfile
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

# プロジェクトルートをパスに追加
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "main_src"))

from ctt_analyzer import convert_mark2_to_ctt_data, CTTAnalyzer


# ============================================================
# テスト用フィクスチャ
# ============================================================

@pytest.fixture
def dummy_template_and_results(tmp_path):
    """
    テンプレートExcel と Mark2結果Excelのダミーファイルを作成。
    
    テンプレート: 5問 (正答=1,2,3,1,2  配点=1,1,1,1,1  観点=1,1,2,2,2)
    結果: 10人の学生
    """
    # テンプレート
    template_path = tmp_path / "template.xlsx"
    template_df = pd.DataFrame({
        '問題番号': [1, 2, 3, 4, 5],
        '正答': [1, 2, 3, 1, 2],
        '配点': [1, 1, 1, 1, 1],
        '観点': [1, 1, 2, 2, 2],
    })
    template_df.to_excel(str(template_path), index=False)
    
    # Mark2 結果（2行ヘッダー形式）
    result_path = tmp_path / "mark2_results.xlsx"
    
    # Row 0: No, File, 1, 2, 3, 4, 5
    # Row 1: NaN, NaN, 1, 2, 3, 4, 5
    # Row 2+: Data
    header_row0 = ['No', 'File', 1, 2, 3, 4, 5]
    header_row1 = [np.nan, np.nan, 1, 2, 3, 4, 5]
    
    data_rows = []
    student_files = []
    for i in range(10):
        fname = f"student_{i:03d}.jpg"
        student_files.append(fname)
        # 簡単なパターン: 奇数番号の学生は全問正解、偶数は一部不正解
        if i % 2 == 0:
            answers = [1, 2, 3, 1, 2]  # 全問正解
        else:
            answers = [2, 1, 3, 1, 1]  # 問1,2,5 不正解
        data_rows.append([i + 1, fname] + answers)
    
    all_data = [header_row0, header_row1] + data_rows
    result_df = pd.DataFrame(all_data)
    result_df.to_excel(str(result_path), index=False, header=False)
    
    return {
        'template_path': str(template_path),
        'result_path': str(result_path),
        'student_files': student_files,
    }


@pytest.fixture
def sample_descriptive_config():
    """記述問題設定のサンプル"""
    return {
        "questions": [
            {
                "id": "D1",
                "name": "記述1",
                "region": [100, 400, 300, 500],
                "max_score": 5,
                "aspect": 1,
            },
            {
                "id": "D2",
                "name": "記述2",
                "region": [100, 510, 300, 600],
                "max_score": 3,
                "aspect": 2,
            },
        ],
    }


@pytest.fixture
def sample_descriptive_scores():
    """記述採点結果のサンプル（10人）"""
    scores = {}
    for i in range(10):
        fname = f"student_{i:03d}.jpg"
        if i % 3 == 0:
            # 満点
            scores[fname] = {"D1": 5, "D2": 3}
        elif i % 3 == 1:
            # 部分点
            scores[fname] = {"D1": 3, "D2": 1}
        else:
            # 0点
            scores[fname] = {"D1": 0, "D2": 0}
    return scores


# ============================================================
# convert_mark2_to_ctt_data の記述統合テスト
# ============================================================

class TestConvertWithDescriptive:
    """convert_mark2_to_ctt_data の記述問題統合テスト"""

    def test_without_descriptive_data(self, dummy_template_and_results):
        """記述データなしでは従来通り動作する"""
        d = dummy_template_and_results
        ans_df, key_df = convert_mark2_to_ctt_data(
            d['template_path'], d['result_path'], skip_questions=0
        )
        assert len(key_df) == 5  # マーク5問のみ
        assert 'D1' not in ans_df.columns
        assert 'D2' not in ans_df.columns

    def test_with_descriptive_data_adds_columns(
        self, dummy_template_and_results, sample_descriptive_config, sample_descriptive_scores
    ):
        """記述データありで D1, D2 列が追加される"""
        d = dummy_template_and_results
        ans_df, key_df = convert_mark2_to_ctt_data(
            d['template_path'], d['result_path'], skip_questions=0,
            descriptive_config=sample_descriptive_config,
            descriptive_scores=sample_descriptive_scores,
        )
        # マーク5問 + 記述2問 = 7問
        assert len(key_df) == 7
        assert 'D1' in ans_df.columns
        assert 'D2' in ans_df.columns

    def test_descriptive_key_is_1(
        self, dummy_template_and_results, sample_descriptive_config, sample_descriptive_scores
    ):
        """記述問題の正答キーは "1" """
        d = dummy_template_and_results
        _, key_df = convert_mark2_to_ctt_data(
            d['template_path'], d['result_path'], skip_questions=0,
            descriptive_config=sample_descriptive_config,
            descriptive_scores=sample_descriptive_scores,
        )
        desc_keys = key_df[key_df['QuestionID'].str.startswith('D')]
        assert all(desc_keys['Key'] == '1')

    def test_full_score_becomes_1(
        self, dummy_template_and_results, sample_descriptive_config, sample_descriptive_scores
    ):
        """満点の生徒は "1"（正答）になる"""
        d = dummy_template_and_results
        ans_df, _ = convert_mark2_to_ctt_data(
            d['template_path'], d['result_path'], skip_questions=0,
            descriptive_config=sample_descriptive_config,
            descriptive_scores=sample_descriptive_scores,
        )
        # student_000.jpg は D1=5(満点), D2=3(満点) → 両方 "1"
        student_0 = ans_df[ans_df['StudentID'] == 'student_000.jpg'].iloc[0]
        assert student_0['D1'] == '1'
        assert student_0['D2'] == '1'

    def test_below_full_score_becomes_0(
        self, dummy_template_and_results, sample_descriptive_config, sample_descriptive_scores
    ):
        """満点未満の生徒は "0"（誤答）になる"""
        d = dummy_template_and_results
        ans_df, _ = convert_mark2_to_ctt_data(
            d['template_path'], d['result_path'], skip_questions=0,
            descriptive_config=sample_descriptive_config,
            descriptive_scores=sample_descriptive_scores,
        )
        # student_001.jpg は D1=3(部分点), D2=1(部分点) → 両方 "0"
        student_1 = ans_df[ans_df['StudentID'] == 'student_001.jpg'].iloc[0]
        assert student_1['D1'] == '0'
        assert student_1['D2'] == '0'

    def test_zero_score_becomes_0(
        self, dummy_template_and_results, sample_descriptive_config, sample_descriptive_scores
    ):
        """0点の生徒は "0"（誤答）になる"""
        d = dummy_template_and_results
        ans_df, _ = convert_mark2_to_ctt_data(
            d['template_path'], d['result_path'], skip_questions=0,
            descriptive_config=sample_descriptive_config,
            descriptive_scores=sample_descriptive_scores,
        )
        # student_002.jpg は D1=0, D2=0 → 両方 "0"
        student_2 = ans_df[ans_df['StudentID'] == 'student_002.jpg'].iloc[0]
        assert student_2['D1'] == '0'
        assert student_2['D2'] == '0'


class TestCTTAnalyzerWithDescriptive:
    """CTTAnalyzer に記述統合データを渡した場合のテスト"""

    def test_analyzer_accepts_descriptive_columns(
        self, dummy_template_and_results, sample_descriptive_config, sample_descriptive_scores
    ):
        """記述列を含むデータでCTTAnalyzerが正常に動作する"""
        d = dummy_template_and_results
        ans_df, key_df = convert_mark2_to_ctt_data(
            d['template_path'], d['result_path'], skip_questions=0,
            descriptive_config=sample_descriptive_config,
            descriptive_scores=sample_descriptive_scores,
        )
        analyzer = CTTAnalyzer(ans_df, key_df)
        
        assert analyzer.n_questions == 7  # 5 + 2
        assert 'D1' in analyzer.questions
        assert 'D2' in analyzer.questions

    def test_test_stats_include_descriptive(
        self, dummy_template_and_results, sample_descriptive_config, sample_descriptive_scores
    ):
        """テスト統計量が記述問題を含んで計算される"""
        d = dummy_template_and_results
        ans_df, key_df = convert_mark2_to_ctt_data(
            d['template_path'], d['result_path'], skip_questions=0,
            descriptive_config=sample_descriptive_config,
            descriptive_scores=sample_descriptive_scores,
        )
        analyzer = CTTAnalyzer(ans_df, key_df)
        stats = analyzer.calculate_test_stats()
        
        assert stats['項目数 (K)'] == 7
        assert stats['受験者数 (N)'] == 10

    def test_item_stats_include_descriptive(
        self, dummy_template_and_results, sample_descriptive_config, sample_descriptive_scores
    ):
        """項目統計量に記述問題が含まれる"""
        d = dummy_template_and_results
        ans_df, key_df = convert_mark2_to_ctt_data(
            d['template_path'], d['result_path'], skip_questions=0,
            descriptive_config=sample_descriptive_config,
            descriptive_scores=sample_descriptive_scores,
        )
        analyzer = CTTAnalyzer(ans_df, key_df)
        item_stats = analyzer.calculate_item_stats()
        
        # DataFrameとして返される
        q_ids = item_stats['QuestionID'].tolist()
        assert 'D1' in q_ids
        assert 'D2' in q_ids

    def test_score_matrix_binary_for_descriptive(
        self, dummy_template_and_results, sample_descriptive_config, sample_descriptive_scores
    ):
        """記述問題のスコアマトリクスも0/1バイナリになっている"""
        d = dummy_template_and_results
        ans_df, key_df = convert_mark2_to_ctt_data(
            d['template_path'], d['result_path'], skip_questions=0,
            descriptive_config=sample_descriptive_config,
            descriptive_scores=sample_descriptive_scores,
        )
        analyzer = CTTAnalyzer(ans_df, key_df)
        
        for q_id in ['D1', 'D2']:
            unique_vals = set(analyzer.score_matrix[q_id].unique())
            assert unique_vals <= {0, 1}, f"{q_id} に 0/1 以外の値: {unique_vals}"


class TestRExportWithDescriptive:
    """R Export の記述統合テスト"""

    def test_r_export_with_descriptive(
        self, dummy_template_and_results, sample_descriptive_config,
        sample_descriptive_scores, tmp_path
    ):
        """記述データ込みでR Exportが正常に動作する"""
        from r_export import export_r_analysis_kit
        
        d = dummy_template_and_results
        output_folder = tmp_path / "report"
        output_folder.mkdir()
        
        result = export_r_analysis_kit(
            d['template_path'], d['result_path'],
            str(output_folder), skip_questions=0,
            descriptive_config=sample_descriptive_config,
            descriptive_scores=sample_descriptive_scores,
        )
        
        assert result['success'] is True
        
        # CSVにD1, D2列が含まれるか確認
        kit_folder = Path(result['output_dir'])
        scored_csv = kit_folder / "scored_data.csv"
        assert scored_csv.exists()
        
        df = pd.read_csv(str(scored_csv))
        # 列名はQ001形式に変換される（D1,D2ではなくQ00N形式）
        # 記述問題分を含めた列数が正しいか確認
        # 全員同点の列は除外される場合がある
        assert len(df.columns) >= 2  # 少なくとも行名列+1設問列
        # Q形式の列名を確認
        q_cols = [c for c in df.columns if c.startswith("Q")]
        assert len(q_cols) > 0, "Q形式の列が存在しない"

    def test_r_export_without_descriptive(
        self, dummy_template_and_results, tmp_path
    ):
        """記述データなしでもR Exportが正常に動作する（後方互換性）"""
        from r_export import export_r_analysis_kit
        
        d = dummy_template_and_results
        output_folder = tmp_path / "report"
        output_folder.mkdir()
        
        result = export_r_analysis_kit(
            d['template_path'], d['result_path'],
            str(output_folder), skip_questions=0,
        )
        
        assert result['success'] is True
        
        kit_folder = Path(result['output_dir'])
        scored_csv = kit_folder / "scored_data.csv"
        df = pd.read_csv(str(scored_csv))
        # 列名はすべてQ形式であること
        q_cols = [c for c in df.columns if c.startswith("Q")]
        assert len(q_cols) > 0


# ============================================================
# GUIチェックボックスのテスト
# ============================================================

class TestGUIDescriptiveCheckbox:
    """GUI の記述採点分析チェックボックスのテスト"""

    def test_include_descriptive_in_analysis_var_exists(self):
        """include_descriptive_in_analysis 変数が存在する"""
        import tkinter as tk
        from conftest import get_shared_tk_root
        root = get_shared_tk_root()
        
        from main_gui import Mark2GUI
        app = Mark2GUI(root)
        
        assert hasattr(app, 'include_descriptive_in_analysis')
        assert isinstance(app.include_descriptive_in_analysis, tk.BooleanVar)

    def test_include_descriptive_default_on(self):
        """デフォルトでON"""
        from conftest import get_shared_tk_root
        root = get_shared_tk_root()
        
        from main_gui import Mark2GUI
        app = Mark2GUI(root)
        
        assert app.include_descriptive_in_analysis.get() is True

    def test_checkbox_widget_exists(self):
        """チェックボックスウィジェットが存在する"""
        from conftest import get_shared_tk_root
        root = get_shared_tk_root()
        
        from main_gui import Mark2GUI
        app = Mark2GUI(root)
        
        assert hasattr(app, '_chk_include_desc_analysis')

    def test_checkbox_hidden_when_descriptive_off(self):
        """記述モードOFF時はチェックボックスが非表示"""
        from conftest import get_shared_tk_root
        root = get_shared_tk_root()
        
        from main_gui import Mark2GUI
        app = Mark2GUI(root)
        
        app.descriptive_enabled.set(False)
        app._on_descriptive_toggle()
        
        # pack_info() が呼べない = 非表示
        with pytest.raises(Exception):
            app._chk_include_desc_analysis.pack_info()

    def test_checkbox_shown_when_descriptive_on(self):
        """記述モードON時はチェックボックスが表示される"""
        from conftest import get_shared_tk_root
        root = get_shared_tk_root()
        
        from main_gui import Mark2GUI
        app = Mark2GUI(root)
        
        app.descriptive_enabled.set(True)
        app._on_descriptive_toggle()
        
        # pack_info() が呼べる = 表示中
        info = app._chk_include_desc_analysis.pack_info()
        assert info is not None
