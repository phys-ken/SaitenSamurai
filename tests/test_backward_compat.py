#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
後方互換テスト: saitensamurai.py から全シンボルがインポート可能か検証
============================================================================
リファクタリングで各モジュールに分割した後も、
既存コード（テスト含む）が ``from saitensamurai import X`` で
全機能にアクセスできることを保証する。
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "main_src"))


# ================================================================
# 全 re-export シンボルのインポート可否チェック
# ================================================================

EXPECTED_SYMBOLS = [
    # --- constants.py ---
    "safe_print",
    "extract_pdf_to_images",
    "combine_images_to_pdf",
    "HAS_PYMUPDF",
    "MARK2_WIDTH",
    "MARK2_HEIGHT",
    "RESULTS_FOLDER",
    "BOXED_FOLDER",
    "RESULTS_DATA_FOLDER",
    "SCORED_FOLDER",
    "FINAL_REPORT_FOLDER",
    "MARK_AREAS_FILE",
    "ANSWER_KEY_FILE",
    "SESSION_STATE_FILE",
    "ERROR_TYPE_NO_MARK",
    "ERROR_TYPE_DOUBLE_MARK",
    "ERROR_TYPE_INVALID",
    "DEFAULT_CORRECTION",
    "DEFAULT_SCALE_FACTOR",
    "DEFAULT_EXPAND_FACTOR",
    "DEFAULT_EXPAND_FACTOR_Y",
    "DEFAULT_BACKUP_FOLDER",
    "MAX_DISPLAY_WIDTH",
    "MAX_DISPLAY_HEIGHT",
    "MARK2_BASE_WIDTH",
    "MARK2_BASE_HEIGHT",
    # --- scoring_engine.py ---
    "number_to_circled",
    "normalize_value",
    "load_template",
    "load_mark2_results",
    "score_answers",
    # --- omr_engine.py ---
    "imread_unicode",
    "parse_excel_coordinates",
    "save_template_coordinates_debug",
    "load_coordinates_from_csv",
    "detect_corner_markers",
    "apply_perspective_transform",
    "compute_output_scale",
    "draw_all_areas",
    "generate_template",
    "save_coordinates_to_csv",
    "_save_coordinates_to_csv_impl",
    "recognize_marks",
    "save_recognition_results",
    "process_box_drawer",
    "process_folder",
    # --- threshold_calibrator.py ---
    "collect_mark_fill_ratios",
    "estimate_color_threshold_from_pixels",
    "kmeans_2class",
    "analyze_fill_ratio_distribution",
    "run_threshold_calibration",
    "reclassify_with_threshold",
    "recollect_and_reclassify",
    # --- image_renderer.py ---
    "draw_text_on_image",
    "draw_mixed_text_on_image",
    "draw_scoring_results",
    "draw_total_score",
    "_draw_total_score_in_box",
    "_draw_total_score_fallback",
    "process_scoring",
    # --- summary_generator.py ---
    "generate_student_summary",
    "generate_exam_summary",
    "process_summary_generation",
    # --- ctt_analyzer.py ---
    "convert_mark2_to_ctt_data",
    "_sort_choices",
    "CTTAnalyzer",
    "CTTPlotGenerator",
    "CTTExcelExporter",
    "CTTPDFReporter",
    "generate_ctt_analysis",
    # --- mark_checker.py ---
    "create_backup_checker",
    "update_xlsx_from_csv_checker",
    "apply_corrections_checker",
    "detect_errors_checker",
    "load_errors_checker",
    "save_errors_checker",
    "load_coordinates_csv_checker",
    "get_bbox_for_question_checker",
    "crop_and_scale_image_checker",
    "get_display_image_checker",
    "fit_image_to_display",
    "pil_to_imagetk_checker",
    "CorrectedImageCache",
    "_load_and_correct_image",
    "crop_from_corrected_image",
    # --- gui_components.py ---
    "MarkCheckerGUI",
    "StudentAnswerSheetViewer",
    "ThresholdCalibratorGUI",
    # --- main_gui.py ---
    "Mark2GUI",
    # --- フラグ ---
    "HAS_MATPLOTLIB",
    "HAS_REPORTLAB",
]


@pytest.mark.parametrize("symbol", EXPECTED_SYMBOLS)
def test_symbol_importable(symbol):
    """saitensamurai から各シンボルがインポート可能"""
    import saitensamurai
    assert hasattr(saitensamurai, symbol), (
        f"saitensamurai に '{symbol}' が存在しません — "
        f"re-export が漏れています"
    )


def test_no_extra_missing_constants():
    """定数モジュール由来の主要定数が正しく利用可能"""
    from saitensamurai import (
        RESULTS_FOLDER,
        RESULTS_DATA_FOLDER,
        SESSION_STATE_FILE,
    )
    assert RESULTS_FOLDER == "_saiten_grading_results"
    assert RESULTS_DATA_FOLDER == "01_Results"
    assert SESSION_STATE_FILE == "session_state.json"


def test_mark2gui_is_class():
    """Mark2GUI がクラスとしてインポートされる"""
    from saitensamurai import Mark2GUI
    assert isinstance(Mark2GUI, type)
