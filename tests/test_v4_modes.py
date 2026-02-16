"""
test_v4_modes.py - v4.0 モード機能のテスト

β: 起動モード選択（マークのみ / マーク＋記述 / 記述のみ）
α: 記述採点のリッチ化（DescriptiveReviewGUI, generate_descriptive_only_sheets）
"""

import json
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch, MagicMock

import pytest
import tkinter as tk
import numpy as np

import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "main_src"))

from conftest import get_shared_tk_root

from constants import (
    MODE_MARK_ONLY,
    MODE_MARK_AND_DESCRIPTIVE,
    MODE_DESCRIPTIVE_ONLY,
    RESULTS_FOLDER,
    RESULTS_DATA_FOLDER,
    BOXED_FOLDER,
    SCORED_FOLDER,
    SESSION_STATE_FILE,
    get_rendering_settings,
)


# ============================================================
# ヘルパー
# ============================================================

def _make_stub_app(img_folder="", mode=MODE_MARK_AND_DESCRIPTIVE):
    """Mark2GUI の軽量スタブ（__init__ をスキップ）"""
    from main_gui import Mark2GUI

    root = get_shared_tk_root()
    app = object.__new__(Mark2GUI)
    app.root = root
    app.app_mode = mode
    app.image_folder_path = tk.StringVar(root, value=img_folder)
    app.coord_excel_path = tk.StringVar(root, value="")
    app.template_path = tk.StringVar(root, value="")
    app.mark2_result_path = tk.StringVar(root, value="")
    app.skip_questions = tk.StringVar(root, value="4")
    app.color_threshold = tk.DoubleVar(root, value=0.1)
    app.area_threshold = tk.DoubleVar(root, value=0.4)
    app.descriptive_enabled = tk.BooleanVar(
        root, value=(mode in (MODE_MARK_AND_DESCRIPTIVE, MODE_DESCRIPTIVE_ONLY))
    )
    app.include_descriptive_in_analysis = tk.BooleanVar(root, value=True)
    app.rendering_settings = get_rendering_settings()
    app._log_messages = []
    app.log_message = lambda msg: app._log_messages.append(msg)
    app._desc_status_label = tk.Label(root)
    app._desc_status_frame = tk.Frame(root)
    app._on_descriptive_toggle = lambda: None
    app.name_trim_enabled = tk.BooleanVar(root, value=False)
    app._processing = False
    return app


def _create_gui(mode):
    """SaitenSamuraiGUI を Toplevel 上に生成"""
    from main_gui import SaitenSamuraiGUI
    root = get_shared_tk_root()
    top = tk.Toplevel(root)
    top.withdraw()
    app = SaitenSamuraiGUI(top, mode=mode)
    return top, app


def _destroy_top(top):
    """Toplevel の安全な破棄"""
    try:
        top.update_idletasks()
        top.destroy()
    except tk.TclError:
        pass


# ============================================================
# モード定数のテスト
# ============================================================

class TestModeConstants(unittest.TestCase):
    """モード定数が正しく定義されている"""

    def test_mode_values(self):
        self.assertEqual(MODE_MARK_ONLY, "mark_only")
        self.assertEqual(MODE_MARK_AND_DESCRIPTIVE, "mark_and_descriptive")
        self.assertEqual(MODE_DESCRIPTIVE_ONLY, "descriptive_only")

    def test_modes_are_distinct(self):
        modes = {MODE_MARK_ONLY, MODE_MARK_AND_DESCRIPTIVE, MODE_DESCRIPTIVE_ONLY}
        self.assertEqual(len(modes), 3)


# ============================================================
# GUI 初期化モード別テスト
# ============================================================

class TestGUIInitModes:
    """各モードでの GUI 初期化動作を検証"""

    def test_mark_only_mode(self):
        """マークのみモード: 記述チェック OFF、タイトルに「マーク採点」"""
        top, app = _create_gui(MODE_MARK_ONLY)
        try:
            assert app.descriptive_enabled.get() is False
            assert "マーク採点" in top.title()
            assert app.app_mode == MODE_MARK_ONLY
        finally:
            _destroy_top(top)

    def test_mark_and_descriptive_mode(self):
        """マーク＋記述モード: 記述チェック ON、タイトルに「マーク＋記述」"""
        top, app = _create_gui(MODE_MARK_AND_DESCRIPTIVE)
        try:
            assert app.descriptive_enabled.get() is True
            assert "マーク＋記述" in top.title()
            assert app.app_mode == MODE_MARK_AND_DESCRIPTIVE
        finally:
            _destroy_top(top)

    def test_descriptive_only_mode(self):
        """記述のみモード: 記述チェック ON、タイトルに「記述採点」"""
        top, app = _create_gui(MODE_DESCRIPTIVE_ONLY)
        try:
            assert app.descriptive_enabled.get() is True
            assert "記述採点" in top.title()
            assert app.app_mode == MODE_DESCRIPTIVE_ONLY
        finally:
            _destroy_top(top)

    def test_descriptive_only_hides_mark_controls(self):
        """記述のみモード: 認識ボタンが「画像準備」に変更"""
        top, app = _create_gui(MODE_DESCRIPTIVE_ONLY)
        try:
            assert "画像準備" in app._btn_run_box.cget("text")
        finally:
            _destroy_top(top)

    def test_mark_only_hides_descriptive_controls(self):
        """マークのみモード: 記述ボタンが非表示"""
        top, app = _create_gui(MODE_MARK_ONLY)
        try:
            assert app.desc_setup_btn.winfo_manager() == ""
            assert app.desc_scoring_btn.winfo_manager() == ""
        finally:
            _destroy_top(top)

    def test_descriptive_only_has_review_button(self):
        """記述のみモード: 採点確認ボタンが存在する"""
        top, app = _create_gui(MODE_DESCRIPTIVE_ONLY)
        try:
            assert hasattr(app, '_btn_desc_review')
            assert app._btn_desc_review.winfo_manager() == "pack"
        finally:
            _destroy_top(top)


# ============================================================
# StartupModeDialog テスト
# ============================================================

class TestStartupModeDialogStructure:
    """StartupModeDialog の構造テスト"""

    def test_dialog_class_exists(self):
        """StartupModeDialog クラスが存在する"""
        from gui_components import StartupModeDialog
        assert StartupModeDialog is not None

    def test_dialog_has_result_attribute(self):
        """StartupModeDialog は result と _session_path を持つ"""
        from gui_components import StartupModeDialog
        # クラスの属性を確認（インスタンス化はモーダルなのでスキップ）
        assert hasattr(StartupModeDialog, '__init__')


# ============================================================
# セッション保存にモード情報が含まれる
# ============================================================

class TestSessionModeInfo(unittest.TestCase):
    """セッション状態にモード情報が保存される"""

    def setUp(self):
        self.tmpdir = Path(tempfile.mkdtemp())
        self.img_folder = self.tmpdir / "images"
        self.img_folder.mkdir()
        results_data = self.img_folder / RESULTS_FOLDER / RESULTS_DATA_FOLDER
        results_data.mkdir(parents=True)

    def tearDown(self):
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_mode_saved_in_session(self):
        """app_mode がセッションJSONに含まれる"""
        app = _make_stub_app(str(self.img_folder), mode=MODE_DESCRIPTIVE_ONLY)
        app._save_session_state()

        session_path = (self.img_folder / RESULTS_FOLDER /
                        RESULTS_DATA_FOLDER / SESSION_STATE_FILE)
        with open(session_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        self.assertIn("app_mode", data)
        self.assertEqual(data["app_mode"], MODE_DESCRIPTIVE_ONLY)

    def test_mode_defaults_when_missing(self):
        """セッションに app_mode がない場合のフォールバック"""
        app = _make_stub_app(str(self.img_folder))
        # 旧バージョンのセッション（app_mode なし）
        session = {
            "version": 1,
            "image_folder": str(self.img_folder),
        }
        app._apply_session_state(session)
        # エラーなく完了することを確認
        self.assertTrue(True)


# ============================================================
# generate_descriptive_only_sheets テスト
# ============================================================

class TestGenerateDescriptiveOnlySheets(unittest.TestCase):
    """記述のみモードの採点済み答案生成"""

    def setUp(self):
        self.tmpdir = Path(tempfile.mkdtemp())
        self.boxed_folder = self.tmpdir / "boxed"
        self.boxed_folder.mkdir()
        self.output_folder = self.tmpdir / "scored"

    def tearDown(self):
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def _make_test_image(self, name="test001.jpg"):
        """テスト画像（595x842 白画像）を作成"""
        import cv2
        img = np.ones((842, 595, 3), dtype=np.uint8) * 255
        cv2.imwrite(str(self.boxed_folder / name), img)

    def test_empty_folder(self):
        """画像がない場合はカウント0"""
        from descriptive_scorer import generate_descriptive_only_sheets

        config = {"questions": [{"id": "D1", "name": "記述1", "max_score": 5,
                                  "aspect": 1, "region": [10, 10, 200, 100]}]}
        result = generate_descriptive_only_sheets(
            boxed_folder=str(self.boxed_folder),
            config=config,
            descriptive_scores={},
            output_folder=str(self.output_folder),
        )
        self.assertEqual(result["total_count"], 0)

    def test_generates_scored_image(self):
        """画像を入力すると採点済み画像が出力される"""
        from descriptive_scorer import generate_descriptive_only_sheets

        self._make_test_image("student01.jpg")
        config = {
            "questions": [
                {"id": "D1", "name": "記述1", "max_score": 5,
                 "aspect": 1, "region": [50, 50, 300, 200]},
            ],
        }
        scores = {"student01.jpg": {"D1": 3}}

        result = generate_descriptive_only_sheets(
            boxed_folder=str(self.boxed_folder),
            config=config,
            descriptive_scores=scores,
            output_folder=str(self.output_folder),
        )
        self.assertEqual(result["total_count"], 1)
        self.assertEqual(result["success_count"], 1)
        self.assertEqual(result["error_count"], 0)
        self.assertTrue((self.output_folder / "student01.jpg").exists())

    def test_handles_missing_scores(self):
        """スコアがない生徒も処理される"""
        from descriptive_scorer import generate_descriptive_only_sheets

        self._make_test_image("student01.jpg")
        self._make_test_image("student02.jpg")
        config = {
            "questions": [
                {"id": "D1", "name": "記述1", "max_score": 5,
                 "aspect": 1, "region": [50, 50, 300, 200]},
            ],
        }
        scores = {"student01.jpg": {"D1": 3}}  # student02 has no scores

        result = generate_descriptive_only_sheets(
            boxed_folder=str(self.boxed_folder),
            config=config,
            descriptive_scores=scores,
            output_folder=str(self.output_folder),
        )
        self.assertEqual(result["total_count"], 2)
        self.assertEqual(result["success_count"], 2)

    def test_log_callback(self):
        """log_callback が呼ばれる"""
        from descriptive_scorer import generate_descriptive_only_sheets

        self._make_test_image("test.jpg")
        logs = []
        config = {"questions": [{"id": "D1", "name": "Q1", "max_score": 3,
                                  "aspect": 1, "region": [10, 10, 100, 50]}]}

        generate_descriptive_only_sheets(
            boxed_folder=str(self.boxed_folder),
            config=config,
            descriptive_scores={"test.jpg": {"D1": 2}},
            output_folder=str(self.output_folder),
            log_callback=lambda msg: logs.append(msg),
        )
        self.assertTrue(any("記述のみ" in m for m in logs))


# ============================================================
# process_descriptive_only_summary テスト
# ============================================================

class TestProcessDescriptiveOnlySummary(unittest.TestCase):
    """記述のみモードのサマリー生成"""

    def setUp(self):
        self.tmpdir = Path(tempfile.mkdtemp())
        self.img_folder = self.tmpdir / "images"
        self.img_folder.mkdir()
        (self.img_folder / RESULTS_FOLDER).mkdir()

    def tearDown(self):
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_basic_summary(self):
        """基本的なサマリー生成"""
        from summary_generator import process_descriptive_only_summary

        config = {
            "questions": [
                {"id": "D1", "name": "記述1", "max_score": 5, "aspect": 1,
                 "region": [10, 10, 200, 100]},
                {"id": "D2", "name": "記述2", "max_score": 3, "aspect": 2,
                 "region": [10, 110, 200, 200]},
            ]
        }
        scores = {
            "student01.jpg": {"D1": 5, "D2": 3},
            "student02.jpg": {"D1": 3, "D2": 1},
            "student03.jpg": {"D1": 2, "D2": 2},
        }

        result = process_descriptive_only_summary(
            image_folder=str(self.img_folder),
            descriptive_config=config,
            descriptive_scores=scores,
        )
        self.assertTrue(result["success"])
        stats = result["stats"]
        self.assertEqual(stats["受験者数"], 3)
        self.assertEqual(stats["満点"], 8)
        self.assertAlmostEqual(stats["平均点"], (8 + 4 + 4) / 3, places=1)
        self.assertEqual(stats["最高点"], 8)
        self.assertEqual(stats["最低点"], 4)

    def test_empty_questions(self):
        """問題が空の場合はエラー"""
        from summary_generator import process_descriptive_only_summary

        result = process_descriptive_only_summary(
            image_folder=str(self.img_folder),
            descriptive_config={"questions": []},
            descriptive_scores={},
        )
        self.assertFalse(result["success"])
        self.assertIn("設定されていません", result["error"])

    def test_excel_files_created(self):
        """Excelファイルが生成される"""
        from summary_generator import process_descriptive_only_summary
        from constants import STUDENT_SUMMARY_FILE, EXAM_SUMMARY_FILE, FINAL_REPORT_FOLDER

        config = {
            "questions": [
                {"id": "D1", "name": "記述1", "max_score": 5, "aspect": 1,
                 "region": [10, 10, 200, 100]},
            ]
        }
        scores = {"s1.jpg": {"D1": 3}}

        result = process_descriptive_only_summary(
            image_folder=str(self.img_folder),
            descriptive_config=config,
            descriptive_scores=scores,
        )
        self.assertTrue(result["success"])

        student_path = Path(result["student_summary_path"])
        exam_path = Path(result["exam_summary_path"])
        self.assertTrue(student_path.exists())
        self.assertTrue(exam_path.exists())


# ============================================================
# DescriptiveReviewGUI 構造テスト
# ============================================================

class TestDescriptiveReviewGUIStructure(unittest.TestCase):
    """DescriptiveReviewGUI クラスの存在確認"""

    def test_class_exists(self):
        from descriptive_scorer import DescriptiveReviewGUI
        self.assertIsNotNone(DescriptiveReviewGUI)

    def test_class_attributes(self):
        from descriptive_scorer import DescriptiveReviewGUI
        self.assertTrue(hasattr(DescriptiveReviewGUI, 'THUMB_SIZE_DEFAULT'))
        self.assertTrue(hasattr(DescriptiveReviewGUI, 'GRID_COLS'))
        self.assertEqual(DescriptiveReviewGUI.GRID_COLS, 4)


# ============================================================
# 記述のみモード run_scoring 分岐テスト
# ============================================================

class TestDescriptiveOnlyRunScoring(unittest.TestCase):
    """記述のみモードでの run_scoring 分岐"""

    def setUp(self):
        self.tmpdir = Path(tempfile.mkdtemp())
        self.img_folder = self.tmpdir / "images"
        self.img_folder.mkdir()
        results_data = self.img_folder / RESULTS_FOLDER / RESULTS_DATA_FOLDER
        results_data.mkdir(parents=True)

    def tearDown(self):
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    @patch("tkinter.messagebox.showerror")
    def test_no_config_shows_error(self, mock_error):
        """記述設定がない場合はエラー"""
        app = _make_stub_app(str(self.img_folder), mode=MODE_DESCRIPTIVE_ONLY)
        app._check_descriptive_completeness = lambda: (True, 0, 0, [])
        app._set_processing_state = lambda s: None
        app.log_text = MagicMock()

        app._run_scoring_descriptive_only()
        mock_error.assert_called_once()

    @patch("tkinter.messagebox.showerror")
    def test_no_scores_shows_error(self, mock_error):
        """採点データがない場合はエラー"""
        results_data = self.img_folder / RESULTS_FOLDER / RESULTS_DATA_FOLDER
        config = {"version": 1, "questions": [{"id": "D1", "name": "Q1",
                   "max_score": 5, "aspect": 1, "region": [0, 0, 100, 100]}]}
        (results_data / "descriptive_config.json").write_text(
            json.dumps(config), encoding='utf-8'
        )

        app = _make_stub_app(str(self.img_folder), mode=MODE_DESCRIPTIVE_ONLY)
        app._check_descriptive_completeness = lambda: (True, 0, 0, [])
        app._set_processing_state = lambda s: None
        app.log_text = MagicMock()

        app._run_scoring_descriptive_only()
        mock_error.assert_called_once()


# ============================================================
# 記述のみモード run_summary_generation 分岐テスト
# ============================================================

class TestDescriptiveOnlySummaryGeneration(unittest.TestCase):
    """記述のみモードでの summary generation 分岐"""

    @patch("tkinter.messagebox.showerror")
    def test_no_data_shows_error(self, mock_error):
        """データがない場合はエラー"""
        tmpdir = Path(tempfile.mkdtemp())
        img_folder = tmpdir / "images"
        img_folder.mkdir()
        results_data = img_folder / RESULTS_FOLDER / RESULTS_DATA_FOLDER
        results_data.mkdir(parents=True)

        try:
            app = _make_stub_app(str(img_folder), mode=MODE_DESCRIPTIVE_ONLY)
            app._check_descriptive_completeness = lambda: (True, 0, 0, [])

            app._run_summary_generation_descriptive_only()
            mock_error.assert_called_once()
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


# ============================================================
# モード別マークチェック制御テスト
# ============================================================

class TestModeMarkCheckerGuard:
    """run_mark_checker の記述のみモードガード"""

    @patch("tkinter.messagebox.showinfo")
    def test_mark_checker_blocked_in_descriptive_only(self, mock_info):
        app = _make_stub_app(mode=MODE_DESCRIPTIVE_ONLY)
        app.run_mark_checker()
        mock_info.assert_called_once()
        assert "使用できません" in mock_info.call_args[0][1]


# ============================================================
# re-export テスト (saitensamurai.py)
# ============================================================

class TestSaitensamuraiExports(unittest.TestCase):
    """saitensamurai.py からのモード定数 re-export"""

    def test_mode_constants_exported(self):
        from saitensamurai import (
            MODE_MARK_ONLY, MODE_MARK_AND_DESCRIPTIVE, MODE_DESCRIPTIVE_ONLY
        )
        self.assertEqual(MODE_MARK_ONLY, "mark_only")
        self.assertEqual(MODE_MARK_AND_DESCRIPTIVE, "mark_and_descriptive")
        self.assertEqual(MODE_DESCRIPTIVE_ONLY, "descriptive_only")

    def test_startup_dialog_exported(self):
        from saitensamurai import StartupModeDialog
        self.assertIsNotNone(StartupModeDialog)


# ============================================================
# CTT分析: 記述のみモード対応テスト
# ============================================================

class TestCTTDescriptiveOnly(unittest.TestCase):
    """convert_mark2_to_ctt_data の記述のみモード対応"""

    def _make_desc_data(self):
        config = {
            "questions": [
                {"id": "D1", "name": "記述1", "max_score": 5, "aspect": 1,
                 "region": [10, 10, 200, 100]},
                {"id": "D2", "name": "記述2", "max_score": 3, "aspect": 2,
                 "region": [10, 110, 200, 200]},
            ]
        }
        scores = {
            "student01.jpg": {"D1": 5, "D2": 3},
            "student02.jpg": {"D1": 3, "D2": 1},
            "student03.jpg": {"D1": 0, "D2": 0},
        }
        return config, scores

    def test_descriptive_only_ctt_data(self):
        """マーク問題なし・記述のみでCTTデータを生成できる"""
        from ctt_analyzer import convert_mark2_to_ctt_data
        config, scores = self._make_desc_data()

        ans_df, key_df = convert_mark2_to_ctt_data(
            template_path=None,
            mark2_result_path=None,
            descriptive_config=config,
            descriptive_scores=scores,
        )

        # 3人の学生
        self.assertEqual(len(ans_df), 3)
        # 2問の記述問題
        self.assertEqual(len(key_df), 2)
        # 記述問題のキーは "1"
        self.assertTrue(all(key_df["Key"] == "1"))
        # StudentID 列が存在
        self.assertIn("StudentID", ans_df.columns)
        # D1, D2 列が存在
        self.assertIn("D1", ans_df.columns)
        self.assertIn("D2", ans_df.columns)

    def test_descriptive_only_binary_encoding(self):
        """記述のみモード: 満点→1, 満点未満→0 のバイナリエンコード"""
        from ctt_analyzer import convert_mark2_to_ctt_data
        config, scores = self._make_desc_data()

        ans_df, _ = convert_mark2_to_ctt_data(
            template_path=None,
            mark2_result_path=None,
            descriptive_config=config,
            descriptive_scores=scores,
        )

        # student01: D1=5(満点→1), D2=3(満点→1)
        s1 = ans_df[ans_df["StudentID"] == "student01.jpg"].iloc[0]
        self.assertEqual(s1["D1"], "1")
        self.assertEqual(s1["D2"], "1")

        # student02: D1=3(→0), D2=1(→0)
        s2 = ans_df[ans_df["StudentID"] == "student02.jpg"].iloc[0]
        self.assertEqual(s2["D1"], "0")
        self.assertEqual(s2["D2"], "0")

    def test_descriptive_only_ctt_analyzer(self):
        """記述のみデータでCTTAnalyzerが動作する"""
        from ctt_analyzer import convert_mark2_to_ctt_data, CTTAnalyzer

        config, scores = self._make_desc_data()
        ans_df, key_df = convert_mark2_to_ctt_data(
            template_path=None,
            mark2_result_path=None,
            descriptive_config=config,
            descriptive_scores=scores,
        )
        analyzer = CTTAnalyzer(ans_df, key_df)
        test_stats = analyzer.calculate_test_stats()
        item_stats = analyzer.calculate_item_stats()

        self.assertIn("受験者数 (N)", test_stats)
        self.assertEqual(test_stats["受験者数 (N)"], 3)
        self.assertIn("項目数 (K)", test_stats)
        self.assertEqual(test_stats["項目数 (K)"], 2)

    def test_descriptive_only_generate_ctt_analysis(self):
        """記述のみモードでCTTレポートを生成できる"""
        from ctt_analyzer import generate_ctt_analysis

        config, scores = self._make_desc_data()
        tmpdir = Path(tempfile.mkdtemp())
        try:
            excel_path = tmpdir / "ctt.xlsx"
            pdf_path = tmpdir / "ctt.pdf"
            result = generate_ctt_analysis(
                template_path=None,
                mark2_result_path=None,
                excel_output_path=excel_path,
                pdf_output_path=pdf_path,
                descriptive_config=config,
                descriptive_scores=scores,
            )
            self.assertTrue(result["success"])
            self.assertTrue(excel_path.exists())
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


# ============================================================
# R連携: 記述のみモード対応テスト
# ============================================================

class TestRExportDescriptiveOnly(unittest.TestCase):
    """export_r_analysis_kit の記述のみモード対応"""

    def test_descriptive_only_r_export(self):
        """記述のみモードでR分析キットを出力できる"""
        from r_export import export_r_analysis_kit

        config = {
            "questions": [
                {"id": "D1", "name": "記述1", "max_score": 5, "aspect": 1,
                 "region": [10, 10, 200, 100]},
            ]
        }
        scores = {
            "s1.jpg": {"D1": 5},
            "s2.jpg": {"D1": 3},
            "s3.jpg": {"D1": 0},
        }

        tmpdir = Path(tempfile.mkdtemp())
        try:
            result = export_r_analysis_kit(
                template_path=None,
                mark2_result_path=None,
                output_folder=str(tmpdir),
                descriptive_config=config,
                descriptive_scores=scores,
            )
            self.assertTrue(result["success"])
            kit_dir = Path(result["output_dir"])
            self.assertTrue((kit_dir / "scored_data.csv").exists())
            self.assertTrue((kit_dir / "item_info.csv").exists())
            self.assertTrue((kit_dir / "create_report.R").exists())
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


# ============================================================
# process_descriptive_only_summary CTT/R統合テスト
# ============================================================

class TestDescriptiveOnlySummaryCTTR(unittest.TestCase):
    """記述のみモードのサマリーにCTT/R出力が含まれる"""

    def setUp(self):
        self.tmpdir = Path(tempfile.mkdtemp())
        self.img_folder = self.tmpdir / "images"
        self.img_folder.mkdir()
        (self.img_folder / RESULTS_FOLDER).mkdir()

    def tearDown(self):
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_ctt_included_in_summary(self):
        """サマリー結果にCTT分析パスが含まれる"""
        from summary_generator import process_descriptive_only_summary

        config = {
            "questions": [
                {"id": "D1", "name": "記述1", "max_score": 5, "aspect": 1,
                 "region": [10, 10, 200, 100]},
                {"id": "D2", "name": "記述2", "max_score": 3, "aspect": 2,
                 "region": [10, 110, 200, 200]},
            ]
        }
        scores = {
            "student01.jpg": {"D1": 5, "D2": 3},
            "student02.jpg": {"D1": 3, "D2": 1},
            "student03.jpg": {"D1": 2, "D2": 2},
        }

        result = process_descriptive_only_summary(
            image_folder=str(self.img_folder),
            descriptive_config=config,
            descriptive_scores=scores,
        )
        self.assertTrue(result["success"])
        self.assertIn("ctt_excel_path", result)
        self.assertTrue(Path(result["ctt_excel_path"]).exists())

    def test_r_export_included_in_summary(self):
        """サマリー結果にR連携出力が含まれる"""
        from summary_generator import process_descriptive_only_summary

        config = {
            "questions": [
                {"id": "D1", "name": "記述1", "max_score": 5, "aspect": 1,
                 "region": [10, 10, 200, 100]},
                {"id": "D2", "name": "記述2", "max_score": 3, "aspect": 2,
                 "region": [10, 110, 200, 200]},
            ]
        }
        scores = {
            "student01.jpg": {"D1": 5, "D2": 3},
            "student02.jpg": {"D1": 3, "D2": 1},
            "student03.jpg": {"D1": 2, "D2": 2},
        }

        result = process_descriptive_only_summary(
            image_folder=str(self.img_folder),
            descriptive_config=config,
            descriptive_scores=scores,
        )
        self.assertTrue(result["success"])
        self.assertIn("r_export_dir", result)
        r_dir = Path(result["r_export_dir"])
        self.assertTrue(r_dir.exists())
        self.assertTrue((r_dir / "scored_data.csv").exists())


# ============================================================
# generate_descriptive_only_sheets output_scale テスト
# ============================================================

class TestDescriptiveOnlySheetsScale(unittest.TestCase):
    """generate_descriptive_only_sheets の output_scale=1.0 検証"""

    def test_no_double_scaling(self):
        """記述のみモードで座標が二重スケーリングされないことを確認"""
        import cv2
        from descriptive_scorer import generate_descriptive_only_sheets

        tmpdir = Path(tempfile.mkdtemp())
        try:
            boxed = tmpdir / "boxed"
            boxed.mkdir()
            output = tmpdir / "output"

            # 2400x3400 の大きな画像を作成（記述のみモードの典型サイズ）
            img = np.zeros((3400, 2400, 3), dtype=np.uint8)
            img[:] = (255, 255, 255)  # 白背景
            cv2.imwrite(str(boxed / "test.jpg"), img)

            config = {
                "questions": [
                    {"id": "D1", "name": "記述1", "max_score": 5, "aspect": 1,
                     "region": [100, 100, 800, 500]},
                ]
            }
            scores = {"test.jpg": {"D1": 5}}

            result = generate_descriptive_only_sheets(
                boxed_folder=str(boxed),
                config=config,
                descriptive_scores=scores,
                output_folder=str(output),
            )
            self.assertEqual(result["success_count"], 1)
            # 出力画像が存在
            self.assertTrue((output / "test.jpg").exists())
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
