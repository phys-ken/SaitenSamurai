#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUIウィジェットテスト — Mark2GUI の自動化可能なGUI検証
======================================================

tkinter_test_vision.md の方針に基づき、
実際の Mark2GUI インスタンスを生成して以下を自動テスト:

1. 初期ウィジェット生成: ボタン・ラベル・入力欄が存在するか
2. 初期状態: テキスト・state 属性が期待通りか
3. 入力バリデーション: 各アクションの事前チェックが正しく動作するか
4. ボタン操作: invoke() で呼び出したときの状態遷移
5. 自動検出ロジック: auto_detect_template, log_message 等

Tkinter が利用できない環境では全テストを自動スキップする。
"""

import sys
import os
import json
import tempfile
import shutil
from pathlib import Path
from unittest.mock import patch, MagicMock, call

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "main_src"))

# Tk が使えない環境では全体スキップ
try:
    import tkinter as tk
    _root_test = tk.Tk()
    _root_test.withdraw()
    _root_test.destroy()
    _root_test = None
    HAS_TK = True
except Exception:
    HAS_TK = False

pytestmark = pytest.mark.skipif(not HAS_TK, reason="Tkinter not available")


# ================================================================
# テスト用ヘルパー
# ================================================================

def _make_gui():
    """Mark2GUI を Toplevel 上に生成して (top, app) を返す

    共有 Tk ルート上に Toplevel を作成し、テスト終了時に
    Toplevel のみ destroy する（Tcl インタプリタは残る）。
    """
    from conftest import get_shared_tk_root
    from main_gui import Mark2GUI
    root = get_shared_tk_root()
    top = tk.Toplevel(root)
    app = Mark2GUI(top)
    top.update_idletasks()
    return top, app


def _destroy_gui(top):
    """Toplevel を安全に破棄"""
    try:
        top.destroy()
    except Exception:
        pass


# ================================================================
# 1. 初期状態テスト — ウィジェットの存在と属性
# ================================================================

class TestInitialState:
    """Mark2GUI の初期生成直後の状態を検証"""

    def setup_method(self):
        self.root, self.app = _make_gui()

    def teardown_method(self):
        _destroy_gui(self.root)

    def test_window_title(self):
        """ウィンドウタイトルが設定されている"""
        title = self.root.title()
        assert "採点侍" in title
        assert "v4.1" in title

    def test_window_geometry(self):
        """初期サイズが設定されている"""
        # geometry は "WxH+X+Y" 形式
        geom = self.root.geometry()
        # CI 環境ではウィンドウマネージャが要求サイズを縮小する場合がある
        w, rest = geom.split("x", 1)
        h = rest.split("+")[0]
        assert int(w) >= 1000, f"幅が小さすぎます: {w}"
        assert int(h) >= 500, f"高さが小さすぎます: {h}"

    def test_main_action_buttons_exist(self):
        """主要アクションボタンが存在する"""
        assert hasattr(self.app, '_btn_run_box')
        assert hasattr(self.app, '_btn_mark_check')
        assert hasattr(self.app, '_btn_total_pos')
        assert hasattr(self.app, '_btn_run_scoring')
        assert hasattr(self.app, '_btn_run_summary')

    def test_button_labels(self):
        """ボタンのテキストが期待通り"""
        assert "認識実行" in self.app._btn_run_box["text"]
        assert "マークチェック" in self.app._btn_mark_check["text"]
        assert "合計点位置" in self.app._btn_total_pos["text"]
        assert "採点済み答案" in self.app._btn_run_scoring["text"]
        assert "集計" in self.app._btn_run_summary["text"]

    def test_buttons_initially_normal(self):
        """Step 1 ボタンはフォルダ/座標ファイル未設定時 disabled、Step 2/3 も disabled（Step 進行ガード）"""
        # Step 1 は画像フォルダ＋座標ファイル未設定なので disabled
        assert str(self.app._btn_run_box["state"]) == "disabled", "Step1 OMRボタンが disabled でない"

        # Step 2 ボタンはフォルダ未選択時 disabled
        step2_buttons = [
            self.app._btn_mark_check,
            self.app._btn_total_pos,
            self.app._btn_run_scoring,
        ]
        for btn in step2_buttons:
            assert str(btn["state"]) == "disabled", f"{btn['text']} should be disabled before Step 1"

        # Step 3 集計ボタンも初期 disabled
        assert str(self.app._btn_run_summary["state"]) == "disabled", "集計ボタンが disabled でない"

    def test_folder_buttons_initially_disabled(self):
        """📁 ボタン（枠結果・採点結果・集計結果）は初期 disabled"""
        assert str(self.app.open_boxed_btn["state"]) == "disabled"
        assert str(self.app.open_scored_btn["state"]) == "disabled"
        assert str(self.app.open_results_btn["state"]) == "disabled"

    def test_input_variables_empty(self):
        """入力変数が初期状態で空"""
        assert self.app.image_folder_path.get() == ""
        assert self.app.coord_excel_path.get() == ""
        assert self.app.template_path.get() == ""
        assert self.app.mark2_result_path.get() == ""

    def test_default_option_values(self):
        """オプションのデフォルト値が正しい"""
        assert self.app.skip_questions.get() == "4"
        assert self.app.color_threshold.get() == pytest.approx(0.1)
        assert self.app.area_threshold.get() == pytest.approx(0.4)
        # v4.0: デフォルトモード(マーク＋記述)では記述が有効
        assert self.app.descriptive_enabled.get() is True

    def test_descriptive_buttons_visible_by_default(self):
        """v4.0: デフォルトモード(マーク＋記述)では記述ボタンが表示"""
        assert self.app.desc_setup_btn.winfo_manager() == "pack"
        assert self.app.desc_scoring_btn.winfo_manager() == "pack"
        assert self.app._desc_status_frame.winfo_manager() == "pack"

    def test_log_text_exists(self):
        """ログテキストウィジェットが存在する"""
        assert hasattr(self.app, 'log_text')
        # 初期状態で disabled (読み取り専用)
        assert str(self.app.log_text["state"]) == "disabled"

    def test_processing_flag_false(self):
        """初期状態で処理中フラグが False"""
        assert self.app._processing is False

    def test_checkbutton_descriptive_exists(self):
        """記述チェックボックスが存在し normal 状態"""
        assert hasattr(self.app, '_chk_descriptive')
        assert str(self.app._chk_descriptive["state"]) == "normal"


# ================================================================
# 2. validate_inputs — 入力バリデーション
# ================================================================

class TestValidateInputs:
    """validate_inputs の全分岐を検証"""

    def setup_method(self):
        self.root, self.app = _make_gui()

    def teardown_method(self):
        _destroy_gui(self.root)

    @patch("main_gui.messagebox")
    def test_no_image_folder(self, mock_mb):
        """画像フォルダ未設定 → False"""
        result = self.app.validate_inputs()
        assert result is False
        mock_mb.showerror.assert_called_once()
        assert "画像フォルダ" in mock_mb.showerror.call_args[0][1]

    @patch("main_gui.messagebox")
    def test_no_coord_excel(self, mock_mb):
        """座標ファイル未設定 → False"""
        self.app.image_folder_path.set("/some/path")
        result = self.app.validate_inputs()
        assert result is False
        mock_mb.showerror.assert_called_once()
        assert "座標ファイル" in mock_mb.showerror.call_args[0][1]

    @patch("main_gui.messagebox")
    def test_image_folder_not_exist(self, mock_mb):
        """画像フォルダが存在しない → False"""
        self.app.image_folder_path.set("/nonexistent/folder")
        self.app.coord_excel_path.set("/nonexistent/file.xlsx")
        result = self.app.validate_inputs()
        assert result is False
        mock_mb.showerror.assert_called_once()
        assert "存在しません" in mock_mb.showerror.call_args[0][1]

    @patch("main_gui.messagebox")
    def test_invalid_skip_questions(self, mock_mb):
        """skip_questions が不正 → False"""
        tmpdir = tempfile.mkdtemp()
        coord_file = Path(tmpdir) / "coord.xlsx"
        coord_file.touch()
        try:
            self.app.image_folder_path.set(tmpdir)
            self.app.coord_excel_path.set(str(coord_file))
            self.app.skip_questions.set("abc")
            result = self.app.validate_inputs()
            assert result is False
            mock_mb.showerror.assert_called_once()
            assert "整数" in mock_mb.showerror.call_args[0][1]
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    @patch("main_gui.messagebox")
    def test_negative_skip_questions(self, mock_mb):
        """skip_questions が負 → False"""
        tmpdir = tempfile.mkdtemp()
        coord_file = Path(tmpdir) / "coord.xlsx"
        coord_file.touch()
        try:
            self.app.image_folder_path.set(tmpdir)
            self.app.coord_excel_path.set(str(coord_file))
            self.app.skip_questions.set("-1")
            result = self.app.validate_inputs()
            assert result is False
            mock_mb.showerror.assert_called_once()
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    @patch("main_gui.messagebox")
    def test_valid_inputs(self, mock_mb):
        """全項目が有効 → True"""
        tmpdir = tempfile.mkdtemp()
        coord_file = Path(tmpdir) / "coord.xlsx"
        coord_file.touch()
        try:
            self.app.image_folder_path.set(tmpdir)
            self.app.coord_excel_path.set(str(coord_file))
            self.app.skip_questions.set("4")
            result = self.app.validate_inputs()
            assert result is True
            mock_mb.showerror.assert_not_called()
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


# ================================================================
# 3. run_box_drawer — ガード条件
# ================================================================

class TestRunBoxDrawerGuard:
    """run_box_drawer の入力チェックと処理中ガード"""

    def setup_method(self):
        self.root, self.app = _make_gui()

    def teardown_method(self):
        _destroy_gui(self.root)

    @patch("main_gui.messagebox")
    def test_blocked_during_processing(self, mock_mb):
        """処理中は即 return"""
        self.app._processing = True
        self.app.run_box_drawer()
        mock_mb.showerror.assert_not_called()

    @patch("main_gui.messagebox")
    def test_no_inputs_shows_error(self, mock_mb):
        """入力未設定 → エラーダイアログ"""
        self.app.run_box_drawer()
        mock_mb.showerror.assert_called()


# ================================================================
# 4. run_mark_checker — 入力チェック
# ================================================================

class TestRunMarkCheckerGuard:
    """run_mark_checker の入力チェック 4 パターン"""

    def setup_method(self):
        self.root, self.app = _make_gui()

    def teardown_method(self):
        _destroy_gui(self.root)

    @patch("main_gui.messagebox")
    def test_no_image_folder(self, mock_mb):
        """画像フォルダ未設定 → エラー"""
        self.app.run_mark_checker()
        mock_mb.showerror.assert_called_once()
        assert "画像フォルダ" in mock_mb.showerror.call_args[0][1]

    @patch("main_gui.messagebox")
    def test_no_coordinates_csv(self, mock_mb):
        """coordinates.csv が存在しない → エラー"""
        tmpdir = tempfile.mkdtemp()
        try:
            self.app.image_folder_path.set(tmpdir)
            self.app.run_mark_checker()
            mock_mb.showerror.assert_called_once()
            assert "coordinates.csv" in mock_mb.showerror.call_args[0][1]
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    @patch("main_gui.messagebox")
    def test_no_mark2_result(self, mock_mb):
        """Mark2結果ファイル未設定 → エラー"""
        tmpdir = tempfile.mkdtemp()
        try:
            # coordinates.csv を作成
            results_data = Path(tmpdir) / "_saiten_grading_results" / "01_Results"
            results_data.mkdir(parents=True)
            (results_data / "coordinates.csv").touch()
            self.app.image_folder_path.set(tmpdir)
            self.app.run_mark_checker()
            mock_mb.showerror.assert_called_once()
            assert "OMR読取結果" in mock_mb.showerror.call_args[0][1]
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    @patch("main_gui.messagebox")
    def test_mark2_result_not_exist(self, mock_mb):
        """Mark2結果ファイルが存在しない → エラー"""
        tmpdir = tempfile.mkdtemp()
        try:
            results_data = Path(tmpdir) / "_saiten_grading_results" / "01_Results"
            results_data.mkdir(parents=True)
            (results_data / "coordinates.csv").touch()
            self.app.image_folder_path.set(tmpdir)
            self.app.mark2_result_path.set("/nonexistent/result.xlsx")
            self.app.run_mark_checker()
            mock_mb.showerror.assert_called_once()
            assert "存在しません" in mock_mb.showerror.call_args[0][1]
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


# ================================================================
# 5. run_scoring — 入力チェック（全分岐）
# ================================================================

class TestRunScoringGuard:
    """run_scoring の入力バリデーション網羅"""

    def setup_method(self):
        self.root, self.app = _make_gui()

    def teardown_method(self):
        _destroy_gui(self.root)

    @patch("main_gui.messagebox")
    def test_no_image_folder(self, mock_mb):
        self.app.run_scoring()
        mock_mb.showerror.assert_called_once()
        assert "画像フォルダ" in mock_mb.showerror.call_args[0][1]

    @patch("main_gui.messagebox")
    def test_no_coord_excel(self, mock_mb):
        self.app.image_folder_path.set("/path")
        self.app.run_scoring()
        mock_mb.showerror.assert_called_once()
        assert "座標ファイル" in mock_mb.showerror.call_args[0][1]

    @patch("main_gui.messagebox")
    def test_no_template(self, mock_mb):
        self.app.image_folder_path.set("/path")
        self.app.coord_excel_path.set("/path.xlsx")
        self.app.run_scoring()
        mock_mb.showerror.assert_called_once()
        assert "正答データ" in mock_mb.showerror.call_args[0][1]

    @patch("main_gui.messagebox")
    def test_no_mark2_result(self, mock_mb):
        self.app.image_folder_path.set("/path")
        self.app.coord_excel_path.set("/path.xlsx")
        self.app.template_path.set("/template.xlsx")
        self.app.run_scoring()
        mock_mb.showerror.assert_called_once()
        assert "OMR読取結果" in mock_mb.showerror.call_args[0][1]

    @patch("main_gui.messagebox")
    def test_image_folder_not_exist(self, mock_mb):
        """全パス設定済みだがフォルダが存在しない"""
        self.app.image_folder_path.set("/nonexistent")
        self.app.coord_excel_path.set("/path.xlsx")
        self.app.template_path.set("/template.xlsx")
        self.app.mark2_result_path.set("/result.xlsx")
        self.app.run_scoring()
        mock_mb.showerror.assert_called_once()
        assert "存在しません" in mock_mb.showerror.call_args[0][1]


# ================================================================
# 6. run_summary_generation — 入力チェック
# ================================================================

class TestRunSummaryGuard:
    """run_summary_generation の入力バリデーション"""

    def setup_method(self):
        self.root, self.app = _make_gui()

    def teardown_method(self):
        _destroy_gui(self.root)

    @patch("main_gui.messagebox")
    def test_blocked_during_processing(self, mock_mb):
        self.app._processing = True
        self.app.run_summary_generation()
        mock_mb.showerror.assert_not_called()

    @patch("main_gui.messagebox")
    def test_no_image_folder(self, mock_mb):
        self.app.run_summary_generation()
        mock_mb.showerror.assert_called_once()
        assert "画像フォルダ" in mock_mb.showerror.call_args[0][1]

    @patch("main_gui.messagebox")
    def test_no_coord_excel(self, mock_mb):
        self.app.image_folder_path.set("/path")
        self.app.run_summary_generation()
        mock_mb.showerror.assert_called_once()
        assert "座標ファイル" in mock_mb.showerror.call_args[0][1]

    @patch("main_gui.messagebox")
    def test_no_template(self, mock_mb):
        self.app.image_folder_path.set("/path")
        self.app.coord_excel_path.set("/path.xlsx")
        self.app.run_summary_generation()
        mock_mb.showerror.assert_called_once()
        assert "正答データ" in mock_mb.showerror.call_args[0][1]

    @patch("main_gui.messagebox")
    def test_no_mark2_result(self, mock_mb):
        self.app.image_folder_path.set("/path")
        self.app.coord_excel_path.set("/path.xlsx")
        self.app.template_path.set("/template.xlsx")
        self.app.run_summary_generation()
        mock_mb.showerror.assert_called_once()
        assert "OMR読取結果" in mock_mb.showerror.call_args[0][1]


# ================================================================
# 7. open_threshold_calibrator — 入力チェック
# ================================================================

class TestOpenThresholdCalibratorGuard:
    """open_threshold_calibrator の前提条件チェック"""

    def setup_method(self):
        self.root, self.app = _make_gui()

    def teardown_method(self):
        _destroy_gui(self.root)

    @patch("main_gui.messagebox")
    def test_no_inputs(self, mock_mb):
        """画像フォルダ・座標ファイル未設定 → 警告"""
        self.app.open_threshold_calibrator()
        mock_mb.showwarning.assert_called_once()
        assert "入力不足" in mock_mb.showwarning.call_args[0][0]

    @patch("main_gui.messagebox")
    def test_image_folder_not_exist(self, mock_mb):
        """画像フォルダが存在しない → エラー"""
        self.app.image_folder_path.set("/nonexistent")
        self.app.coord_excel_path.set("/nonexistent.xlsx")
        self.app.open_threshold_calibrator()
        mock_mb.showerror.assert_called_once()
        assert "画像フォルダ" in mock_mb.showerror.call_args[0][1]

    @patch("main_gui.messagebox")
    def test_coord_excel_not_exist(self, mock_mb):
        """座標ファイルが存在しない → エラー"""
        tmpdir = tempfile.mkdtemp()
        try:
            self.app.image_folder_path.set(tmpdir)
            self.app.coord_excel_path.set("/nonexistent.xlsx")
            self.app.open_threshold_calibrator()
            mock_mb.showerror.assert_called_once()
            assert "座標ファイル" in mock_mb.showerror.call_args[0][1]
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


# ================================================================
# 8. setup_total_position — 入力チェック
# ================================================================

class TestSetupTotalPositionGuard:
    """setup_total_position の事前チェック"""

    def setup_method(self):
        self.root, self.app = _make_gui()

    def teardown_method(self):
        _destroy_gui(self.root)

    @patch("main_gui.messagebox")
    def test_no_image_folder(self, mock_mb):
        self.app.setup_total_position()
        mock_mb.showerror.assert_called_once()
        assert "画像フォルダ" in mock_mb.showerror.call_args[0][1]

    @patch("main_gui.messagebox")
    def test_no_boxed_folder(self, mock_mb):
        """boxed_folder (00_Processing) が存在しない"""
        tmpdir = tempfile.mkdtemp()
        try:
            self.app.image_folder_path.set(tmpdir)
            self.app.setup_total_position()
            mock_mb.showerror.assert_called_once()
            assert "補正済み画像" in mock_mb.showerror.call_args[0][1]
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    @patch("main_gui.messagebox")
    def test_no_images_in_boxed_folder(self, mock_mb):
        """boxed_folder は存在するがJPG/PNGなし"""
        tmpdir = tempfile.mkdtemp()
        try:
            boxed = Path(tmpdir) / "_saiten_grading_results" / "00_Processing"
            boxed.mkdir(parents=True)
            self.app.image_folder_path.set(tmpdir)
            self.app.setup_total_position()
            mock_mb.showerror.assert_called_once()
            assert "画像が見つかりません" in mock_mb.showerror.call_args[0][1]
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    @patch("main_gui.messagebox")
    def test_no_template(self, mock_mb):
        """正答データ未設定"""
        tmpdir = tempfile.mkdtemp()
        try:
            boxed = Path(tmpdir) / "_saiten_grading_results" / "00_Processing"
            boxed.mkdir(parents=True)
            # ダミー画像を配置
            (boxed / "test.jpg").touch()
            self.app.image_folder_path.set(tmpdir)
            self.app.template_path.set("")
            self.app.setup_total_position()
            mock_mb.showerror.assert_called_once()
            assert "正答データ" in mock_mb.showerror.call_args[0][1]
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


# ================================================================
# 9. 記述トグル — ボタンの表示/非表示
# ================================================================

class TestDescriptiveToggle:
    """記述ON/OFFの UI 切り替え"""

    def setup_method(self):
        self.root, self.app = _make_gui()

    def teardown_method(self):
        _destroy_gui(self.root)

    def test_toggle_on_shows_buttons(self):
        """記述ON → 3つのウィジェットが pack される"""
        self.app.descriptive_enabled.set(True)
        self.app._on_descriptive_toggle()
        self.root.update_idletasks()

        assert self.app.desc_setup_btn.winfo_manager() == "pack"
        assert self.app.desc_scoring_btn.winfo_manager() == "pack"
        assert self.app._desc_status_frame.winfo_manager() == "pack"

    def test_toggle_off_hides_buttons(self):
        """記述ON → OFF でウィジェットが非表示に戻る"""
        self.app.descriptive_enabled.set(True)
        self.app._on_descriptive_toggle()
        self.app.descriptive_enabled.set(False)
        self.app._on_descriptive_toggle()
        self.root.update_idletasks()

        assert self.app.desc_setup_btn.winfo_manager() == ""
        assert self.app.desc_scoring_btn.winfo_manager() == ""
        assert self.app._desc_status_frame.winfo_manager() == ""

    def test_toggle_roundtrip(self):
        """OFF → ON → OFF → ON: 2回目の ON でもボタンが正しく表示"""
        for _ in range(2):
            self.app.descriptive_enabled.set(True)
            self.app._on_descriptive_toggle()
            self.root.update_idletasks()
            assert self.app.desc_setup_btn.winfo_manager() == "pack"

            self.app.descriptive_enabled.set(False)
            self.app._on_descriptive_toggle()
            self.root.update_idletasks()
            assert self.app.desc_setup_btn.winfo_manager() == ""


# ================================================================
# 10. log_message — ログ表示
# ================================================================

class TestLogMessage:
    """log_message の基本動作"""

    def setup_method(self):
        self.root, self.app = _make_gui()

    def teardown_method(self):
        _destroy_gui(self.root)

    def test_log_appends_text(self):
        """メッセージがログに追加される"""
        self.app.log_message("テスト行1")
        self.app.log_message("テスト行2")
        content = self.app.log_text.get("1.0", tk.END)
        assert "テスト行1" in content
        assert "テスト行2" in content

    def test_log_replace_last(self):
        """replace_last=True で最終行が上書きされる"""
        self.app.log_message("行1")
        self.app.log_message("行2 (消される)")
        self.app.log_message("行2 (置換)", replace_last=True)
        content = self.app.log_text.get("1.0", tk.END)
        assert "行1" in content
        assert "行2 (置換)" in content
        # 上書きされた行は残っていないはず
        assert "行2 (消される)" not in content

    def test_log_text_stays_disabled(self):
        """ログ書き込み後も状態は disabled"""
        self.app.log_message("test")
        assert str(self.app.log_text["state"]) == "disabled"


# ================================================================
# 11. auto_detect_template — テンプレート自動検出
# ================================================================

class TestAutoDetectTemplate:
    """auto_detect_template の動作"""

    def setup_method(self):
        self.root, self.app = _make_gui()

    def teardown_method(self):
        _destroy_gui(self.root)

    def test_no_folder_does_nothing(self):
        """画像フォルダ未設定 → 何もしない"""
        self.app.auto_detect_template()
        content = self.app.log_text.get("1.0", tk.END).strip()
        assert content == ""  # ログ出力なし

    def test_template_found(self):
        """テンプレートが存在する場合、ログに通知"""
        tmpdir = tempfile.mkdtemp()
        try:
            template = Path(tmpdir) / "_saiten_grading_results" / "01_Results" / "answer_key.xlsx"
            template.parent.mkdir(parents=True)
            template.touch()
            self.app.image_folder_path.set(tmpdir)
            self.app.auto_detect_template()
            content = self.app.log_text.get("1.0", tk.END)
            assert "正答データを自動検出" in content
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    def test_template_not_found(self):
        """テンプレートが存在しない場合はログ出力なし"""
        tmpdir = tempfile.mkdtemp()
        try:
            self.app.image_folder_path.set(tmpdir)
            self.app.auto_detect_template()
            content = self.app.log_text.get("1.0", tk.END).strip()
            assert "テンプレート" not in content
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


# ================================================================
# 12. _set_processing_state — 状態遷移
# ================================================================

class TestProcessingState:
    """処理中/待機中の状態切り替え"""

    def setup_method(self):
        self.root, self.app = _make_gui()

    def teardown_method(self):
        _destroy_gui(self.root)

    def test_busy_disables_all(self):
        """busy=True → 全アクションボタン disabled"""
        self.app._set_processing_state(True)
        self.root.update_idletasks()

        assert self.app._processing is True
        for btn in [self.app._btn_run_box, self.app._btn_mark_check,
                     self.app._btn_total_pos, self.app._btn_run_scoring,
                     self.app._btn_run_summary]:
            assert str(btn["state"]) == "disabled"
        # チェックボックスも disabled
        assert str(self.app._chk_descriptive["state"]) == "disabled"

    def test_idle_enables_all(self):
        """busy=False → _processing解除 & Stepガードに従いボタン状態復元"""
        self.app._set_processing_state(True)
        self.app._set_processing_state(False)
        self.root.update_idletasks()

        assert self.app._processing is False
        # Stepガードにより、前提条件未達のボタンはdisabledのまま
        # チェックボックスとStep3ボタンは常にnormalに戻る
        assert str(self.app._chk_descriptive["state"]) == "normal"

    def test_progress_bar_visibility(self):
        """busy=True でプログレスバー表示、False で非表示"""
        self.app._set_processing_state(True)
        self.root.update_idletasks()
        assert self.app._progress_bar.winfo_manager() == "pack"

        self.app._set_processing_state(False)
        self.root.update_idletasks()
        assert self.app._progress_bar.winfo_manager() == ""


# ================================================================
# 13. select_folder — フォルダ選択チェーン
# ================================================================

class TestSelectFolderChain:
    """select_folder がコールバックチェーンを正しく起動するか"""

    def setup_method(self):
        self.root, self.app = _make_gui()

    def teardown_method(self):
        _destroy_gui(self.root)

    @patch("main_gui.filedialog")
    def test_cancel_does_nothing(self, mock_fd):
        """キャンセル時はフォルダ変更なし"""
        mock_fd.askdirectory.return_value = ""
        self.app.select_folder()
        assert self.app.image_folder_path.get() == ""

    @patch.object(sys.modules.get('main_gui', MagicMock()), 'Mark2GUI.auto_detect_template',
                  create=True, new_callable=MagicMock)
    @patch("main_gui.filedialog")
    def test_folder_selected_triggers_chain(self, mock_fd, _):
        """フォルダ選択後、auto_detect_template が呼ばれる"""
        tmpdir = tempfile.mkdtemp()
        try:
            mock_fd.askdirectory.return_value = tmpdir
            with patch.object(self.app, '_try_auto_restore') as mock_restore, \
                 patch.object(self.app, 'auto_detect_template') as mock_detect:
                self.app.select_folder()
                assert self.app.image_folder_path.get() == tmpdir
                mock_restore.assert_called_once()
                mock_detect.assert_called_once()
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


# ================================================================
# 14. _auto_detect_omr_result — OMR結果自動検出
# ================================================================

class TestAutoDetectOMRResult:
    """_auto_detect_omr_result の動作"""

    def setup_method(self):
        self.root, self.app = _make_gui()

    def teardown_method(self):
        _destroy_gui(self.root)

    def test_no_reading_folder(self):
        """reading_results フォルダがない → 何もしない"""
        tmpdir = tempfile.mkdtemp()
        try:
            self.app._auto_detect_omr_result(tmpdir)
            assert self.app.mark2_result_path.get() == ""
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    def test_detects_latest_result(self):
        """Mark2-Result-*.xlsx の最新ファイルを検出"""
        tmpdir = tempfile.mkdtemp()
        try:
            reading = Path(tmpdir) / "01_Results" / "reading_results"
            reading.mkdir(parents=True)
            (reading / "Mark2-Result-20260101.xlsx").touch()
            (reading / "Mark2-Result-20260210.xlsx").touch()
            self.app._auto_detect_omr_result(tmpdir)
            assert "20260210" in self.app.mark2_result_path.get()
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    def test_no_result_files(self):
        """reading_results はあるが結果ファイルがない"""
        tmpdir = tempfile.mkdtemp()
        try:
            reading = Path(tmpdir) / "01_Results" / "reading_results"
            reading.mkdir(parents=True)
            self.app._auto_detect_omr_result(tmpdir)
            assert self.app.mark2_result_path.get() == ""
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


# ================================================================
# 15. run_scoring 記述ON時の事前チェック
# ================================================================

class TestRunScoringDescriptiveGuard:
    """run_scoring で記述ONかつデータ不足のケース"""

    def setup_method(self):
        self.root, self.app = _make_gui()

    def teardown_method(self):
        _destroy_gui(self.root)

    @patch("main_gui.messagebox")
    def test_descriptive_on_no_config(self, mock_mb):
        """記述ON + descriptive_config.json なし → エラー"""
        tmpdir = tempfile.mkdtemp()
        template = Path(tmpdir) / "template.xlsx"
        template.touch()
        result_file = Path(tmpdir) / "result.xlsx"
        result_file.touch()
        coord_file = Path(tmpdir) / "coord.xlsx"
        coord_file.touch()
        try:
            self.app.image_folder_path.set(tmpdir)
            self.app.coord_excel_path.set(str(coord_file))
            self.app.template_path.set(str(template))
            self.app.mark2_result_path.set(str(result_file))
            self.app.descriptive_enabled.set(True)
            self.app.run_scoring()
            # 記述設定ファイルなしのエラーが出るはず
            assert mock_mb.showerror.called
            error_msg = mock_mb.showerror.call_args[0][1]
            assert "記述" in error_msg
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


# ================================================================
# Phase 1A: ThresholdCalibratorGUI コード構造テスト
# ================================================================

class TestThresholdCalibratorGUIStructure:
    """ThresholdCalibratorGUI の v4.0 改善が正しく反映されているかをコード構造で検証"""

    def test_japanese_labels_in_source(self):
        """スライダーラベルが日本語化されている"""
        import inspect
        from gui_components import ThresholdCalibratorGUI
        src = inspect.getsource(ThresholdCalibratorGUI)
        assert "色の読取感度" in src, "Color Threshold が日本語化されていない"
        assert "面積の読取感度" in src, "Area Threshold が日本語化されていない"
        # 英語ラベルが残っていないことを確認
        assert "Color Threshold" not in src, "英語の Color Threshold ラベルが残っている"
        assert "Area Threshold" not in src, "英語の Area Threshold ラベルが残っている"

    def test_detail_section_collapsible(self):
        """テスト読み取りセクションが折りたたみ機能を持つ"""
        import inspect
        from gui_components import ThresholdCalibratorGUI
        src = inspect.getsource(ThresholdCalibratorGUI)
        assert "_toggle_detail_section" in src, "折りたたみトグルメソッドがない"
        assert "_detail_frame" in src, "詳細設定フレームが定義されていない"
        assert "_detail_expanded" in src, "展開フラグがない"

    def test_gallery_summary_header(self):
        """ギャラリーにサマリヘッダーが含まれる"""
        import inspect
        from gui_components import ThresholdCalibratorGUI
        src = inspect.getsource(ThresholdCalibratorGUI)
        assert "境界付近" in src, "ギャラリーヘッダーに境界付近の情報がない"
        assert "summary_frame" in src or "summary_text" in src, "サマリフレームがない"

    def test_window_size_reduced(self):
        """ウィンドウサイズが縮小されている (1200x780 → 1100x700)"""
        import inspect
        from gui_components import ThresholdCalibratorGUI
        src = inspect.getsource(ThresholdCalibratorGUI)
        assert "1100x700" in src, "ウィンドウサイズが1100x700に変更されていない"
        assert "1200x780" not in src, "旧サイズ1200x780が残っている"

    def test_stats_japanese_labels(self):
        """統計情報テキストのラベルが日本語化されている"""
        import inspect
        from gui_components import ThresholdCalibratorGUI
        src = inspect.getsource(ThresholdCalibratorGUI)
        assert "色=" in src and "面積=" in src, "統計情報の閾値ラベルが日本語化されていない"


# ================================================================
# Phase 1B: MarkCheckerGUI コード構造テスト
# ================================================================

class TestMarkCheckerGUIStructure:
    """MarkCheckerGUI の v4.0 改善が正しく反映されているかをコード構造で検証"""

    def test_choice_buttons_exist(self):
        """選択肢ボタン行が存在する"""
        import inspect
        from gui_components import MarkCheckerGUI
        src = inspect.getsource(MarkCheckerGUI)
        assert "_choice_btn_frame" in src, "選択肢ボタンフレームがない"
        assert "_build_choice_buttons" in src, "ボタン構築メソッドがない"
        assert "_on_choice_button" in src, "ボタンハンドラがない"

    def test_keyboard_shortcuts(self):
        """数字キーショートカットが実装されている"""
        import inspect
        from gui_components import MarkCheckerGUI
        src = inspect.getsource(MarkCheckerGUI)
        assert "_on_key_press" in src, "キープレスハンドラがない"
        assert "1234567890" in src, "数字キーマッピングがない"

    def test_dynamic_choice_count(self):
        """座標CSVから選択肢数を動的取得するメソッドがある"""
        import inspect
        from gui_components import MarkCheckerGUI
        src = inspect.getsource(MarkCheckerGUI)
        assert "_get_num_choices_for_question" in src, "問題別選択肢数取得メソッドがない"
        assert "_update_max_choices_from_coords" in src, "最大選択肢数更新メソッドがない"
        assert "mark_coords" in src, "mark_coordsカラム参照がない"

    def test_hint_label_dynamic(self):
        """ヒントテキストが動的に更新される"""
        import inspect
        from gui_components import MarkCheckerGUI
        src = inspect.getsource(MarkCheckerGUI)
        assert "_hint_label" in src, "ヒントラベルがない"
        assert "マークなし" in src, "マークなし表示がない"
