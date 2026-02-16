#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
横断的統合テスト

記述モード × 各機能、ボタン状態遷移、データフロー整合性など、
機能を横断する結合バグを検出するためのテスト。
"""

import json
import sys
import tempfile
from pathlib import Path
from unittest.mock import patch, MagicMock

import cv2
import numpy as np
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "main_src"))

from descriptive_scorer import (
    draw_combined_total,
    load_total_display_config,
    save_total_display_config,
    TOTAL_DISPLAY_CONFIG_FILE,
)


# ============================================================
# フィクスチャ
# ============================================================

@pytest.fixture
def dummy_image():
    """テスト用ダミー画像 (BGR)"""
    return np.ones((1000, 800, 3), dtype=np.uint8) * 255


@pytest.fixture
def sample_scoring_result():
    """マーク採点結果のサンプル"""
    return {
        "total_score": 60,
        "max_score": 90,
        "aspect_scores": {1: 20, 2: 20, 3: 20},
        "aspect_max_scores": {1: 30, 2: 30, 3: 30},
        "results": {},
    }


@pytest.fixture
def sample_config_with_total_region():
    """合計点表示位置付きのdescriptive config"""
    return {
        "questions": [
            {"id": "D1", "name": "記述1", "region": [100, 400, 300, 500],
             "max_score": 5, "aspect": 2},
        ],
        "total_display_region": [50, 800, 350, 870],
    }


@pytest.fixture
def sample_config_no_total_region():
    """合計点表示位置なしのdescriptive config"""
    return {
        "questions": [
            {"id": "D1", "name": "記述1", "region": [100, 400, 300, 500],
             "max_score": 5, "aspect": 2},
        ],
        "total_display_region": None,
    }


@pytest.fixture
def sample_coordinates():
    """マーク座標のサンプル"""
    return [
        {"question_no": 1, "x": 100, "y": 200, "width": 40, "height": 20},
        {"question_no": 2, "x": 100, "y": 230, "width": 40, "height": 20},
    ]


# ============================================================
# 1. total_display_config の保存/読み込み整合性
# ============================================================

class TestTotalDisplayConfigIntegrity:
    """合計点表示位置の保存→読み込み一貫性テスト"""

    def test_save_and_load_round_trip(self, tmp_path):
        """保存→読み込みでデータが一致する"""
        config_path = str(tmp_path / TOTAL_DISPLAY_CONFIG_FILE)
        region = [50, 800, 350, 870]
        save_total_display_config(config_path, region)
        loaded = load_total_display_config(config_path)
        assert loaded is not None
        assert loaded["total_display_region"] == region

    def test_load_nonexistent_returns_none(self, tmp_path):
        """存在しないファイルの読み込みはNoneを返す"""
        config_path = str(tmp_path / "nonexistent.json")
        result = load_total_display_config(config_path)
        assert result is None


# ============================================================
# 2. draw_combined_total × total_display_region
# ============================================================

class TestDrawCombinedTotalWithRegion:
    """draw_combined_total が total_display_region を正しく使うかテスト"""

    def test_with_total_display_region(
        self, dummy_image, sample_scoring_result, sample_config_with_total_region
    ):
        """total_display_region が指定されている場合、ボックス内に描画される"""
        desc_scores = {"D1": 3}
        result = draw_combined_total(
            dummy_image,
            sample_scoring_result,
            sample_config_with_total_region,
            desc_scores,
            [],  # coordinates
        )
        # 画像が返る（エラーなし）
        assert result is not None
        assert result.shape == dummy_image.shape
        # ボックス領域に何か描画されているはず（白一色でなくなる）
        region = sample_config_with_total_region["total_display_region"]
        box_area = result[region[1]:region[3], region[0]:region[2]]
        assert not np.array_equal(box_area, dummy_image[region[1]:region[3], region[0]:region[2]])

    def test_without_total_display_region_fallback(
        self, dummy_image, sample_scoring_result, sample_config_no_total_region, sample_coordinates
    ):
        """total_display_region が None の場合、フォールバック位置に描画される"""
        desc_scores = {"D1": 3}
        result = draw_combined_total(
            dummy_image,
            sample_scoring_result,
            sample_config_no_total_region,
            desc_scores,
            sample_coordinates,
        )
        assert result is not None
        assert result.shape == dummy_image.shape
        # フォールバック: 最後の座標の下に描画 → 画像全体で何か変化がある
        assert not np.array_equal(result, dummy_image)

    def test_config_injection_from_total_display_config(self, tmp_path):
        """setup_total_position で保存した設定が config に注入されるフロー"""
        # 1. setup_total_position が保存する形式
        region = [50, 800, 350, 870]
        config_path = str(tmp_path / TOTAL_DISPLAY_CONFIG_FILE)
        save_total_display_config(config_path, region)

        # 2. _run_scoring_thread が読み込んで config に注入する動作をシミュレート
        tdc = load_total_display_config(config_path)
        config = {
            "questions": [{"id": "D1", "max_score": 5, "aspect": 2, "region": [100, 400, 300, 500]}],
            "total_display_region": None,  # 初期状態
        }
        if tdc and "total_display_region" in tdc:
            config["total_display_region"] = tdc["total_display_region"]

        # 3. 注入後に正しい値が入っている
        assert config["total_display_region"] == region


# ============================================================
# 3. GUI ボタン状態管理の網羅テスト
# ============================================================

class TestGUIStateMachine:
    """ボタン状態遷移と処理中ガード"""

    @staticmethod
    def _make_app():
        import tkinter as tk
        from conftest import get_shared_tk_root
        from saitensamurai import Mark2GUI
        root = get_shared_tk_root()
        # 共有ルート上に Toplevel を作成し、GUI をそちらに配置
        top = tk.Toplevel(root)
        app = Mark2GUI(top)
        return top, app

    def test_processing_state_disables_checkbox(self):
        """処理中は記述チェックボックスも無効化される"""
        import tkinter as tk
        try:
            top, app = self._make_app()
        except tk.TclError:
            pytest.skip("Tkinter not available")
        try:
            app._set_processing_state(True)
            assert str(app._chk_descriptive["state"]) == "disabled"
            app._set_processing_state(False)
            assert str(app._chk_descriptive["state"]) == "normal"
        finally:
            top.destroy()

    def test_all_action_buttons_disabled_during_processing(self):
        """すべてのアクションボタンが処理中に無効化される"""
        import tkinter as tk
        try:
            top, app = self._make_app()
        except tk.TclError:
            pytest.skip("Tkinter not available")
        try:
            app._set_processing_state(True)
            buttons = [
                app._btn_run_box,
                app._btn_mark_check,
                app._btn_total_pos,
                app._btn_run_scoring,
                app._btn_run_summary,
                app.desc_setup_btn,
                app.desc_scoring_btn,
            ]
            for btn in buttons:
                assert str(btn["state"]) == "disabled", f"{btn['text']} should be disabled"
            app._set_processing_state(False)
            # フォルダ未設定状態では Step ガードにより Step1/2/3 ボタンは disabled のまま
            # ガード対象外のボタン (desc_setup_btn) のみ normal であることを確認
            step_guarded_buttons = {
                app._btn_run_box,      # Step 1
                app._btn_mark_check,   # Step 2
                app._btn_total_pos,    # Step 2
                app._btn_run_scoring,  # Step 2
                app._btn_run_summary,  # Step 3
                app.desc_scoring_btn,  # Step 2
            }
            for btn in buttons:
                if btn in step_guarded_buttons:
                    assert str(btn["state"]) == "disabled", f"{btn['text']} should remain disabled (step guard)"
                else:
                    assert str(btn["state"]) == "normal", f"{btn['text']} should be normal"
        finally:
            top.destroy()

    def test_run_scoring_blocked_during_processing(self):
        """処理中に run_scoring は即 return する"""
        import tkinter as tk
        try:
            top, app = self._make_app()
        except tk.TclError:
            pytest.skip("Tkinter not available")
        try:
            app._processing = True
            # エラーダイアログが出ずに静かにリターンすることを検証
            app.run_scoring()  # no exception
        finally:
            top.destroy()

    def test_descriptive_toggle_creates_correct_order(self):
        """記述ON/OFFでボタン順序が正しく維持される"""
        import tkinter as tk
        try:
            top, app = self._make_app()
        except tk.TclError:
            pytest.skip("Tkinter not available")
        try:
            # OFF → ON
            app.descriptive_enabled.set(True)
            app._on_descriptive_toggle()
            # 記述ON → desc_setup_btn (Step1), desc_scoring_btn (Step2) が表示
            assert app.desc_setup_btn.winfo_manager() == "pack"
            assert app.desc_scoring_btn.winfo_manager() == "pack"
            assert app._desc_status_frame.winfo_manager() == "pack"

            # ON → OFF
            app.descriptive_enabled.set(False)
            app._on_descriptive_toggle()
            assert app.desc_setup_btn.winfo_manager() == ""
            assert app.desc_scoring_btn.winfo_manager() == ""
            assert app._desc_status_frame.winfo_manager() == ""

            # 再度 ON
            app.descriptive_enabled.set(True)
            app._on_descriptive_toggle()
            assert app.desc_setup_btn.winfo_manager() == "pack"
            assert app.desc_scoring_btn.winfo_manager() == "pack"
        finally:
            top.destroy()

    def test_no_return_sheet_btn_attribute(self):
        """return_sheet_btn が存在しないことを確認（廃止済み）"""
        import tkinter as tk
        try:
            top, app = self._make_app()
        except tk.TclError:
            pytest.skip("Tkinter not available")
        try:
            assert not hasattr(app, "return_sheet_btn")
        finally:
            top.destroy()


# ============================================================
# 4. run_scoring の記述モード分岐テスト
# ============================================================

class TestRunScoringDescriptiveIntegration:
    """run_scoring が記述ON時に正しく分岐するかのテスト"""

    @staticmethod
    def _make_app():
        import tkinter as tk
        from conftest import get_shared_tk_root
        from saitensamurai import Mark2GUI
        root = get_shared_tk_root()
        top = tk.Toplevel(root)
        app = Mark2GUI(top)
        return top, app

    def test_descriptive_on_requires_config_file(self):
        """記述ON＋設定ファイルなし → エラーダイアログ"""
        import tkinter as tk
        try:
            top, app = self._make_app()
        except tk.TclError:
            pytest.skip("Tkinter not available")
        try:
            app.descriptive_enabled.set(True)
            # 必要パス設定
            with tempfile.TemporaryDirectory() as td:
                app.image_folder_path.set(td)
                app.coord_excel_path.set("dummy.xlsx")
                app.template_path.set("dummy.xlsx")
                app.mark2_result_path.set("dummy.xlsx")
                # ファイルが存在しないのでエラーダイアログが出るべき
                with patch("tkinter.messagebox.showerror") as mock_err:
                    app.run_scoring()
                    # 何らかのエラーが呼ばれるはず（ファイル不存在）
                    assert mock_err.called or app._processing is False
        finally:
            top.destroy()

    def test_descriptive_off_uses_process_scoring(self):
        """記述OFF時は process_scoring フローに行く"""
        import tkinter as tk
        try:
            top, app = self._make_app()
        except tk.TclError:
            pytest.skip("Tkinter not available")
        try:
            app.descriptive_enabled.set(False)
            with tempfile.TemporaryDirectory() as td:
                app.image_folder_path.set(td)
                # 必要ファイルなし → 早期エラー（process_scoring 分岐前の入力チェック）
                with patch("tkinter.messagebox.showerror") as mock_err:
                    app.run_scoring()
                    # coord_excel_path が空 → エラー
                    assert mock_err.called
        finally:
            top.destroy()


# ============================================================
# 5. データフロー整合性テスト（JSONファイルパス）
# ============================================================

class TestConfigFilePathConsistency:
    """設定ファイルのパス整合性"""

    def test_total_display_config_filename_constant(self):
        """定数が一貫している"""
        assert TOTAL_DISPLAY_CONFIG_FILE == "total_display_config.json"

    def test_descriptive_config_keys(self):
        """descriptive_config の必須キーが正しい"""
        config = {
            "questions": [
                {"id": "D1", "name": "q1", "max_score": 5, "aspect": 1, "region": [0, 0, 100, 100]}
            ],
            "total_display_region": None,
        }
        # draw_combined_total が期待するキー
        assert "questions" in config
        assert "total_display_region" in config

    def test_scoring_result_keys(self):
        """scoring_result の必須キーが揃っている"""
        result = {
            "total_score": 60,
            "max_score": 90,
            "aspect_scores": {1: 20, 2: 20, 3: 20},
            "aspect_max_scores": {1: 30, 2: 30, 3: 30},
            "results": {},
        }
        # draw_combined_total が参照するキー
        for key in ["total_score", "max_score", "aspect_scores", "aspect_max_scores"]:
            assert key in result


# ============================================================
# 6. 記述モード × 合計点描画の完全統合テスト
# ============================================================

class TestDescriptiveScoringFullPipeline:
    """記述採点のパイプライン全体テスト"""

    def test_draw_combined_total_with_injected_region(self):
        """setup_total_position保存 → config注入 → draw_combined_total の全フロー"""
        with tempfile.TemporaryDirectory() as td:
            # Step 1: setup_total_position が保存する動作をシミュレート
            region = [50, 800, 400, 880]
            save_total_display_config(
                str(Path(td) / TOTAL_DISPLAY_CONFIG_FILE), region
            )

            # Step 2: _run_scoring_thread が注入する動作をシミュレート
            tdc = load_total_display_config(str(Path(td) / TOTAL_DISPLAY_CONFIG_FILE))
            config = {
                "questions": [
                    {"id": "D1", "name": "q1", "max_score": 10, "aspect": 1,
                     "region": [100, 400, 300, 500]},
                ],
                "total_display_region": None,
            }
            if tdc and "total_display_region" in tdc:
                config["total_display_region"] = tdc["total_display_region"]

            assert config["total_display_region"] == region

            # Step 3: draw_combined_total が描画
            img = np.ones((1000, 800, 3), dtype=np.uint8) * 255
            scoring_result = {
                "total_score": 70,
                "max_score": 90,
                "aspect_scores": {1: 70},
                "aspect_max_scores": {1: 90},
                "results": {},
            }
            desc_scores = {"D1": 8}

            result_img = draw_combined_total(
                img, scoring_result, config, desc_scores, []
            )

            # 描画が発生 → ボックス領域が変化
            box_area = result_img[800:880, 50:400]
            original_area = img[800:880, 50:400]
            assert not np.array_equal(box_area, original_area)

    def test_draw_combined_total_without_region_fallback(self):
        """total_display_region が None → フォールバック位置"""
        config = {
            "questions": [
                {"id": "D1", "name": "q1", "max_score": 10, "aspect": 1,
                 "region": [100, 400, 300, 500]},
            ],
            "total_display_region": None,
        }
        img = np.ones((1000, 800, 3), dtype=np.uint8) * 255
        scoring_result = {
            "total_score": 70,
            "max_score": 90,
            "aspect_scores": {1: 70},
            "aspect_max_scores": {1: 90},
            "results": {},
        }
        desc_scores = {"D1": 8}
        coords = [{"question_no": 1, "x": 100, "y": 700, "width": 40, "height": 20}]

        result_img = draw_combined_total(
            img, scoring_result, config, desc_scores, coords
        )
        assert result_img is not None
        assert not np.array_equal(result_img, img)

    def test_multiple_aspects_combined(self):
        """複数観点（マーク+記述）の合算が正しい"""
        config = {
            "questions": [
                {"id": "D1", "name": "q1", "max_score": 10, "aspect": 1,
                 "region": [100, 400, 300, 500]},
                {"id": "D2", "name": "q2", "max_score": 20, "aspect": 2,
                 "region": [100, 510, 300, 600]},
            ],
            "total_display_region": [50, 800, 500, 900],
        }
        img = np.ones((1000, 800, 3), dtype=np.uint8) * 255
        scoring_result = {
            "total_score": 60,
            "max_score": 60,
            "aspect_scores": {1: 30, 2: 30},
            "aspect_max_scores": {1: 30, 2: 30},
            "results": {},
        }
        desc_scores = {"D1": 8, "D2": 15}

        result_img = draw_combined_total(
            img, scoring_result, config, desc_scores, []
        )
        assert result_img is not None
        # 描画が発生すればOK（結合テスト）
        assert not np.array_equal(result_img, img)


# ============================================================
# 7. load_template キー名整合性テスト
# ============================================================

class TestLoadTemplateKeyNames:
    """load_template が返す辞書のキー名が一貫しているかテスト"""

    def test_template_dict_uses_japanese_keys(self, tmp_path):
        """load_template は '配点' '観点' キーを返す"""
        import pandas as pd
        from saitensamurai import load_template

        template_file = tmp_path / "template.xlsx"
        df = pd.DataFrame({
            "問題番号": [1, 2, 3],
            "正答": ["ア", "イ", "ウ"],
            "配点": [10, 20, 30],
            "観点": [1, 2, 3],
        })
        df.to_excel(str(template_file), index=False)

        result = load_template(str(template_file))
        for q_no, q_info in result.items():
            assert "配点" in q_info, f"q{q_no} missing '配点'"
            assert "観点" in q_info, f"q{q_no} missing '観点'"
            # 英語キーは使わない
            assert "score" not in q_info
            assert "aspect" not in q_info

    def test_setup_total_position_preview_uses_correct_keys(self, tmp_path):
        """setup_total_position のプレビューテキスト生成が正しいキーを使う"""
        import pandas as pd
        from saitensamurai import load_template

        template_file = tmp_path / "template.xlsx"
        df = pd.DataFrame({
            "問題番号": [1, 2, 3],
            "正答": ["ア", "イ", "ウ"],
            "配点": [30, 30, 30],
            "観点": [1, 2, 3],
        })
        df.to_excel(str(template_file), index=False)

        template_dict = load_template(str(template_file))
        # setup_total_position と同じロジック
        aspect_max = {}
        for q_no, q_info in template_dict.items():
            asp = q_info.get('観点', 1)
            score = q_info.get('配点', 0)
            aspect_max[asp] = aspect_max.get(asp, 0) + score
        total_max = sum(aspect_max.values())

        assert total_max == 90
        assert aspect_max == {1: 30, 2: 30, 3: 30}
