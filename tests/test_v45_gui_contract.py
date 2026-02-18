#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""v4.5 GUI 契約テスト（実コード経由）

方針:
- 手組みモックUIではなく、実際の GUI クラスを生成して検証する
- 視覚キャプチャではなく、状態・ウィジェット契約を高速に検証する
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

import pandas as pd
import pytest
import tkinter as tk

sys.path.insert(0, str(Path(__file__).parent.parent / "main_src"))

from conftest import get_shared_tk_root
from constants import (
    APP_VERSION,
    MODE_MARK_AND_DESCRIPTIVE,
    MODE_DESCRIPTIVE_ONLY,
    OMR_MODE_KMEANS,
    OMR_MODE_THRESHOLD,
    ERROR_TYPE_NO_MARK,
    ERROR_TYPE_DOUBLE_MARK,
)
from main_gui import SaitenSamuraiGUI
from gui_components import MarkCheckerGUI


# ============================================================
# SaitenSamuraiGUI (v4.5 OMRモード)
# ============================================================


def _create_main_gui(mode=MODE_MARK_AND_DESCRIPTIVE):
    root = get_shared_tk_root()
    top = tk.Toplevel(root)
    top.withdraw()
    app = SaitenSamuraiGUI(top, mode=mode)
    top.update_idletasks()
    return top, app


def _destroy_top(top):
    try:
        top.destroy()
    except Exception:
        pass


class TestV45OmrModeContract:
    def test_default_mode_is_kmeans(self):
        top, app = _create_main_gui()
        try:
            assert app.omr_mode.get() == OMR_MODE_KMEANS
            # コンボボックスの表示ラベルがユーザーフレンドリーな日本語であること
            combo_values = tuple(app._omr_mode_combo["values"])
            assert len(combo_values) == 2
            assert "クラスタリング" in combo_values[0]
            assert "しきい値" in combo_values[1]
        finally:
            _destroy_top(top)

    def test_threshold_slider_toggle(self):
        top, app = _create_main_gui()
        try:
            top.update_idletasks()
            assert app._omr_slider_row.winfo_manager() == ""

            # 表示ラベル経由で切り替え（実際のユーザー操作を再現）
            app._omr_display_var.set(app._omr_value_to_label[OMR_MODE_THRESHOLD])
            app._on_omr_mode_changed()
            top.update_idletasks()
            assert app.omr_mode.get() == OMR_MODE_THRESHOLD
            assert app._omr_slider_row.winfo_manager() == "pack"

            app._omr_display_var.set(app._omr_value_to_label[OMR_MODE_KMEANS])
            app._on_omr_mode_changed()
            top.update_idletasks()
            assert app.omr_mode.get() == OMR_MODE_KMEANS
            assert app._omr_slider_row.winfo_manager() == ""
        finally:
            _destroy_top(top)

    def test_descriptive_only_hides_omr_mode_row(self):
        top, app = _create_main_gui(mode=MODE_DESCRIPTIVE_ONLY)
        try:
            top.update_idletasks()
            assert app._omr_mode_row.winfo_manager() == ""
            assert app._omr_slider_row.winfo_manager() == ""
        finally:
            _destroy_top(top)

    def test_title_uses_app_version_constant(self):
        top, _app = _create_main_gui()
        try:
            assert f"v{APP_VERSION}" in top.title()
        finally:
            _destroy_top(top)


# ============================================================
# MarkCheckerGUI (v4.5 グリッド表示)
# ============================================================


@pytest.fixture
def mark_checker_gui(monkeypatch):
    monkeypatch.setattr(MarkCheckerGUI, "load_data", lambda self: None)

    root = get_shared_tk_root()
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        coords_csv = tmp / "coords.csv"
        xlsx_path = tmp / "results.xlsx"
        coords_csv.write_text("", encoding="utf-8")
        xlsx_path.write_text("", encoding="utf-8")

        gui = MarkCheckerGUI(
            parent_window=root,
            image_folder=str(tmp),
            coords_csv_path=str(coords_csv),
            xlsx_path=str(xlsx_path),
            skip_questions=0,
            template_path=None,
        )
        gui.window.update_idletasks()
        yield gui

        try:
            gui.on_close()
        except Exception:
            try:
                gui.window.destroy()
            except Exception:
                pass


class TestV45MarkCheckerGridContract:
    def test_grid_controls_exist_and_initial_mode(self, mark_checker_gui):
        gui = mark_checker_gui
        assert gui._view_mode == "single"
        assert gui._btn_toggle_view["text"] == "グリッド表示"
        assert gui._single_view_frame.winfo_manager() == "pack"
        assert gui._grid_view_frame.winfo_manager() == ""

    def test_toggle_to_grid_and_back(self, mark_checker_gui):
        gui = mark_checker_gui

        gui._toggle_view_mode()
        gui.window.update_idletasks()
        assert gui._view_mode == "grid"
        assert gui._btn_toggle_view["text"] == "単体表示"
        assert gui._grid_view_frame.winfo_manager() == "pack"

        gui._toggle_view_mode()
        gui.window.update_idletasks()
        assert gui._view_mode == "single"
        assert gui._btn_toggle_view["text"] == "グリッド表示"
        assert gui._single_view_frame.winfo_manager() == "pack"

    def test_grid_card_size_slider(self, mark_checker_gui):
        gui = mark_checker_gui
        # 初期サイズ 160
        assert gui._grid_thumb_size == 160
        assert gui._grid_size_var.get() == 160
        # スライダーが存在すること
        assert hasattr(gui, '_grid_size_slider')
        # サイズ変更が反映されること
        gui._grid_size_var.set(200)
        gui._on_grid_size_changed()
        gui.window.update_idletasks()
        assert gui._grid_thumb_size == 200

    def test_filter_logic_contract(self, mark_checker_gui):
        gui = mark_checker_gui
        gui.error_df = pd.DataFrame(
            [
                {"error_type": ERROR_TYPE_NO_MARK, "after": ""},
                {"error_type": ERROR_TYPE_DOUBLE_MARK, "after": "2"},
                {"error_type": ERROR_TYPE_NO_MARK, "after": "-1"},
                {"error_type": ERROR_TYPE_DOUBLE_MARK, "after": ""},
            ]
        )

        gui._grid_filter_var.set("全て")
        assert len(gui._get_filtered_error_indices()) == 4

        gui._grid_filter_var.set("無マーク")
        assert len(gui._get_filtered_error_indices()) == 2

        gui._grid_filter_var.set("ダブルマーク")
        assert len(gui._get_filtered_error_indices()) == 2

        gui._grid_filter_var.set("チェック済み")
        assert len(gui._get_filtered_error_indices()) == 2

        gui._grid_filter_var.set("未チェック")
        assert len(gui._get_filtered_error_indices()) == 2
