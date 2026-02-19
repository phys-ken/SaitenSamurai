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
    def test_grid_default_view_and_side_panel(self, mark_checker_gui):
        """デフォルトがグリッド表示で、サイドパネルが存在すること"""
        gui = mark_checker_gui
        assert gui._view_mode == "grid"
        assert gui._grid_view_frame.winfo_manager() == "pack"
        assert gui._single_view_frame.winfo_manager() == ""
        # サイドパネル
        assert hasattr(gui, '_category_listbox')
        assert hasattr(gui, '_sort_combo')
        assert hasattr(gui, '_side_panel')

    def test_switch_to_single_and_back(self, mark_checker_gui):
        """単体表示への切り替えとグリッドへの復帰"""
        gui = mark_checker_gui
        # all_entries_df を設定して単体表示に切り替え
        gui._all_entries_df = pd.DataFrame([
            {"filename": "test.jpg", "question_no": 1, "before": "1",
             "after": "", "error_type": "", "category": "1"},
        ])
        gui.coords_df = pd.DataFrame(columns=['image_path', 'question_no', 'choices_bbox', 'mark_coords'])

        gui._switch_to_single(0)
        gui.window.update_idletasks()
        assert gui._view_mode == "single"
        assert gui._single_view_frame.winfo_manager() == "pack"

        gui._switch_to_grid()
        gui.window.update_idletasks()
        assert gui._view_mode == "grid"
        assert gui._grid_view_frame.winfo_manager() == "pack"

    def test_grid_card_size_slider(self, mark_checker_gui):
        gui = mark_checker_gui
        assert gui._grid_thumb_size == 160
        assert gui._grid_size_var.get() == 160
        assert hasattr(gui, '_grid_size_slider')
        gui._grid_size_var.set(200)
        gui._on_grid_size_changed()
        gui.window.update_idletasks()
        assert gui._grid_thumb_size == 200

    def test_category_filter_logic(self, mark_checker_gui):
        """カテゴリフィルタが正しく動作すること"""
        gui = mark_checker_gui
        gui._all_entries_df = pd.DataFrame([
            {"filename": "a.jpg", "question_no": 1, "before": "1",
             "after": "", "error_type": "", "category": "1"},
            {"filename": "a.jpg", "question_no": 2, "before": "",
             "after": "", "error_type": ERROR_TYPE_NO_MARK, "category": "ノーマーク"},
            {"filename": "b.jpg", "question_no": 1, "before": "2",
             "after": "", "error_type": "", "category": "2"},
            {"filename": "b.jpg", "question_no": 2, "before": "1;3",
             "after": "-1", "error_type": ERROR_TYPE_DOUBLE_MARK, "category": "複数マーク"},
        ])

        # カテゴリリスト構築
        gui._rebuild_category_list()

        # 全て
        gui._category_listbox.selection_clear(0, tk.END)
        gui._category_listbox.selection_set(0)  # "全て"
        assert len(gui._get_filtered_indices()) == 4

        # 要チェック: error_type ありかつ after != -1
        gui._category_listbox.selection_clear(0, tk.END)
        gui._category_listbox.selection_set(1)  # "要チェック"
        filtered = gui._get_filtered_indices()
        # ノーマーク(after='')のみが要チェック。複数マーク(after='-1')は修正済み
        assert len(filtered) == 1

    def test_sort_options(self, mark_checker_gui):
        """ソートプルダウンの選択肢が正しいこと"""
        gui = mark_checker_gui
        sort_values = list(gui._sort_combo['values'])
        assert "画像名順" in sort_values
        assert "白さ順（白い順）" in sort_values
