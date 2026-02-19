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
        assert gui._side_panel.winfo_manager() == ""

        gui._switch_to_grid()
        gui.window.update_idletasks()
        assert gui._view_mode == "grid"
        assert gui._grid_view_frame.winfo_manager() == "pack"
        assert gui._side_panel.winfo_manager() == "pack"

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

        # 要チェック（インデックス0）: error_type ありかつ after != -1
        gui._category_listbox.selection_clear(0, tk.END)
        gui._category_listbox.selection_set(0)  # "要チェック"
        filtered = gui._get_filtered_indices()
        # ノーマーク(after='')のみが要チェック。複数マーク(after='-1')は修正済み
        assert len(filtered) == 1

    def test_sort_options(self, mark_checker_gui):
        """ソートプルダウンの選択肢とデフォルト値が正しいこと"""
        gui = mark_checker_gui
        sort_values = list(gui._sort_combo['values'])
        assert "画像名順" in sort_values
        assert "白さ順（白い順）" in sort_values
        # デフォルトは白さ順
        assert gui._sort_var.get() == "白さ順（白い順）"

    def test_pager_always_pages_at_100(self, mark_checker_gui):
        """100件上限で常にページ分割されること"""
        gui = mark_checker_gui
        assert gui._grid_page_size == 100

        # 50件 → 1ページ
        gui._update_grid_pager(50)
        gui.window.update_idletasks()
        assert gui._page_label.cget("text") == "1/1"
        assert str(gui._btn_prev_page.cget("state")) == "disabled"
        assert str(gui._btn_next_page.cget("state")) == "disabled"

        # 250件 → 3ページ
        gui._update_grid_pager(250)
        gui.window.update_idletasks()
        assert gui._page_label.cget("text") == "1/3"
        assert str(gui._btn_next_page.cget("state")) == "normal"

    def test_pager_navigation(self, mark_checker_gui):
        """ページ移動が正しく動作すること"""
        gui = mark_checker_gui
        gui._grid_filtered_indices = list(range(250))
        gui._grid_current_page = 0
        gui._update_grid_pager(250)

        # 50件 → next は無効（1ページしかない）
        gui._grid_filtered_indices = list(range(50))
        gui._grid_current_page = 0
        gui._update_grid_pager(50)
        gui._next_grid_page()
        assert gui._grid_current_page == 0  # 進めない

        # prev は最初のページでは無効
        gui._prev_grid_page()
        assert gui._grid_current_page == 0

    def test_choice_tab_buttons_exist(self, mark_checker_gui):
        """選択肢カテゴリ用のタブボタンが存在すること"""
        gui = mark_checker_gui
        assert hasattr(gui, '_btn_tab_light')
        assert hasattr(gui, '_btn_tab_dark')
        assert hasattr(gui, '_tab_frame')
        assert hasattr(gui, '_pager_frame')
        # 初期状態ではタブは非表示、ページャーが表示
        assert gui._tab_frame.winfo_manager() == ""
        assert gui._pager_frame.winfo_manager() == "pack"

    def test_choice_tab_toggle_on_category_selection(self, mark_checker_gui):
        """選択肢カテゴリ選択時にタブモードに切り替わること"""
        gui = mark_checker_gui
        gui._all_entries_df = pd.DataFrame([
            {"filename": "a.jpg", "question_no": 1, "before": "1",
             "after": "", "error_type": "", "category": "1"},
            {"filename": "b.jpg", "question_no": 1, "before": "1",
             "after": "", "error_type": "", "category": "1"},
        ])
        gui._whiteness_cache = {0: 200.0, 1: 150.0}
        gui._rebuild_category_list()

        # 選択肢カテゴリを選択
        for i in range(gui._category_listbox.size()):
            text = gui._category_listbox.get(i)
            if '選択肢 1' in text:
                gui._category_listbox.selection_clear(0, tk.END)
                gui._category_listbox.selection_set(i)
                gui._on_category_selected()
                break

        gui.window.update_idletasks()
        assert gui._choice_tab_active is True
        assert gui._tab_frame.winfo_manager() == "pack"
        assert gui._pager_frame.winfo_manager() == ""

        # 非選択肢カテゴリに切替 → ページャーに戻る
        gui._category_listbox.selection_clear(0, tk.END)
        gui._category_listbox.selection_set(0)  # 要チェック
        gui._on_category_selected()
        gui.window.update_idletasks()
        assert gui._choice_tab_active is False
        assert gui._pager_frame.winfo_manager() == "pack"
        assert gui._tab_frame.winfo_manager() == ""

    def test_choice_tab_default_is_light(self, mark_checker_gui):
        """選択肢カテゴリ選択時のデフォルトタブが「薄い」であること"""
        gui = mark_checker_gui
        gui._all_entries_df = pd.DataFrame([
            {"filename": "a.jpg", "question_no": 1, "before": "1",
             "after": "", "error_type": "", "category": "1"},
        ])
        gui._whiteness_cache = {0: 200.0}
        gui._rebuild_category_list()

        for i in range(gui._category_listbox.size()):
            text = gui._category_listbox.get(i)
            if '選択肢 1' in text:
                gui._category_listbox.selection_clear(0, tk.END)
                gui._category_listbox.selection_set(i)
                gui._on_category_selected()
                break

        assert gui._choice_tab_current == "薄い"

    def test_apply_button_label(self, mark_checker_gui):
        """反映ボタンのラベルが変更されていること"""
        gui = mark_checker_gui
        assert gui._btn_apply.cget("text") == "データの更新(再読み込み)"

    def test_whiteness_json_loader(self, mark_checker_gui, tmp_path):
        """白さキャッシュJSONから正しく読み込めること"""
        import json
        gui = mark_checker_gui
        gui._all_entries_df = pd.DataFrame([
            {"filename": "a.jpg", "question_no": 1, "before": "1",
             "after": "", "error_type": "", "category": "1"},
            {"filename": "a.jpg", "question_no": 2, "before": "2",
             "after": "", "error_type": "", "category": "2"},
            {"filename": "b.jpg", "question_no": 1, "before": "1",
             "after": "", "error_type": "", "category": "1"},
        ])

        # JSONファイルを作成
        whiteness_data = {
            "a.jpg": {"1": 200.5, "2": 180.3},
            "b.jpg": {"1": 195.0},
        }
        json_path = tmp_path / "whiteness_cache.json"
        json_path.write_text(json.dumps(whiteness_data), encoding="utf-8")

        gui.coords_csv_path = tmp_path / "coords.csv"
        result = gui._load_whiteness_from_json()

        assert result is True
        assert gui._whiteness_cache[0] == 200.5
        assert gui._whiteness_cache[1] == 180.3
        assert gui._whiteness_cache[2] == 195.0
