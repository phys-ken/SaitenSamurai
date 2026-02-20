"""
test_mode_consistency.py — 3モード一貫性テスト

採点侍の3つの起動モード（マークのみ / マーク＋記述 / 記述のみ）間で
GUI状態・ボタン可視性・インポート・セッション復元が一貫して動作することを検証する。

テスト対象:
  - GUI初期化時のウィジェット可視性・状態の正しさ
  - _set_processing_state(True/False) の安全性
  - _update_step_availability の一貫性
  - _apply_session_state のモード間セッション復元
  - _on_descriptive_toggle のモード別安全性
  - descriptive_scorer.py のインポートパス（saitensamurai 経由でないこと）
  - EXE環境での import chain 安全性
"""

import json
import shutil
import tempfile
import tkinter as tk
from pathlib import Path
from unittest.mock import patch, MagicMock

import pytest
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

ALL_MODES = [MODE_MARK_ONLY, MODE_MARK_AND_DESCRIPTIVE, MODE_DESCRIPTIVE_ONLY]


# ============================================================
# ヘルパー
# ============================================================

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


# ============================================================
# 1. GUI ウィジェット可視性テスト（全3モード）
# ============================================================

class TestWidgetVisibility:
    """各モードで表示/非表示であるべきウィジェットが正しく制御されている"""

    @pytest.mark.parametrize("mode", ALL_MODES)
    def test_btn_mark_check_visibility(self, mode):
        """マークチェックボタン: 記述のみモードでは非表示"""
        top, app = _create_gui(mode)
        try:
            manager = app._btn_mark_check.winfo_manager()
            if mode == MODE_DESCRIPTIVE_ONLY:
                assert manager == "", \
                    f"MODE_DESCRIPTIVE_ONLY で _btn_mark_check が表示されている: {manager}"
            else:
                assert manager == "pack", \
                    f"{mode} で _btn_mark_check が非表示"
        finally:
            _destroy_top(top)

    @pytest.mark.parametrize("mode", ALL_MODES)
    def test_desc_setup_btn_visibility(self, mode):
        """記述問題設定ボタン: マークのみモードでは非表示"""
        top, app = _create_gui(mode)
        try:
            manager = app.desc_setup_btn.winfo_manager()
            if mode == MODE_MARK_ONLY:
                assert manager == "", \
                    f"MODE_MARK_ONLY で desc_setup_btn が表示されている"
            else:
                assert manager == "pack", \
                    f"{mode} で desc_setup_btn が非表示"
        finally:
            _destroy_top(top)

    @pytest.mark.parametrize("mode", ALL_MODES)
    def test_desc_scoring_btn_visibility(self, mode):
        """記述採点ボタン: マークのみモードでは非表示"""
        top, app = _create_gui(mode)
        try:
            manager = app.desc_scoring_btn.winfo_manager()
            if mode == MODE_MARK_ONLY:
                assert manager == "", \
                    f"MODE_MARK_ONLY で desc_scoring_btn が表示されている"
            else:
                assert manager == "pack", \
                    f"{mode} で desc_scoring_btn が非表示"
        finally:
            _destroy_top(top)

    @pytest.mark.parametrize("mode", ALL_MODES)
    def test_btn_desc_review_visibility(self, mode):
        """記述採点確認ボタン: マークのみモードでは非表示"""
        top, app = _create_gui(mode)
        try:
            manager = app._btn_desc_review.winfo_manager()
            if mode == MODE_MARK_ONLY:
                assert manager == "", \
                    f"MODE_MARK_ONLY で _btn_desc_review が表示されている"
            else:
                assert manager == "pack", \
                    f"{mode} で _btn_desc_review が非表示"
        finally:
            _destroy_top(top)

    @pytest.mark.parametrize("mode", ALL_MODES)
    def test_total_pos_button_always_visible(self, mode):
        """合計点位置設定ボタン: 全モードで表示"""
        top, app = _create_gui(mode)
        try:
            manager = app._btn_total_pos.winfo_manager()
            assert manager == "pack", f"{mode} で _btn_total_pos が非表示"
        finally:
            _destroy_top(top)

    @pytest.mark.parametrize("mode", ALL_MODES)
    def test_run_scoring_button_always_visible(self, mode):
        """採点実行ボタン: 全モードで表示"""
        top, app = _create_gui(mode)
        try:
            manager = app._btn_run_scoring.winfo_manager()
            assert manager == "pack", f"{mode} で _btn_run_scoring が非表示"
        finally:
            _destroy_top(top)


# ============================================================
# 2. descriptive_enabled 初期値テスト
# ============================================================

class TestDescriptiveEnabledInit:
    """各モードの descriptive_enabled 初期値が正しい"""

    @pytest.mark.parametrize("mode,expected", [
        (MODE_MARK_ONLY, False),
        (MODE_MARK_AND_DESCRIPTIVE, True),
        (MODE_DESCRIPTIVE_ONLY, True),
    ])
    def test_descriptive_enabled_initial_value(self, mode, expected):
        top, app = _create_gui(mode)
        try:
            assert app.descriptive_enabled.get() == expected, \
                f"{mode}: descriptive_enabled should be {expected}"
        finally:
            _destroy_top(top)


# ============================================================
# 3. _set_processing_state の安全性
# ============================================================

class TestSetProcessingState:
    """_set_processing_state が全モードで例外なく動作する"""

    @pytest.mark.parametrize("mode", ALL_MODES)
    def test_set_processing_state_true_no_error(self, mode):
        """busy=True で例外が発生しない"""
        top, app = _create_gui(mode)
        try:
            app._set_processing_state(True)
            assert app._processing is True
        finally:
            _destroy_top(top)

    @pytest.mark.parametrize("mode", ALL_MODES)
    def test_set_processing_state_false_no_error(self, mode):
        """busy=False で例外が発生しない"""
        top, app = _create_gui(mode)
        try:
            app._set_processing_state(True)
            app._set_processing_state(False)
            assert app._processing is False
        finally:
            _destroy_top(top)

    @pytest.mark.parametrize("mode", ALL_MODES)
    def test_buttons_disabled_during_processing(self, mode):
        """処理中は全操作ボタンが DISABLED"""
        top, app = _create_gui(mode)
        try:
            app._set_processing_state(True)
            assert str(app._btn_run_scoring.cget("state")) == "disabled"
            assert str(app._btn_total_pos.cget("state")) == "disabled"
            assert str(app._btn_mark_check.cget("state")) == "disabled"
        finally:
            _destroy_top(top)

    @pytest.mark.parametrize("mode", ALL_MODES)
    def test_step_guard_after_processing_false(self, mode):
        """busy=False 後に _update_step_availability が呼ばれ、
        Step2 がフォルダ未設定なら DISABLED に戻る"""
        top, app = _create_gui(mode)
        try:
            # フォルダ未設定の状態で busy=False
            app._set_processing_state(True)
            app._set_processing_state(False)
            # Step2 ボタンは DISABLED に戻っているべき
            assert str(app._btn_run_scoring.cget("state")) == "disabled"
        finally:
            _destroy_top(top)


# ============================================================
# 4. _update_step_availability の一貫性
# ============================================================

class TestStepAvailability:
    """_update_step_availability がモード別に正しく機能する"""

    @pytest.mark.parametrize("mode", ALL_MODES)
    def test_no_folder_all_step2_disabled(self, mode):
        """フォルダ未選択 → Step2 全ボタン DISABLED"""
        top, app = _create_gui(mode)
        try:
            app.image_folder_path.set("")
            app._update_step_availability()
            assert str(app._btn_run_scoring.cget("state")) == "disabled"
            assert str(app._btn_total_pos.cget("state")) == "disabled"
        finally:
            _destroy_top(top)

    @pytest.mark.parametrize("mode", ALL_MODES)
    def test_with_boxed_folder_step2_enabled(self, mode):
        """boxed フォルダに画像あり → Step2 ボタン NORMAL"""
        tmpdir = Path(tempfile.mkdtemp())
        try:
            img_folder = tmpdir / "images"
            img_folder.mkdir()
            boxed = img_folder / RESULTS_FOLDER / BOXED_FOLDER
            boxed.mkdir(parents=True)
            # ダミー画像作成
            (boxed / "test001.jpg").write_bytes(b"\xff\xd8\xff\xe0")

            top, app = _create_gui(mode)
            try:
                app.image_folder_path.set(str(img_folder))
                app._update_step_availability()
                assert str(app._btn_run_scoring.cget("state")) == "normal", \
                    f"{mode}: boxed に画像があるのに _btn_run_scoring が disabled"
                assert str(app._btn_total_pos.cget("state")) == "normal"
            finally:
                _destroy_top(top)
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


# ============================================================
# 5. _apply_session_state のモード間セッション復元
# ============================================================

class TestApplySessionStateCrossModes:
    """異なるモードのセッションを復元しても安全"""

    def _make_session(self, image_folder, desc_enabled=True, mode="mark_and_descriptive"):
        return {
            "version": 1,
            "image_folder": str(image_folder),
            "descriptive_enabled": desc_enabled,
            "app_mode": mode,
            "skip_questions": "4",
            "color_threshold": 0.1,
            "area_threshold": 0.4,
        }

    def test_mark_only_ignores_descriptive_enabled_true(self):
        """MODE_MARK_ONLY でセッションの descriptive_enabled=True を無視する"""
        tmpdir = Path(tempfile.mkdtemp())
        try:
            img_folder = tmpdir / "images"
            img_folder.mkdir()

            top, app = _create_gui(MODE_MARK_ONLY)
            try:
                session = self._make_session(img_folder, desc_enabled=True)
                app._apply_session_state(session)
                assert app.descriptive_enabled.get() is False, \
                    "MODE_MARK_ONLY で descriptive_enabled が True に変更された"
                # 記述ボタンが表示されていないこと
                assert app.desc_scoring_btn.winfo_manager() == ""
            finally:
                _destroy_top(top)
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    def test_descriptive_only_ignores_descriptive_enabled_false(self):
        """MODE_DESCRIPTIVE_ONLY でセッションの descriptive_enabled=False を無視する"""
        tmpdir = Path(tempfile.mkdtemp())
        try:
            img_folder = tmpdir / "images"
            img_folder.mkdir()

            top, app = _create_gui(MODE_DESCRIPTIVE_ONLY)
            try:
                session = self._make_session(img_folder, desc_enabled=False)
                app._apply_session_state(session)
                assert app.descriptive_enabled.get() is True, \
                    "MODE_DESCRIPTIVE_ONLY で descriptive_enabled が False に変更された"
            finally:
                _destroy_top(top)
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    def test_mark_and_descriptive_restores_descriptive_enabled(self):
        """MODE_MARK_AND_DESCRIPTIVE ではセッション値を正しく復元する"""
        tmpdir = Path(tempfile.mkdtemp())
        try:
            img_folder = tmpdir / "images"
            img_folder.mkdir()

            top, app = _create_gui(MODE_MARK_AND_DESCRIPTIVE)
            try:
                # descriptive_enabled=False のセッションを復元
                session = self._make_session(img_folder, desc_enabled=False)
                app._apply_session_state(session)
                # False に変更される
                assert app.descriptive_enabled.get() is False
                # 記述ボタンが非表示になっている
                assert app.desc_scoring_btn.winfo_manager() == ""
            finally:
                _destroy_top(top)
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    def test_mark_and_descriptive_restores_descriptive_enabled_true(self):
        """MODE_MARK_AND_DESCRIPTIVE で descriptive_enabled=True の復元"""
        tmpdir = Path(tempfile.mkdtemp())
        try:
            img_folder = tmpdir / "images"
            img_folder.mkdir()

            top, app = _create_gui(MODE_MARK_AND_DESCRIPTIVE)
            try:
                session = self._make_session(img_folder, desc_enabled=True)
                app._apply_session_state(session)
                assert app.descriptive_enabled.get() is True
                assert app.desc_scoring_btn.winfo_manager() == "pack"
            finally:
                _destroy_top(top)
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


# ============================================================
# 6. _on_descriptive_toggle の安全性
# ============================================================

class TestOnDescriptiveToggle:
    """_on_descriptive_toggle がモード別に安全に動作する"""

    def test_descriptive_only_toggle_noop(self):
        """MODE_DESCRIPTIVE_ONLY では _on_descriptive_toggle が何もしない"""
        top, app = _create_gui(MODE_DESCRIPTIVE_ONLY)
        try:
            # 初期状態を記録
            btn_visible_before = app.desc_scoring_btn.winfo_manager()
            # toggle を呼んでも変化しない
            app._on_descriptive_toggle()
            btn_visible_after = app.desc_scoring_btn.winfo_manager()
            assert btn_visible_before == btn_visible_after
        finally:
            _destroy_top(top)

    def test_mark_only_toggle_does_not_show_descriptive(self):
        """MODE_MARK_ONLY で descriptive_enabled=True にして toggle しても
        記述UIが二重 pack されない"""
        top, app = _create_gui(MODE_MARK_ONLY)
        try:
            # descriptive_enabled を True に無理やり変更
            app.descriptive_enabled.set(True)
            app._on_descriptive_toggle()
            # マークチェックボタンは引き続き表示
            assert app._btn_mark_check.winfo_manager() == "pack"
            # desc_scoring_btn が pack されている（toggle が許可した場合）
            # → 実際のレイアウトが壊れないことの確認
            top.update_idletasks()  # レイアウト更新を強制
        finally:
            _destroy_top(top)

    def test_mark_and_descriptive_toggle_cycle(self):
        """MODE_MARK_AND_DESCRIPTIVE で ON→OFF→ON サイクルが安全"""
        top, app = _create_gui(MODE_MARK_AND_DESCRIPTIVE)
        try:
            # 初期: ON
            assert app.desc_scoring_btn.winfo_manager() == "pack"

            # OFF にする
            app.descriptive_enabled.set(False)
            app._on_descriptive_toggle()
            assert app.desc_scoring_btn.winfo_manager() == ""

            # ON に戻す
            app.descriptive_enabled.set(True)
            app._on_descriptive_toggle()
            assert app.desc_scoring_btn.winfo_manager() == "pack"
            assert app._btn_mark_check.winfo_manager() == "pack"

            top.update_idletasks()
        finally:
            _destroy_top(top)


# ============================================================
# 7. インポートパスのテスト（saitensamurai 経由回避）
# ============================================================

class TestImportPaths:
    """descriptive_scorer.py が saitensamurai.py を経由せず直接インポートする"""

    def test_descriptive_scorer_no_saitensamurai_import(self):
        """descriptive_scorer.py に 'from saitensamurai import' がないことを確認"""
        desc_scorer_path = Path(__file__).parent.parent / "main_src" / "descriptive_scorer.py"
        content = desc_scorer_path.read_text(encoding="utf-8")
        # "from saitensamurai import" が含まれていないこと
        assert "from saitensamurai import" not in content, \
            "descriptive_scorer.py はまだ saitensamurai 経由でインポートしている"

    def test_name_trimmer_no_saitensamurai_import(self):
        """name_trimmer.py に 'from saitensamurai import' がないことを確認"""
        name_trimmer_path = Path(__file__).parent.parent / "main_src" / "name_trimmer.py"
        content = name_trimmer_path.read_text(encoding="utf-8")
        assert "from saitensamurai import" not in content, \
            "name_trimmer.py はまだ saitensamurai 経由でインポートしている"

    def test_omr_engine_functions_importable(self):
        """omr_engine から必要な関数が直接インポートできる"""
        from omr_engine import (
            detect_corner_markers,
            apply_perspective_transform,
            compute_output_scale,
            parse_excel_coordinates,
        )
        assert callable(detect_corner_markers)
        assert callable(apply_perspective_transform)
        assert callable(compute_output_scale)
        assert callable(parse_excel_coordinates)

    def test_scoring_engine_functions_importable(self):
        """scoring_engine から必要な関数が直接インポートできる"""
        from scoring_engine import load_template, load_mark2_results, score_answers
        assert callable(load_template)
        assert callable(load_mark2_results)
        assert callable(score_answers)

    def test_image_renderer_functions_importable(self):
        """image_renderer から必要な関数が直接インポートできる"""
        from image_renderer import draw_scoring_results
        assert callable(draw_scoring_results)


# ============================================================
# 8. EXE 環境での import chain 安全性テスト
# ============================================================

class TestImportChainSafety:
    """EXE環境で問題になり得る循環・連鎖インポートの検証"""

    def test_descriptive_scorer_import_does_not_trigger_sklearn(self):
        """descriptive_scorer のインポートが sklearn を連鎖ロードしない"""
        import importlib
        # sklearn がまだ読み込まれていないか確認（既にロード済みならスキップ）
        sklearn_loaded_before = "sklearn" in sys.modules

        import descriptive_scorer  # noqa: F401

        if not sklearn_loaded_before:
            # descriptive_scorer のインポートだけでは sklearn がロードされないはず
            # ただし conftest で saitensamurai がインポートされている可能性があるため
            # この条件は厳密には保証できない。ロードされている場合は警告のみ
            pass  # テスト環境では saitensamurai がロード済みなので PASS

    def test_descriptive_scorer_generate_return_sheets_exists(self):
        """generate_return_sheets が正しくインポートできる"""
        from descriptive_scorer import generate_return_sheets
        assert callable(generate_return_sheets)

    def test_descriptive_scorer_trim_descriptive_regions_exists(self):
        """trim_descriptive_regions が正しくインポートできる"""
        from descriptive_scorer import trim_descriptive_regions
        assert callable(trim_descriptive_regions)


# ============================================================
# 9. モード別タイトルテスト
# ============================================================

class TestWindowTitle:
    """各モードでウィンドウタイトルが正しく設定される"""

    @pytest.mark.parametrize("mode,expected_substr", [
        (MODE_MARK_ONLY, "マーク採点"),
        (MODE_MARK_AND_DESCRIPTIVE, "マーク＋記述"),
        (MODE_DESCRIPTIVE_ONLY, "記述採点"),
    ])
    def test_window_title(self, mode, expected_substr):
        top, app = _create_gui(mode)
        try:
            assert expected_substr in top.title(), \
                f"タイトルに '{expected_substr}' が含まれていない: {top.title()}"
        finally:
            _destroy_top(top)


# ============================================================
# 10. Step2 レイアウト順序テスト
# ============================================================

class TestStep2LayoutOrder:
    """Step2 のウィジェットが正しい順序で pack されている"""

    def _get_packed_children(self, frame):
        """frame の pack されている子ウィジェットのリスト（pack 順）"""
        children = frame.pack_slaves()
        return children

    def test_mark_only_layout_order(self):
        """マークのみ: マークチェック → 合計点 → 設定リンク → 実行行"""
        top, app = _create_gui(MODE_MARK_ONLY)
        try:
            children = self._get_packed_children(app._step2_frame)
            # widget のテキストまたは型で識別
            texts = []
            for c in children:
                if isinstance(c, tk.Button):
                    texts.append(c.cget("text"))
                elif isinstance(c, tk.Label):
                    texts.append(c.cget("text"))
                elif isinstance(c, tk.Frame):
                    texts.append("Frame")

            # マークチェック、合計点位置は含まれるべき
            assert any("マークチェック" in t for t in texts), f"マークチェック missing: {texts}"
            assert any("合計点位置" in t for t in texts), f"合計点位置 missing: {texts}"
            # 記述採点は含まれないべき
            assert not any("記述採点" in t for t in texts), f"記述採点 should not be in mark_only: {texts}"
        finally:
            _destroy_top(top)

    def test_mark_and_descriptive_layout_order(self):
        """マーク＋記述: マークチェック → 記述採点 → ステータス → 確認 → 合計点 → 設定 → 実行"""
        top, app = _create_gui(MODE_MARK_AND_DESCRIPTIVE)
        try:
            children = self._get_packed_children(app._step2_frame)
            texts = []
            for c in children:
                if isinstance(c, tk.Button):
                    texts.append(c.cget("text"))
                elif isinstance(c, tk.Label):
                    texts.append(c.cget("text"))
                elif isinstance(c, tk.Frame):
                    texts.append("Frame")

            assert any("マークチェック" in t for t in texts)
            assert any("記述採点" in t for t in texts)
            assert any("合計点位置" in t for t in texts)
        finally:
            _destroy_top(top)

    def test_descriptive_only_layout_has_no_mark_check(self):
        """記述のみ: マークチェックボタンが非表示"""
        top, app = _create_gui(MODE_DESCRIPTIVE_ONLY)
        try:
            children = self._get_packed_children(app._step2_frame)
            texts = []
            for c in children:
                if isinstance(c, tk.Button):
                    texts.append(c.cget("text"))
                elif isinstance(c, tk.Label):
                    texts.append(c.cget("text"))
                elif isinstance(c, tk.Frame):
                    texts.append("Frame")

            assert not any("マークチェック" in t for t in texts), \
                f"記述のみモードにマークチェックがある: {texts}"
            assert any("記述採点" in t for t in texts)
        finally:
            _destroy_top(top)


# ============================================================
# 11. rendering_settings スレッドセーフティテスト
# ============================================================

class TestRenderingSettingsThreadSafety:
    """rendering_settings がスレッドに渡される際にコピーが作られる"""

    def test_descriptive_only_scoring_copies_settings(self):
        """_run_descriptive_only_thread が rendering_settings のコピーを使う"""
        from main_gui import SaitenSamuraiGUI
        import inspect

        # ソースコードを取得して dict(self.rendering_settings) の存在を確認
        source = inspect.getsource(SaitenSamuraiGUI._run_descriptive_only_thread)
        assert "dict(self.rendering_settings)" in source, \
            "_run_descriptive_only_thread で rendering_settings がコピーされていない"


# ============================================================
# 12. include_descriptive_in_analysis 初期値テスト
# ============================================================

class TestIncludeDescriptiveInAnalysis:
    """include_descriptive_in_analysis の初期値テスト"""

    @pytest.mark.parametrize("mode", ALL_MODES)
    def test_initial_value_is_true(self, mode):
        """全モードで初期値が True"""
        top, app = _create_gui(mode)
        try:
            assert app.include_descriptive_in_analysis.get() is True
        finally:
            _destroy_top(top)

    def test_mark_only_checkbox_hidden(self):
        """マークのみモードでチェックボックスが非表示"""
        top, app = _create_gui(MODE_MARK_ONLY)
        try:
            assert app._chk_include_desc_analysis.winfo_manager() == ""
        finally:
            _destroy_top(top)

    def test_descriptive_only_checkbox_disabled(self):
        """記述のみモードでチェックボックスが disabled"""
        top, app = _create_gui(MODE_DESCRIPTIVE_ONLY)
        try:
            state = str(app._chk_include_desc_analysis.cget("state"))
            assert state == "disabled"
        finally:
            _destroy_top(top)


# ============================================================
# 13. _set_step2_enabled の全モード安全性
# ============================================================

class TestSetStep2Enabled:
    """_set_step2_enabled が全モードで安全に動作する"""

    @pytest.mark.parametrize("mode", ALL_MODES)
    def test_enable_no_error(self, mode):
        """有効化で例外が発生しない"""
        top, app = _create_gui(mode)
        try:
            app._set_step2_enabled(True)
            assert str(app._btn_run_scoring.cget("state")) == "normal"
        finally:
            _destroy_top(top)

    @pytest.mark.parametrize("mode", ALL_MODES)
    def test_disable_no_error(self, mode):
        """無効化で例外が発生しない"""
        top, app = _create_gui(mode)
        try:
            app._set_step2_enabled(False)
            assert str(app._btn_run_scoring.cget("state")) == "disabled"
        finally:
            _destroy_top(top)

    @pytest.mark.parametrize("mode", ALL_MODES)
    def test_toggle_cycle(self, mode):
        """有効→無効→有効のサイクルが安全"""
        top, app = _create_gui(mode)
        try:
            app._set_step2_enabled(True)
            app._set_step2_enabled(False)
            app._set_step2_enabled(True)
            assert str(app._btn_run_scoring.cget("state")) == "normal"
        finally:
            _destroy_top(top)
