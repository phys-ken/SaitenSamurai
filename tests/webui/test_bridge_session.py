"""webui セッション保存/復元（L1: tk 版 session_state.json 互換）。"""
import json
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "webui"))
sys.path.insert(0, str(PROJECT_ROOT / "main_src"))
sys.path.insert(0, str(PROJECT_ROOT / "tests"))
sys.path.insert(0, str(PROJECT_ROOT / "tests" / "webui"))

from api.bridge import Bridge  # noqa: E402
from test_bridge_jobs import RecordingAdapter, _wait_job_done  # noqa: E402
from test_multi_digit_mode import _create_answer_key  # noqa: E402
from test_multidigit_image_e2e import (  # noqa: E402
    make_coord_xlsx, make_sheet, sym_positions,
)


def _build_inputs(base: Path):
    import cv2
    coord = make_coord_xlsx(base / "coord.xlsx", 4)
    key = base / "key.xlsx"
    _create_answer_key(key, [
        {'問題番号': 1, '正答': '-24', '配点': 3, '観点': 1},
        {'問題番号': 4, '正答': 'a', '配点': 2, '観点': 2},
    ])
    scans = base / "scans"
    scans.mkdir()
    filled = {**{q: p for q, p in zip([1, 2, 3], sym_positions('-24'))},
              4: sym_positions('a')[0]}
    cv2.imwrite(str(scans / "s1.png"), make_sheet(filled, with_markers=True))
    return scans, coord, key


def _make_bridge(scans, coord, key, run_recognition=True):
    adapter = RecordingAdapter()
    b = Bridge(window_adapter=adapter)
    b.set_mode("mark_only", "multi_digit")
    b.set_skip_questions(0)
    adapter.folder_returns = [str(scans)]
    adapter.file_returns = [str(coord), str(key)]
    assert b.select_image_folder()["ok"]
    assert b.select_coord_file()["ok"]
    assert b.select_answer_key()["ok"]
    if run_recognition:
        assert b.run_recognition()["ok"]
        _wait_job_done(b)
    return b, adapter


def _session_file(scans):
    from constants import (RESULTS_FOLDER, RESULTS_DATA_FOLDER,
                           SESSION_STATE_FILE)
    return scans / RESULTS_FOLDER / RESULTS_DATA_FOLDER / SESSION_STATE_FILE


class TestSave:
    def test_autosave_after_select_and_job_tk_compatible_format(self, tmp_path):
        scans, coord, key = _build_inputs(tmp_path)
        b, _ = _make_bridge(scans, coord, key)
        p = _session_file(scans)
        assert p.exists(), "選択・認識の要所で自動保存されていない"
        data = json.loads(p.read_text(encoding="utf-8"))
        # tk 版 _save_session_state と同じキー・同じ表現
        assert data["version"] == 1
        assert data["app_mode"] == "mark_only"
        assert data["mark_format"] == "multi_digit"
        assert data["image_folder"] == str(scans)
        assert data["coord_excel"] == str(coord)  # 別ツリー → 絶対パスのまま
        assert isinstance(data["skip_questions"], str)  # tk は StringVar
        from constants import RESULTS_FOLDER
        assert data["omr_result"].startswith(RESULTS_FOLDER)  # フォルダ内 → 相対パス
        assert data["descriptive_enabled"] is False

    def test_tk_loader_can_read_webui_session(self, tmp_path):
        """tk 側のローダー（load_json_safe + version 必須）で読める"""
        from constants import load_json_safe
        scans, coord, key = _build_inputs(tmp_path)
        _make_bridge(scans, coord, key)
        data = load_json_safe(_session_file(scans), required_keys=["version"])
        assert data is not None and data["app_mode"] == "mark_only"


class TestRestore:
    def test_restore_roundtrip(self, tmp_path):
        scans, coord, key = _build_inputs(tmp_path)
        b1, _ = _make_bridge(scans, coord, key)
        omr = b1.state["omr_result"]

        b2 = Bridge(window_adapter=RecordingAdapter())
        res = b2.restore_session(str(_session_file(scans)))
        assert res["ok"], res
        assert res["warnings"] == []
        st = b2.state
        assert st["app_mode"] == "mark_only"
        assert st["mark_format"] == "multi_digit"
        assert st["image_folder"] == str(scans)
        assert st["coord_file"] == str(coord)
        assert st["answer_key"] == str(key)
        assert st["omr_result"] == omr
        assert st["skip_questions"] == 0
        assert st["coord_summary"] is not None
        assert st["key_summary"] is not None and st["key_summary"]["ok"]

    def test_restore_tk_written_session(self, tmp_path):
        """tk 版が書いた形式（相対パス・skip 文字列・webui キーなし）を読める"""
        scans, coord, key = _build_inputs(tmp_path)
        b1, _ = _make_bridge(scans, coord, key)
        p = _session_file(scans)
        data = json.loads(p.read_text(encoding="utf-8"))
        del data["webui"]                      # tk のファイルには無い
        data["skip_questions"] = "0"
        data["rendering_settings"] = {"circle_thickness": 3}
        p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")

        b2 = Bridge(window_adapter=RecordingAdapter())
        res = b2.restore_session(str(p))
        assert res["ok"], res
        assert b2.state["skip_questions"] == 0
        assert b2.state["omr_mode"] == "kmeans"  # 既定のまま

    def test_restore_reports_broken_paths_as_warnings(self, tmp_path):
        scans, coord, key = _build_inputs(tmp_path)
        _make_bridge(scans, coord, key)
        coord.unlink()  # 座標ファイルを消す
        b2 = Bridge(window_adapter=RecordingAdapter())
        res = b2.restore_session(str(_session_file(scans)))
        assert res["ok"]
        assert any("座標ファイル" in w for w in res["warnings"])
        assert b2.state["coord_file"] is None      # 壊れたパスは適用しない
        assert b2.state["answer_key"] == str(key)  # 生きているパスは適用する

    def test_restore_missing_image_folder_fails(self, tmp_path):
        scans, coord, key = _build_inputs(tmp_path)
        _make_bridge(scans, coord, key)
        p = _session_file(scans)
        data = json.loads(p.read_text(encoding="utf-8"))
        data["image_folder"] = str(tmp_path / "moved_away")
        p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        b2 = Bridge(window_adapter=RecordingAdapter())
        res = b2.restore_session(str(p))
        assert not res["ok"] and "画像フォルダ" in res["error"]

    def test_restore_dialog_cancel(self, tmp_path):
        adapter = RecordingAdapter()
        adapter.file_returns = [None]
        b = Bridge(window_adapter=adapter)
        res = b.restore_session()
        assert res["ok"] and res.get("cancelled")
