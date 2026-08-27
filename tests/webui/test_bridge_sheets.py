"""SheetsMixin（全体画像・合計点位置）の L1 テスト。"""
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "webui"))
sys.path.insert(0, str(PROJECT_ROOT / "main_src"))
sys.path.insert(0, str(PROJECT_ROOT / "tests"))
sys.path.insert(0, str(PROJECT_ROOT / "tests" / "webui"))

from api.bridge import Bridge  # noqa: E402
from test_bridge_jobs import RecordingAdapter  # noqa: E402


@pytest.fixture
def bridge_with_boxed(tmp_path):
    """00_Processing に補正済み画像がある状態"""
    import cv2
    import numpy as np
    from constants import RESULTS_FOLDER, BOXED_FOLDER
    adapter = RecordingAdapter()
    b = Bridge(window_adapter=adapter)
    scans = tmp_path / "scans"
    boxed = scans / RESULTS_FOLDER / BOXED_FOLDER
    boxed.mkdir(parents=True)
    (scans / "orig.png").write_bytes(b"x")
    img = np.full((100, 200, 3), 255, dtype=np.uint8)
    cv2.imwrite(str(boxed / "b.png"), img)
    cv2.imwrite(str(boxed / "a.png"), img)
    adapter.folder_returns = [str(scans)]
    b.select_image_folder()
    return b


def test_sheet_image_returns_first_sorted(bridge_with_boxed):
    res = bridge_with_boxed.get_sheet_image()
    assert res["ok"] and res["filename"] == "a.png"
    assert res["data_url"].startswith("data:image/png;base64,")


def test_total_region_roundtrip(bridge_with_boxed):
    b = bridge_with_boxed
    assert b.get_total_display_region()["region"] is None
    assert b.set_total_display_region([10, 20, 110, 60])["ok"]
    assert b.get_total_display_region()["region"] == [10, 20, 110, 60]
    # tk 版と同じファイル・キーで保存されている（相互運用）
    from descriptive_scorer import load_total_display_config, TOTAL_DISPLAY_CONFIG_FILE
    cfg = load_total_display_config(str(b._results_data_folder() / TOTAL_DISPLAY_CONFIG_FILE))
    assert cfg["total_display_region"] == [10, 20, 110, 60]
    assert b.set_total_display_region(None)["ok"]
    assert b.get_total_display_region()["region"] is None


def test_invalid_region_rejected(bridge_with_boxed):
    assert not bridge_with_boxed.set_total_display_region([1, 2, 3])["ok"]


def test_get_sheet_size_returns_dimensions(bridge_with_boxed):
    res = bridge_with_boxed.get_sheet_size("a.png")
    assert res["ok"] and (res["w"], res["h"]) == (200, 100)


def test_get_sheet_size_missing_file(bridge_with_boxed):
    assert not bridge_with_boxed.get_sheet_size("nope.png")["ok"]


def test_get_sheet_size_rejects_path_traversal(bridge_with_boxed):
    # パス区切りはファイル名だけに丸められ、フォルダ外は読めない
    res = bridge_with_boxed.get_sheet_size("../../orig.png")
    assert not res["ok"]
