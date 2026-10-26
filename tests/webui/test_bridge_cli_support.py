"""外部のツール（コマンド）から採点侍を動かすために足した Bridge の振る舞い（L1）。

- 起動時にフォルダを指定して開く（app.py --folder → open_folder_session）
- 記述のみモードでも、四隅マーカーのある用紙は傾きを補正する
- 補正せずにコピーした答案には、高解像度の切り出し（補正後の座標）を当てない
"""
import shutil
import sys
from pathlib import Path

import pytest
from PIL import Image

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "webui"))
sys.path.insert(0, str(PROJECT_ROOT / "main_src"))
sys.path.insert(0, str(PROJECT_ROOT / "tests"))
sys.path.insert(0, str(PROJECT_ROOT / "tests" / "webui"))

from api.bridge import Bridge  # noqa: E402
from test_bridge_jobs import RecordingAdapter, _wait_job_done  # noqa: E402

SHEET = PROJECT_ROOT / "sample_basefile" / "sample_marksheet.jpg"


def _desc_bridge(scans):
    adapter = RecordingAdapter()
    b = Bridge(window_adapter=adapter)
    b.set_mode("descriptive_only", "standard")
    adapter.folder_returns = [str(scans)]
    assert b.select_image_folder()["ok"]
    return b, adapter


def _boxed(scans):
    from constants import RESULTS_FOLDER, BOXED_FOLDER
    return scans / RESULTS_FOLDER / BOXED_FOLDER


@pytest.fixture
def marker_scans(tmp_path):
    if not SHEET.exists():
        pytest.skip("サンプル用紙がありません")
    scans = tmp_path / "scans"
    scans.mkdir()
    shutil.copy(SHEET, scans / "a.jpg")
    shutil.copy(SHEET, scans / "b.jpg")
    return scans


class TestDescriptiveOnlyDeskew:
    def test_marker_sheets_are_corrected(self, marker_scans):
        b, adapter = _desc_bridge(marker_scans)
        assert b.run_prepare_images()["ok"]
        _wait_job_done(b)
        for p in _boxed(marker_scans).iterdir():
            with Image.open(p) as im:
                assert im.size == (595, 842), p.name
        done = [e for e in adapter.events if e["type"] == "job_done"][-1]
        assert "傾きを補正" in done["message"]

    def test_markerless_sheets_are_copied(self, tmp_path):
        import cv2
        import numpy as np
        scans = tmp_path / "scans"
        scans.mkdir()
        img = np.full((400, 300, 3), 255, np.uint8)
        cv2.putText(img, "answer", (40, 200), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 0), 2)
        cv2.imwrite(str(scans / "s1.png"), img)
        b, _ = _desc_bridge(scans)
        assert b.run_prepare_images()["ok"]
        _wait_job_done(b)
        with Image.open(_boxed(scans) / "s1.png") as im:
            assert im.size == (300, 400)

    def test_previously_copied_folder_stays_copied(self, marker_scans):
        """以前の版で補正せずに準備した試験は、範囲がずれないようコピーのまま続ける"""
        boxed = _boxed(marker_scans)
        boxed.mkdir(parents=True)
        shutil.copy2(marker_scans / "a.jpg", boxed / "a.jpg")
        b, _ = _desc_bridge(marker_scans)
        assert b.run_prepare_images()["ok"]
        _wait_job_done(b)
        for p in boxed.iterdir():
            with Image.open(p) as im:
                assert im.size == Image.open(SHEET).size, p.name

    def test_copied_sheet_crops_use_its_own_coordinates(self, marker_scans):
        """コピーした答案（元の大きさの座標）に、補正後の高解像度切り出しを当てない。

        当てると範囲が補正後の画像の倍率で拡大され、全く違う所を切り出していた
        """
        boxed = _boxed(marker_scans)
        boxed.mkdir(parents=True)
        for name in ("a.jpg", "b.jpg"):
            shutil.copy2(marker_scans / name, boxed / name)
        b, _ = _desc_bridge(marker_scans)
        region = [100, 300, 700, 600]
        assert b.add_descriptive_question("問1", 2, 1, region)["ok"]
        assert b.start_descriptive_scoring()["ok"]
        crop = b._desc_crops["D1"]["a.jpg"]
        with Image.open(crop) as im:
            assert im.size == (600, 300)
        # 中身も元画像の同じ場所と一致する
        import numpy as np
        with Image.open(SHEET) as src, Image.open(crop) as im:
            expected = np.asarray(src.convert("L").crop(region), dtype=np.int16)
            got = np.asarray(im.convert("L"), dtype=np.int16)
        assert np.abs(expected - got).mean() < 8


class TestBulkScores:
    def test_all_or_nothing(self, marker_scans):
        b, _ = _desc_bridge(marker_scans)
        assert b.run_prepare_images()["ok"]
        _wait_job_done(b)
        assert b.add_descriptive_question("問1", 2, 1, [10, 10, 100, 100])["ok"]
        # 1件でも範囲外があれば何も書かない
        res = b.set_descriptive_scores_bulk("D1", {"a.jpg": 2, "b.jpg": 5})
        assert not res["ok"]
        assert b._desc_scores["scores"] == {}
        res = b.set_descriptive_scores_bulk("D1", {"a.jpg": 2, "b.jpg": None})
        assert res["ok"] and res["count"] == 2
        from descriptive_scorer import load_descriptive_scores
        saved = load_descriptive_scores(str(b._desc_scores_path()))
        assert saved["scores"]["a.jpg"] == {"D1": 2}
        assert saved["scores"].get("b.jpg", {}) == {}

    def test_unknown_sheet_is_rejected(self, marker_scans):
        b, _ = _desc_bridge(marker_scans)
        assert b.run_prepare_images()["ok"]
        _wait_job_done(b)
        assert b.add_descriptive_question("問1", 2, 1, [10, 10, 100, 100])["ok"]
        assert not b.set_descriptive_scores_bulk("D1", {"../x.jpg": 1})["ok"]


class TestOpenFolderSession:
    def test_opens_saved_session(self, marker_scans):
        b, _ = _desc_bridge(marker_scans)
        assert b.save_session()["ok"]
        b2 = Bridge(window_adapter=RecordingAdapter())
        res = b2.open_folder_session(str(marker_scans))
        assert res["ok"], res
        assert b2.state["opened_from_command_line"] is True
        assert b2.state["app_mode"] == "descriptive_only"
        assert b2.state["image_folder"] == str(marker_scans)

    def test_accepts_exam_folder_with_scans_inside(self, marker_scans):
        b, _ = _desc_bridge(marker_scans)
        assert b.save_session()["ok"]
        b2 = Bridge(window_adapter=RecordingAdapter())
        assert b2.open_folder_session(str(marker_scans.parent))["ok"]
        assert b2.state["image_folder"] == str(marker_scans)

    def test_folder_without_session_is_an_error(self, tmp_path):
        b = Bridge(window_adapter=RecordingAdapter())
        res = b.open_folder_session(str(tmp_path))
        assert not res["ok"] and "session_state.json" in res["error"]
