"""webui 追加機能（L1）: PDF展開・しきい値自動調整・氏名トリミング入り集計。"""
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


class TestPdfImport:
    def _make_pdf(self, path: Path, pages=2):
        import fitz
        doc = fitz.open()
        for i in range(pages):
            page = doc.new_page(width=595, height=842)
            page.insert_text((100, 100), f"Page {i + 1}")
        doc.save(str(path))
        doc.close()

    def test_single_pdf_extracts_and_sets_folder(self, tmp_path):
        pdf = tmp_path / "toi.pdf"
        self._make_pdf(pdf, pages=2)
        adapter = RecordingAdapter()
        b = Bridge(window_adapter=adapter)
        adapter.files_returns = [[str(pdf)]]
        assert b.select_pdf()["ok"]
        _wait_job_done(b)
        done = [e for e in adapter.events if e["type"] == "job_done"][-1]
        assert done["ok"], done
        # tk と同じ規則: {PDF名}_images/ に展開し、画像フォルダとして設定
        assert b.state["image_folder"].endswith("toi_images")
        assert b.state["image_count"] == 2

    def test_multiple_pdfs_share_common_folder(self, tmp_path):
        pdfs = []
        for name in ("a.pdf", "b.pdf"):
            p = tmp_path / name
            self._make_pdf(p, pages=1)
            pdfs.append(str(p))
        adapter = RecordingAdapter()
        b = Bridge(window_adapter=adapter)
        adapter.files_returns = [pdfs]
        assert b.select_pdf()["ok"]
        _wait_job_done(b)
        assert b.state["image_folder"].endswith("pdf_import_images")
        assert b.state["image_count"] == 2  # {PDF名}_pNNN.png で衝突しない

    def test_cancel_dialog(self):
        adapter = RecordingAdapter()
        b = Bridge(window_adapter=adapter)
        adapter.files_returns = [None]
        res = b.select_pdf()
        assert res["ok"] and res.get("cancelled")


def _build_marked_inputs(base: Path):
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
    cv2.imwrite(str(scans / "s2.png"), make_sheet({}, with_markers=True))
    return scans, coord, key


def _bridge_through_recognition(scans, coord, key):
    adapter = RecordingAdapter()
    b = Bridge(window_adapter=adapter)
    b.set_mode("mark_only", "multi_digit")
    b.set_skip_questions(0)
    adapter.folder_returns = [str(scans)]
    adapter.file_returns = [str(coord), str(key)]
    assert b.select_image_folder()["ok"]
    assert b.select_coord_file()["ok"]
    assert b.select_answer_key()["ok"]
    assert b.run_recognition()["ok"]
    _wait_job_done(b)
    return b, adapter


class TestThresholdCalibration:
    def test_calibration_applies_recommended_values(self, tmp_path):
        scans, coord, key = _build_marked_inputs(tmp_path)
        b, _ = _bridge_through_recognition(scans, coord, key)
        res = b.run_threshold_calibration()
        assert res["ok"], res
        assert 0 < res["color_threshold"] < 1
        assert 0 < res["area_threshold"] < 1
        assert b.state["color_threshold"] == res["color_threshold"]

    def test_calibration_requires_inputs(self):
        b = Bridge(window_adapter=RecordingAdapter())
        res = b.run_threshold_calibration()
        assert not res["ok"] and "画像フォルダ" in res["error"]


class TestNameTrimSummary:
    def test_region_validation(self, tmp_path):
        scans, coord, key = _build_marked_inputs(tmp_path)
        b, _ = _bridge_through_recognition(scans, coord, key)
        assert not b.set_name_trim_region([10, 10])["ok"]
        assert not b.set_name_trim_region([50, 50, 40, 90])["ok"]
        assert b.set_name_trim_region([20, 10, 300, 80])["ok"]
        assert b.set_name_trim_region(None)["ok"]
        assert b.state["name_trim_region"] is None

    def test_summary_with_name_images(self, tmp_path):
        scans, coord, key = _build_marked_inputs(tmp_path)
        b, adapter = _bridge_through_recognition(scans, coord, key)
        assert b.set_name_trim_region([20, 10, 300, 80])["ok"]
        for job in ("run_scoring", "run_summary"):
            assert getattr(b, job)()["ok"], job
            _wait_job_done(b)
            done = [e for e in adapter.events if e["type"] == "job_done"][-1]
            assert done["ok"], done
        # 集計シートに画像が埋め込まれている（openpyxl の _images で確認）
        import openpyxl
        from constants import (RESULTS_FOLDER, FINAL_REPORT_FOLDER,
                               STUDENT_SUMMARY_FILE)
        p = scans / RESULTS_FOLDER / FINAL_REPORT_FOLDER / STUDENT_SUMMARY_FILE
        wb = openpyxl.load_workbook(p)
        n_images = sum(len(ws._images) for ws in wb.worksheets)
        assert n_images >= 2, "氏名画像が集計シートに埋め込まれていない"

    def test_summary_without_region_still_works(self, tmp_path):
        """有効でも領域未指定なら氏名なしで普通に集計できる（エラーにしない）"""
        scans, coord, key = _build_marked_inputs(tmp_path)
        b, adapter = _bridge_through_recognition(scans, coord, key)
        for job in ("run_scoring", "run_summary"):
            assert getattr(b, job)()["ok"], job
            _wait_job_done(b)
            done = [e for e in adapter.events if e["type"] == "job_done"][-1]
            assert done["ok"], done

    def test_session_roundtrip_of_name_trim(self, tmp_path):
        scans, coord, key = _build_marked_inputs(tmp_path)
        b, _ = _bridge_through_recognition(scans, coord, key)
        b.set_name_trim_region([20, 10, 300, 80])
        b.set_name_trim_enabled(False)
        assert b.save_session()["ok"]
        from constants import (RESULTS_FOLDER, RESULTS_DATA_FOLDER,
                               SESSION_STATE_FILE)
        p = scans / RESULTS_FOLDER / RESULTS_DATA_FOLDER / SESSION_STATE_FILE
        b2 = Bridge(window_adapter=RecordingAdapter())
        res = b2.restore_session(str(p))
        assert res["ok"], res
        assert b2.state["name_trim_region"] == [20, 10, 300, 80]
        assert b2.state["name_trim_enabled"] is False
