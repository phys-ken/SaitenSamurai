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


def _has_pymupdf():
    from constants import HAS_PYMUPDF
    return HAS_PYMUPDF


@pytest.mark.skipif(not _has_pymupdf(), reason="PyMuPDF 未インストール（オプショナル依存）")
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


class TestParityAdditions:
    def test_auto_detect_answer_key_and_session_flag(self, tmp_path):
        """フォルダ選択で正答データ自動検出＋既存セッション検出（tk 相当）"""
        scans, coord, key = _build_marked_inputs(tmp_path)
        b1, _ = _bridge_through_recognition(scans, coord, key)
        # answer_key.xlsx を結果フォルダに配置（tk の標準運用と同じ場所）
        from constants import (RESULTS_FOLDER, RESULTS_DATA_FOLDER,
                               ANSWER_KEY_FILE)
        data_dir = scans / RESULTS_FOLDER / RESULTS_DATA_FOLDER
        import shutil
        shutil.copy(key, data_dir / ANSWER_KEY_FILE)

        b2 = Bridge(window_adapter=RecordingAdapter())
        b2.set_mode("mark_only", "multi_digit")
        b2._win.folder_returns = [str(scans)]
        res = b2.select_image_folder()
        assert res["ok"]
        assert res["auto_detected_answer_key"] and \
            res["auto_detected_answer_key"].endswith(ANSWER_KEY_FILE)
        assert b2.state["answer_key"].endswith(ANSWER_KEY_FILE)
        assert b2.state["key_summary"] is not None
        assert res["session_found"] and res["session_found"].endswith(
            "session_state.json")

    def test_recheck_answer_key(self, tmp_path):
        scans, coord, key = _build_marked_inputs(tmp_path)
        b, _ = _bridge_through_recognition(scans, coord, key)
        assert b.recheck_answer_key()["ok"]
        b2 = Bridge(window_adapter=RecordingAdapter())
        assert "正答データ" in b2.recheck_answer_key()["error"]

    def test_include_descriptive_setting_and_session(self, tmp_path):
        scans, coord, key = _build_marked_inputs(tmp_path)
        b, _ = _bridge_through_recognition(scans, coord, key)
        assert b.state["include_descriptive_in_analysis"] is True
        assert b.set_include_descriptive_in_analysis(False)["ok"]
        assert b.save_session()["ok"]
        from constants import (RESULTS_FOLDER, RESULTS_DATA_FOLDER,
                               SESSION_STATE_FILE)
        p = scans / RESULTS_FOLDER / RESULTS_DATA_FOLDER / SESSION_STATE_FILE
        b2 = Bridge(window_adapter=RecordingAdapter())
        assert b2.restore_session(str(p))["ok"]
        assert b2.state["include_descriptive_in_analysis"] is False

    def test_checker_whiteness_and_batch_minus1(self, tmp_path):
        """白さがエントリに付き、ノーマーク一括-1がCSV保存1回で効く"""
        import cv2
        import numpy as np
        coord = make_coord_xlsx(tmp_path / "coord.xlsx", 3)
        key = tmp_path / "key.xlsx"
        _create_answer_key(key, [
            {'問題番号': 1, '正答': '-24', '配点': 3, '観点': 1},
        ])
        scans = tmp_path / "scans"
        scans.mkdir()
        # s1: 完答 / s2: 全行ノーマーク
        filled = {q: p for q, p in zip([1, 2, 3], sym_positions('-24'))}
        cv2.imwrite(str(scans / "s1.png"), make_sheet(filled, with_markers=True))
        cv2.imwrite(str(scans / "s2.png"), make_sheet({}, with_markers=True))
        b, _ = _bridge_through_recognition(scans, coord, key)
        assert b.open_mark_checker()["ok"]
        entries = b._checker["entries"]
        assert all("whiteness" in e for e in entries)
        # 完答者と白紙で白さが分かれている（白紙のほうが白い）
        w_marked = [e["whiteness"] for e in entries if e["filename"] == "s1.png"]
        w_blank = [e["whiteness"] for e in entries if e["filename"] == "s2.png"]
        assert max(w_blank) >= max(w_marked)

        no_marks = [e for e in entries if e["category"] == "ノーマーク"]
        assert no_marks, "ノーマークのエントリが無いとこのテストは成立しない"
        res = b.batch_correct_no_mark()
        assert res["ok"] and res["applied"] == len(no_marks)
        assert all(e["after"] == "-1" for e in no_marks)
        assert res["state"]["checker"]["corrected"] >= len(no_marks)

    def test_open_folder_validation(self, tmp_path):
        b = Bridge(window_adapter=RecordingAdapter())
        assert "画像フォルダ" in b.open_folder("boxed")["error"]
        scans, coord, key = _build_marked_inputs(tmp_path)
        b2, _ = _bridge_through_recognition(scans, coord, key)
        assert "不明" in b2.open_folder("nope")["error"]
        # 存在しないフォルダはエラー（開く処理自体はOS依存なので実行しない）
        assert "まだありません" in b2.open_folder("report")["error"]


class TestResultsReadmeAndProgress:
    def test_readme_written_after_recognition(self, tmp_path):
        scans, coord, key = _build_marked_inputs(tmp_path)
        b, _ = _bridge_through_recognition(scans, coord, key)
        from constants import RESULTS_FOLDER
        readme = scans / RESULTS_FOLDER / "README.txt"
        assert readme.exists()
        text = readme.read_text(encoding="utf-8")
        assert "00_Processing" in text and "再開" in text

    def test_progress_derivation(self, tmp_path):
        scans, coord, key = _build_marked_inputs(tmp_path)
        b, adapter = _bridge_through_recognition(scans, coord, key)
        p = b.get_progress()["progress"]
        assert p == {"prepared": True, "read": True,
                     "scored": False, "summarized": False}
        assert b.run_scoring()["ok"]
        _wait_job_done(b)
        p = b.get_progress()["progress"]
        assert p["scored"] is True and p["summarized"] is False
        assert b.run_summary()["ok"]
        _wait_job_done(b)
        assert b.get_progress()["progress"]["summarized"] is True

    def test_progress_empty_bridge(self):
        b = Bridge(window_adapter=RecordingAdapter())
        p = b.get_progress()["progress"]
        assert p == {"prepared": False, "read": False,
                     "scored": False, "summarized": False}
