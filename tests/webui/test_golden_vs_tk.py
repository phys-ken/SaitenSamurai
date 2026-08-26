"""L4 ゴールデン: 同一入力で webui 経由と tk 版ドライバ直呼びの出力が一致する。

「現在のプロジェクトと同じ入力（答案と模範解答と配点など）であれば、
同じ出力になるべき」という要件そのものを固定するテスト。
- A系: tk ドライバを直接呼ぶ（tk GUI と同じ関数・同じ既定引数）
- B系: webui Bridge 経由（select→認識→採点→集計）
比較: 採点済み答案画像のピクセル一致 / 成績一覧 Excel のセル値一致。
"""
import shutil
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

SKIP = 0


def _build_inputs(base: Path):
    """合成入力一式（答案3枚: 完答 / 1行誤り / 無マーク行あり）"""
    import cv2
    coord = make_coord_xlsx(base / "coord.xlsx", 4)
    key = base / "key.xlsx"
    _create_answer_key(key, [
        {'問題番号': 1, '正答': '-24', '配点': 3, '観点': 1},
        {'問題番号': 4, '正答': 'a', '配点': 2, '観点': 2},
    ])
    scans = base / "scans"
    scans.mkdir()
    sheets = {
        "s1.png": {**{q: p for q, p in zip([1, 2, 3], sym_positions('-24'))},
                   4: sym_positions('a')[0]},                       # 5点
        "s2.png": {1: sym_positions('-')[0], 2: sym_positions('9')[0],
                   3: sym_positions('4')[0], 4: sym_positions('a')[0]},  # 2点
        "s3.png": {2: sym_positions('2')[0], 3: sym_positions('4')[0]},  # 0点
    }
    for name, filled in sheets.items():
        cv2.imwrite(str(scans / name), make_sheet(filled, with_markers=True))
    return scans, coord, key


def _run_tk_pipeline(scans, coord, key):
    """tk 版と同じドライバ直呼び（main_gui が渡すのと同じ引数）"""
    from omr_engine import process_box_drawer
    from image_renderer import process_scoring
    from summary_generator import process_summary_generation
    from constants import (RESULTS_FOLDER, RESULTS_DATA_FOLDER,
                           READING_RESULTS_FOLDER_NAME)
    process_box_drawer(str(scans), str(coord), skip_questions=SKIP,
                       mark_format="multi_digit")
    omr = sorted((scans / RESULTS_FOLDER / RESULTS_DATA_FOLDER /
                  READING_RESULTS_FOLDER_NAME).glob("Mark2-Result-*.xlsx"))[-1]
    process_scoring(str(scans), str(coord), str(key), str(omr),
                    skip_questions=SKIP, mark_format="multi_digit")
    process_summary_generation(str(scans), str(coord), str(key), str(omr),
                               skip_questions=SKIP, mark_format="multi_digit")
    return scans / RESULTS_FOLDER


def _run_webui_pipeline(scans, coord, key):
    adapter = RecordingAdapter()
    b = Bridge(window_adapter=adapter)
    b.set_mode("mark_only", "multi_digit")
    b.set_skip_questions(SKIP)
    adapter.folder_returns = [str(scans)]
    adapter.file_returns = [str(coord), str(key)]
    assert b.select_image_folder()["ok"]
    assert b.select_coord_file()["ok"]
    assert b.select_answer_key()["ok"]
    for job in ("run_recognition", "run_scoring", "run_summary"):
        assert getattr(b, job)()["ok"], job
        _wait_job_done(b)
        done = [e for e in adapter.events if e["type"] == "job_done"][-1]
        assert done["ok"], done
    from constants import RESULTS_FOLDER
    return scans / RESULTS_FOLDER


@pytest.fixture(scope="module")
def golden(tmp_path_factory):
    base_a = tmp_path_factory.mktemp("tk")
    base_b = tmp_path_factory.mktemp("webui")
    res_a = _run_tk_pipeline(*_build_inputs(base_a))
    res_b = _run_webui_pipeline(*_build_inputs(base_b))
    return res_a, res_b


class TestGoldenOutputs:
    def test_scored_images_pixel_identical(self, golden):
        import cv2
        import numpy as np
        from constants import SCORED_FOLDER
        res_a, res_b = golden
        a_files = sorted((res_a / SCORED_FOLDER).glob("*.png"))
        b_files = sorted((res_b / SCORED_FOLDER).glob("*.png"))
        assert [f.name for f in a_files] == [f.name for f in b_files] \
            == ["s1.png", "s2.png", "s3.png"]
        for fa, fb in zip(a_files, b_files):
            ia, ib = cv2.imread(str(fa)), cv2.imread(str(fb))
            assert ia.shape == ib.shape, fa.name
            assert np.array_equal(ia, ib), f"{fa.name} のピクセルが tk 版と不一致"

    def test_student_summary_values_identical(self, golden):
        import openpyxl
        from constants import FINAL_REPORT_FOLDER, STUDENT_SUMMARY_FILE
        res_a, res_b = golden

        def cells(res):
            p = res / FINAL_REPORT_FOLDER / STUDENT_SUMMARY_FILE
            assert p.exists(), p
            wb = openpyxl.load_workbook(p, data_only=True)
            out = {}
            for ws in wb.worksheets:
                out[ws.title] = [[c.value for c in row] for row in ws.iter_rows()]
            return out

        assert cells(res_a) == cells(res_b)

    def test_exam_summary_and_ctt_exist_in_both(self, golden):
        from constants import (FINAL_REPORT_FOLDER, EXAM_SUMMARY_FILE,
                               CTT_ANALYSIS_EXCEL_FILE)
        res_a, res_b = golden
        for res in (res_a, res_b):
            assert (res / FINAL_REPORT_FOLDER / EXAM_SUMMARY_FILE).exists()
            assert (res / FINAL_REPORT_FOLDER / CTT_ANALYSIS_EXCEL_FILE).exists()
