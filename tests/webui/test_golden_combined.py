"""L4 ゴールデン（マーク＋記述の複合モード）: webui と tk 版ドライバの出力一致。

マーク採点（複数桁）＋記述2問という実運用に近い構成で、
採点済み答案画像のピクセル一致と成績一覧 Excel のセル値一致を固定する。
"""
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

# 記述2問（座標は 00_Processing 基準のピクセル座標）
DESC_QUESTIONS = [
    {"id": "D1", "name": "記述1", "max_score": 5, "aspect": 1,
     "region": [40, 600, 400, 700]},
    {"id": "D2", "name": "記述2", "max_score": 3, "aspect": 2,
     "region": [40, 720, 400, 800]},
]
DESC_SCORES = {
    "s1.png": {"D1": 5, "D2": 2},
    "s2.png": {"D1": 3},           # D2 は未採点のまま
}


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
    sheets = {
        "s1.png": {**{q: p for q, p in zip([1, 2, 3], sym_positions('-24'))},
                   4: sym_positions('a')[0]},                       # マーク5点
        "s2.png": {1: sym_positions('-')[0], 2: sym_positions('9')[0],
                   3: sym_positions('4')[0], 4: sym_positions('a')[0]},  # 2点
    }
    for name, filled in sheets.items():
        cv2.imwrite(str(scans / name), make_sheet(filled, with_markers=True))
    return scans, coord, key


def _run_tk_pipeline(scans, coord, key):
    """tk 版と同じドライバ直呼び（main_gui.py:2529 / :3670 と同じ引数）"""
    from omr_engine import process_box_drawer
    from descriptive_scorer import generate_return_sheets
    from summary_generator import process_summary_generation
    from constants import (RESULTS_FOLDER, RESULTS_DATA_FOLDER, SCORED_FOLDER,
                           READING_RESULTS_FOLDER_NAME)
    process_box_drawer(str(scans), str(coord), skip_questions=SKIP,
                       mark_format="multi_digit")
    omr = sorted((scans / RESULTS_FOLDER / RESULTS_DATA_FOLDER /
                  READING_RESULTS_FOLDER_NAME).glob("Mark2-Result-*.xlsx"))[-1]
    out = scans / RESULTS_FOLDER / SCORED_FOLDER
    generate_return_sheets(
        image_folder=str(scans),
        config={"questions": DESC_QUESTIONS},
        descriptive_scores=DESC_SCORES,
        coord_excel_path=str(coord),
        template_path=str(key),
        mark2_result_path=str(omr),
        skip_questions=SKIP,
        output_folder=str(out),
        mark_format="multi_digit",
    )
    process_summary_generation(
        str(scans), str(coord), str(key), str(omr),
        skip_questions=SKIP,
        descriptive_config={"questions": DESC_QUESTIONS},
        descriptive_scores=DESC_SCORES,
        include_descriptive_in_analysis=True,
        mark_format="multi_digit",
    )
    return scans / RESULTS_FOLDER


def _run_webui_pipeline(scans, coord, key):
    adapter = RecordingAdapter()
    b = Bridge(window_adapter=adapter)
    b.set_mode("mark_and_descriptive", "multi_digit")
    b.set_skip_questions(SKIP)
    adapter.folder_returns = [str(scans)]
    adapter.file_returns = [str(coord), str(key)]
    assert b.select_image_folder()["ok"]
    assert b.select_coord_file()["ok"]
    assert b.select_answer_key()["ok"]
    assert b.run_recognition()["ok"]
    _wait_job_done(b)
    # UI と同じ経路で記述問題を登録し、得点を付ける
    for q in DESC_QUESTIONS:
        res = b.add_descriptive_question(q["name"], q["max_score"],
                                         q["aspect"], q["region"])
        assert res["ok"] and res["question_id"] == q["id"], res
    for fname, scores in DESC_SCORES.items():
        for qid, score in scores.items():
            assert b.set_descriptive_score(fname, qid, score)["ok"]
    for job in ("run_scoring", "run_summary"):
        assert getattr(b, job)()["ok"], job
        _wait_job_done(b)
        done = [e for e in adapter.events if e["type"] == "job_done"][-1]
        assert done["ok"], done
    from constants import RESULTS_FOLDER
    return scans / RESULTS_FOLDER


@pytest.fixture(scope="module")
def golden(tmp_path_factory):
    res_a = _run_tk_pipeline(*_build_inputs(tmp_path_factory.mktemp("tk")))
    res_b = _run_webui_pipeline(*_build_inputs(tmp_path_factory.mktemp("webui")))
    return res_a, res_b


class TestCombinedGolden:
    def test_scored_images_pixel_identical(self, golden):
        import cv2
        import numpy as np
        from constants import SCORED_FOLDER
        res_a, res_b = golden
        a_files = sorted((res_a / SCORED_FOLDER).glob("*.png"))
        b_files = sorted((res_b / SCORED_FOLDER).glob("*.png"))
        assert [f.name for f in a_files] == [f.name for f in b_files] \
            == ["s1.png", "s2.png"]
        for fa, fb in zip(a_files, b_files):
            ia, ib = cv2.imread(str(fa)), cv2.imread(str(fb))
            assert ia.shape == ib.shape, fa.name
            assert np.array_equal(ia, ib), f"{fa.name} のピクセルが tk 版と不一致"

    def test_scored_images_show_descriptive_marks(self, golden):
        """記述領域に何かが描き込まれている（素通しではない）ことを確認"""
        import cv2
        import numpy as np
        from constants import SCORED_FOLDER, BOXED_FOLDER
        _, res_b = golden
        scored = cv2.imread(str(res_b / SCORED_FOLDER / "s1.png"))
        base = cv2.imread(str(res_b / BOXED_FOLDER / "s1.png"))
        assert scored is not None and base is not None
        x1, y1, x2, y2 = DESC_QUESTIONS[0]["region"]
        region_scored = scored[y1:y2, x1:x2]
        region_base = base[y1:y2, x1:x2]
        assert not np.array_equal(region_scored, region_base), \
            "記述領域に得点の描画がない"

    def test_student_summary_identical_and_includes_descriptive(self, golden):
        import openpyxl
        from constants import FINAL_REPORT_FOLDER, STUDENT_SUMMARY_FILE
        res_a, res_b = golden

        def cells(res):
            p = res / FINAL_REPORT_FOLDER / STUDENT_SUMMARY_FILE
            assert p.exists(), p
            wb = openpyxl.load_workbook(p, data_only=True)
            return {ws.title: [[c.value for c in row] for row in ws.iter_rows()]
                    for ws in wb.worksheets}

        ca, cb = cells(res_a), cells(res_b)
        assert ca == cb
        flat = [str(v) for rows in cb.values() for row in rows for v in row]
        assert any("D1" in v or "記述1" in v for v in flat), \
            "成績一覧に記述問題の列が見当たらない"
