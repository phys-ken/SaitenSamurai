"""標準モード × 実スキャン画像の通し（webui 既定設定のまま満点になること）。

tk 版 e2e (test_scoring_e2e) と同じ実サンプル
（sample_marksheet.jpg / M2-03-002 / answer_key_sample = 90点満点）を、
webui Bridge の既定設定（kmeans・skip=4）で認識→採点し、
得点まで一致することを確認する。
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

SAMPLE = PROJECT_ROOT / "sample_basefile"
IMAGE = SAMPLE / "sample_marksheet.jpg"
COORD = SAMPLE / "M2-03-002_座標ファイル.xlsx"
KEY = SAMPLE / "answer_key_sample.xlsx"


@pytest.fixture(scope="module")
def piped(tmp_path_factory):
    if not (IMAGE.exists() and COORD.exists() and KEY.exists()):
        pytest.skip("サンプルファイルがありません")
    import shutil
    scans = tmp_path_factory.mktemp("real") / "scans"
    scans.mkdir()
    shutil.copy(IMAGE, scans / IMAGE.name)

    adapter = RecordingAdapter()
    b = Bridge(window_adapter=adapter)          # 既定: standard / skip=4 / kmeans
    adapter.folder_returns = [str(scans)]
    adapter.file_returns = [str(COORD), str(KEY)]
    assert b.select_image_folder()["ok"]
    assert b.select_coord_file()["ok"]
    assert b.select_answer_key()["ok"]
    assert b.state["key_summary"]["ok"] is True

    assert b.run_recognition()["ok"]
    _wait_job_done(b, timeout=120)
    done = [e for e in adapter.events if e["type"] == "job_done"][-1]
    assert done["ok"] and done["success_count"] == 1
    return b


def test_coord_summary_matches_template(piped):
    s = piped.state["coord_summary"]
    assert s["marks_per_row"] == 10 and s["warning"] is None


def test_full_score_via_webui_defaults(piped):
    """実画像を webui 既定設定で認識した結果が tk e2e と同じ90点満点"""
    from scoring_engine import load_template, load_mark2_results, score_answers
    td = load_template(str(KEY))
    students = load_mark2_results(piped.state["omr_result"], skip_questions=4)
    assert len(students) == 1
    result = score_answers(students[0]["answers"], td)
    assert result["max_score"] == 90
    assert result["total_score"] == 90, result["results"]


def test_scoring_job_outputs_image(piped):
    from constants import RESULTS_FOLDER, SCORED_FOLDER
    b = piped
    assert b.run_scoring()["ok"]
    _wait_job_done(b, timeout=120)
    scored = Path(b.state["image_folder"]) / RESULTS_FOLDER / SCORED_FOLDER
    assert (scored / IMAGE.name).exists()
