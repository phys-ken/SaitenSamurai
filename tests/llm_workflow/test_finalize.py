"""個票・集計を作る前の確認と作成（S6）"""
import pytest

from saiten_llm.cli import main
from saiten_llm.cmd_finalize import gate
from saiten_llm.engine import open_bridge
from saiten_llm.workspace import Project
from tests.llm_workflow.test_points import _confirm_all, _points_file
from tests.llm_workflow.test_points import judged_project  # noqa: F401  （フィクスチャ）


@pytest.fixture
def ready_project(judged_project, tmp_path):
    """配点まで確定し、部分点も入れた（個票を作ってよい）試験"""
    root = str(judged_project)
    _confirm_all(root)
    assert main(["set-points", root, "--input", _points_file(tmp_path), "--teacher-approved"]) == 0
    bridge, _ = open_bridge(Project(root))
    assert bridge.set_descriptive_score("小テスト_p004.png", "D1", 6)["ok"]
    assert main(["confirm", root, "question", "--question", "D1"]) == 0
    return judged_project


def _problems(root):
    bridge, _ = open_bridge(Project(root), write=False)
    return gate(Project(root), bridge)


def test_gate_lists_everything_missing(judged_project):
    problems = _problems(judged_project)
    text = "\n".join(problems)
    assert "配点が決まっていません" in text
    assert "正答を先生が確定していません" in text
    assert "マークの確認" in text
    assert main(["finalize", str(judged_project)]) == 2
    from constants import RESULTS_FOLDER, SCORED_FOLDER
    assert not (Project(judged_project).scans / RESULTS_FOLDER / SCORED_FOLDER).exists()


def test_partials_block_until_filled(judged_project, tmp_path):
    root = str(judged_project)
    _confirm_all(root)
    assert main(["set-points", root, "--input", _points_file(tmp_path), "--teacher-approved"]) == 0
    assert any("未採点が 1 人分" in p for p in _problems(root))


def test_ready_project_generates_reports(ready_project):
    from constants import RESULTS_FOLDER, SCORED_FOLDER, FINAL_REPORT_FOLDER
    assert _problems(ready_project) == []
    assert main(["finalize", str(ready_project)]) == 0
    results = Project(ready_project).scans / RESULTS_FOLDER
    assert len(list((results / SCORED_FOLDER).glob("*.png"))) == 6
    assert (results / FINAL_REPORT_FOLDER / "001_student_summary.xlsx").exists()
    assert Project(ready_project).data["finalized"]["at"]


def test_score_changed_after_confirm_is_caught(ready_project):
    bridge, _ = open_bridge(Project(ready_project))
    assert bridge.set_descriptive_score("小テスト_p004.png", "D1", 7)["ok"]
    assert any("得点が書き換わりました" in p for p in _problems(ready_project))
    assert main(["finalize", str(ready_project)]) == 2


def test_answer_key_changed_after_confirm_is_caught(ready_project):
    import openpyxl
    bridge, _ = open_bridge(Project(ready_project), write=False)
    key = bridge.state["answer_key"]
    wb = openpyxl.load_workbook(key)
    wb.active.cell(2, 2).value = "9"
    wb.save(key)
    assert any("answer_key.xlsx が書き換わりました" in p for p in _problems(ready_project))


def test_new_recognition_after_marks_check_is_caught(ready_project):
    bridge, _ = open_bridge(Project(ready_project), write=False)
    omr = bridge.state["omr_result"]
    assert _problems(ready_project) == []
    with open(omr, "ab") as f:
        f.write(b"\0")            # 先生がマーク確認の訂正を反映した等
    assert any("読み取り結果が変わりました" in p for p in _problems(ready_project))
