"""モード・座標ファイル・読み取り・正答（S3）"""
import json
import os

import pytest

from saiten_llm.cli import main
from saiten_llm.engine import open_bridge, session_path
from saiten_llm.mask import mask_status
from saiten_llm.workspace import Project
from tests.llm_workflow.conftest import COORD


def test_mode_needs_teacher_approval(oriented_project):
    assert main(["mode", str(oriented_project), "--set", "mark_only"]) == 2
    assert "mode" not in Project(oriented_project).data
    assert not session_path(Project(oriented_project)).exists()
    assert main(["mode", str(oriented_project), "--set", "mark_only",
                 "--teacher-approved"]) == 0
    data = json.loads(session_path(Project(oriented_project)).read_text(encoding="utf-8"))
    assert data["app_mode"] == "mark_only"


def test_coord_needs_mode_first(oriented_project):
    assert main(["coord", str(oriented_project), str(COORD)]) == 1


def test_recognition_matches_truth(recognized_project, sample_pdf):
    """合成した全員分のマークが、そのまま読み取れている（0 と 10 は同じ列）"""
    from scoring_engine import load_mark2_results
    bridge, _ = open_bridge(Project(recognized_project), write=False)
    students = {s["image"]: s for s in load_mark2_results(
        bridge.state["omr_result"], skip_questions=4)}
    truth = sample_pdf[1]["students"]
    assert len(students) == len(truth)
    checked = 0
    for t in truth:
        got = students[f"小テスト_p{t['page']:03d}.png"]["answers"]
        for q, v in t["marks"].items():
            g = got[int(q)]
            assert (10 if str(g) == "0" else int(g)) == v, (t["page"], q, g, v)
            checked += 1
    assert checked == 45 * len(truth)
    rec = Project(recognized_project).data["recognition"]
    assert rec["single_mark_ratio"] == 1.0 and rec["warnings"] == {}


def test_coord_overlay_and_copies(recognized_project, tmp_path):
    project = Project(recognized_project)
    assert (project.llm_view / "coord_check.jpg").exists()
    # 座標ファイルは採点結果フォルダの中に写し、採点侍の設定は相対パスで持つ
    data = json.loads(session_path(project).read_text(encoding="utf-8"))
    assert not os.path.isabs(data["coord_excel"])
    assert project.data["coord"]["answer_rows"] == 45


def test_coord_is_kept_in_templates(oriented_project):
    templates = oriented_project.parent / "templates"
    templates.mkdir()
    assert main(["mode", str(oriented_project), "--set", "mark_only",
                 "--teacher-approved"]) == 0
    assert main(["coord", str(oriented_project), str(COORD),
                 "--templates", str(templates)]) == 0
    assert (templates / COORD.name).exists()


def test_wrong_coord_file_is_noticed(recognized_project, tmp_path):
    """用紙と合わない座標ファイル（位置を大きくずらしたもの）で読むと、警告して止まる"""
    import openpyxl
    wrong = tmp_path / "別の用紙の座標.xlsx"
    wb = openpyxl.load_workbook(COORD)
    ws = wb.active
    for row in ws.iter_rows(min_row=4):
        for cell in row[2:]:
            if isinstance(cell.value, (int, float)):
                cell.value = cell.value + 37
    wb.save(wrong)
    assert main(["coord", str(recognized_project), str(wrong)]) == 0
    assert main(["recognize", str(recognized_project)]) == 2
    assert Project(recognized_project).data["recognition"]["single_mark_ratio"] < 0.6


def _key_rows(bridge):
    from scoring_engine import load_template
    t = load_template(bridge.state["answer_key"], mark_format="standard")
    return {q: (str(v["正答"]), v["配点"]) for q, v in t.items()}


def test_answer_key_from_json_uses_one_point(recognized_project, tmp_path):
    key = tmp_path / "key.json"
    key.write_text(json.dumps({"rows": [{"問題番号": i, "正答": str((i - 1) % 10 + 1)}
                                        for i in range(1, 46)]}), encoding="utf-8")
    assert main(["answer-key", str(recognized_project), "--from-json", str(key)]) == 0
    bridge, _ = open_bridge(Project(recognized_project), write=False)
    rows = _key_rows(bridge)
    assert len(rows) == 45 and all(p == 1 for _a, p in rows.values())
    assert rows[10][0] == "10"
    assert bridge.state["key_summary"]["ok"]


def test_answer_key_from_teacher_xlsx_is_not_rewritten(recognized_project):
    import hashlib
    src = COORD.parent / "answer_key_sample.xlsx"
    assert main(["answer-key", str(recognized_project), "--from-xlsx", str(src)]) == 0
    bridge, _ = open_bridge(Project(recognized_project), write=False)
    from pathlib import Path
    digest = lambda p: hashlib.sha256(Path(p).read_bytes()).hexdigest()  # noqa: E731
    assert digest(bridge.state["answer_key"]) == digest(src)


def test_answer_key_from_marked_sheet_then_set_aside(recognized_project, sample_pdf):
    """正答をマークした用紙から正答を取り、その用紙を集計から外す"""
    sheet = "小テスト_p002.png"
    assert main(["answer-key", str(recognized_project), "--from-sheet", sheet]) == 0
    bridge, _ = open_bridge(Project(recognized_project), write=False)
    rows = _key_rows(bridge)
    marks = sample_pdf[1]["students"][1]["marks"]
    for q, v in marks.items():
        a = rows[int(q)][0]
        assert (10 if a == "0" else int(a)) == v

    assert main(["set-aside", str(recognized_project), sheet, "--reason", "正答用"]) == 0
    project = Project(recognized_project)
    assert sheet not in [p.name for p in project.scan_files()]
    assert (project.root / "除外" / sheet).exists()
    # 取り除いただけなので黒塗りの承認は生きている
    assert mask_status(project) == "approved"
    assert main(["recognize", str(recognized_project)]) == 0
    bridge, _ = open_bridge(project, write=False)
    from scoring_engine import load_mark2_results
    names = [s["image"] for s in load_mark2_results(bridge.state["omr_result"], skip_questions=4)]
    assert sheet not in names and len(names) == 5


def test_writes_are_refused_while_screen_is_open(recognized_project):
    from api import uilock
    project = Project(recognized_project)
    uilock.acquire(project.scans)          # このテスト自身のプロセスが「画面」役
    try:
        assert main(["recognize", str(recognized_project)]) == 2
        assert main(["set-aside", str(recognized_project), "小テスト_p001.png"]) == 2
        assert (project.scans / "小テスト_p001.png").exists()
    finally:
        uilock.release(project.scans)
    assert main(["status", str(recognized_project)]) == 0


def test_stale_lock_from_crashed_screen_is_ignored(recognized_project):
    from api import uilock
    project = Project(recognized_project)
    path = uilock.lock_path(project.scans)
    path.write_text(json.dumps({"pid": 999999, "opened_at": "x"}), encoding="utf-8")
    assert uilock.holder(project.scans) is None


def test_facts_reports_markers_and_bubbles(recognized_project, capsys):
    assert main(["facts", str(recognized_project)]) == 0
    out = capsys.readouterr().out
    assert "四隅のマーカー: 6 / 6 枚" in out
    assert "マーク採点＋記述採点" in out


@pytest.mark.parametrize("mode", ["descriptive_only"])
def test_answer_key_refused_in_descriptive_only(oriented_project, mode, tmp_path):
    assert main(["mode", str(oriented_project), "--set", mode, "--teacher-approved"]) == 0
    key = tmp_path / "k.json"
    key.write_text('{"rows": [{"問題番号": 1, "正答": "1"}]}', encoding="utf-8")
    assert main(["answer-key", str(oriented_project), "--from-json", str(key)]) == 1
