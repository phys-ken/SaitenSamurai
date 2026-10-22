"""配点の統計・配点案・換算（S5）"""
import json

import numpy as np
import pytest

from saiten_llm import allocate
from saiten_llm import judgments as jmod
from saiten_llm.cli import main
from saiten_llm.engine import open_bridge
from saiten_llm.workspace import Project
from tests.llm_workflow.conftest import COORD
from tests.llm_workflow.test_desc_flow import _index, _truth_verdicts, _write


# --- 計算 ---------------------------------------------------------------

def test_integer_points_sum_and_minimum():
    pts = allocate.integer_points([1, 1, 1], 100)
    assert sum(pts) == 100 and min(pts) >= 1 and max(pts) - min(pts) <= 1
    pts = allocate.integer_points([3, 1], 10)
    assert sum(pts) == 10 and pts[0] > pts[1]
    with pytest.raises(ValueError):
        allocate.integer_points([1] * 5, 4)


def test_item_stats_signs():
    rng = np.random.default_rng(0)
    ability = rng.normal(size=200)
    good = (ability > 0).astype(float)            # 力のある生徒ほど正答
    bad = (ability < 0).astype(float)             # 力のある生徒ほど誤答
    noise = [(ability + rng.normal(size=200) > t).astype(float) for t in (-0.5, 0, 0.5)]
    # 良い項目だけの表では識別力がはっきり正
    stats = allocate.item_stats(np.column_stack([good] + noise))
    assert stats[0]["r"] > 0.3 and stats[0]["ul"] > 0.5
    assert abs(stats[0]["p"] - good.mean()) < 1e-9
    # 逆向きの項目を混ぜると、その項目の識別力は負
    stats = allocate.item_stats(np.column_stack([bad] + noise))
    assert stats[0]["r"] < 0 and stats[0]["ul"] < 0


def test_alpha_of_identical_items_is_one():
    col = np.array([0, 1, 1, 0, 1, 0, 1, 1], dtype=float)
    assert allocate.cronbach_alpha(np.column_stack([col, col, col])) == pytest.approx(1.0)


def test_constant_item_has_no_discrimination():
    m = np.array([[1, 0], [1, 1], [1, 0], [1, 1]], dtype=float)
    assert allocate.item_stats(m)[0]["r"] is None


# --- 流れ ---------------------------------------------------------------

@pytest.fixture
def judged_project(recognized_project, sample_pdf, tmp_path):
    """記述を判定し、マーク・正答・記述の確定がそろう直前の試験"""
    x1, y1, x2, y2 = sample_pdf[1]["desc_box_595"]
    regions = _write(tmp_path, "regions.json",
                     {"questions": [{"name": "問46", "region": [x1 * 2, y1 * 2, x2 * 2, y2 * 2]}]})
    root = str(recognized_project)
    assert main(["desc-set", root, "--input", regions]) == 0
    assert main(["crops", root]) == 0
    verdicts = _write(tmp_path, "v.json", _truth_verdicts(sample_pdf, _index(Project(root))))
    assert main(["judge", root, "--question", "D1", "--input", verdicts]) == 0
    key = COORD.parent / "answer_key_sample.xlsx"
    assert main(["answer-key", root, "--from-xlsx", str(key)]) == 0
    return recognized_project


def _confirm_all(root):
    bridge, _ = open_bridge(Project(root))
    assert bridge.set_descriptive_score("小テスト_p003.png", "D1", 2)["ok"]   # 保留に先生が点
    assert main(["confirm", root, "question", "--question", "D1"]) == 0
    assert main(["confirm", root, "answer-key"]) == 0
    assert main(["confirm", root, "marks"]) == 0


def _points_file(tmp_path, desc_max=10):
    return _write(tmp_path, "points.json",
                  {"marks": {str(i): 2 for i in range(1, 46)}, "descriptive": {"D1": desc_max}})


def test_set_points_needs_approval_and_confirmations(judged_project, tmp_path):
    root = str(judged_project)
    pts = _points_file(tmp_path)
    assert main(["set-points", root, "--input", pts]) == 2              # 先生の承認なし
    assert main(["set-points", root, "--input", pts, "--teacher-approved"]) == 2   # 未確定あり
    bridge, _ = open_bridge(Project(root), write=False)
    assert bridge.state["descriptive"]["questions"][0]["max_score"] == 2


def test_stats_writes_plans(judged_project, capsys):
    root = str(judged_project)
    _confirm_all(root)
    capsys.readouterr()
    assert main(["stats", root, "--total", "100"]) == 0
    out = capsys.readouterr().out
    assert "まだ確定していない" not in out
    data = json.loads((Project(root).root / "配点の検討.json").read_text(encoding="utf-8"))
    assert data["students"] == 6 and len(data["items"]) == 46
    assert [p["key"] for p in data["plans"]] == ["equal", "difficulty", "discrimination"]
    assert all(sum(p["points"]) == 100 for p in data["plans"])


def test_set_points_converts_and_waits_for_partials(judged_project, sample_pdf, tmp_path):
    root = str(judged_project)
    _confirm_all(root)
    assert main(["set-points", root, "--input", _points_file(tmp_path), "--teacher-approved"]) == 0
    project = Project(root)
    bridge, _ = open_bridge(project, write=False)
    assert bridge.state["descriptive"]["questions"][0]["max_score"] == 10
    assert bridge.state["key_summary"]["ok"]
    assert "満点90点" in bridge.state["key_summary"]["stats_line"]

    scores = {f: s.get("D1") for f, s in bridge._desc_scores["scores"].items()}
    by_page = {f"小テスト_p{s['page']:03d}.png": s["descriptive"] for s in sample_pdf[1]["students"]}
    for f, d in by_page.items():
        if d["text"] == "こうごうせい" or d["truth"] == "正":
            assert scores[f] == 10, f
        elif d["truth"] == "部分":
            assert scores.get(f) is None, f     # 先生が個別に入れる
        else:
            assert scores[f] == 0, f
    record = jmod.load(project)["questions"]["D1"]
    assert record["confirmed"] is None and record["converted"]["partials"] == ["小テスト_p004.png"]

    # 先生が部分点を入れて確定
    bridge, _ = open_bridge(project)
    assert bridge.set_descriptive_score("小テスト_p004.png", "D1", 6)["ok"]
    assert main(["confirm", root, "question", "--question", "D1"]) == 0
    record = jmod.load(project)["questions"]["D1"]
    assert record["confirmed"]["max_score"] == 10

    # 2回目は換算し直さない（10点を「2点満点の正」と取り違えない）
    assert main(["set-points", root, "--input", _points_file(tmp_path), "--teacher-approved"]) == 0
    bridge, _ = open_bridge(project, write=False)
    assert bridge._desc_scores["scores"]["小テスト_p004.png"]["D1"] == 6
