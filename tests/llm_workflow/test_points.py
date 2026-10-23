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


def test_keep_teacher_mark_points(judged_project, tmp_path):
    """先生が answer_key.xlsx に入れた配点をそのまま使う（"marks": "keep"）"""
    root = str(judged_project)
    _confirm_all(root)
    before = Project(root).data["answer_key"]["hash"]
    pts = _write(tmp_path, "keep.json", {"marks": "keep", "descriptive": {"D1": 10}})
    assert main(["set-points", root, "--input", pts, "--teacher-approved"]) == 0
    data = Project(root).data
    assert data["answer_key"]["points_set_at"] and data["answer_key"]["hash"] == before
    assert data["points"]["marks"] == "keep"


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


# --- 部分の点の指定 ------------------------------------------------------

PARTIAL_FILE = "小テスト_p004.png"     # 合成答案で「部分」の生徒


def test_partial_points_rules():
    from saiten_llm.cmd_points import partial_points
    assert partial_points({"D1": 10, "D2": 2}, {"D1": 3}, None) == {"D1": 3}     # D2 は指定なし
    assert partial_points({"D1": 10, "D2": 2}, {}, 0.5) == {"D1": 5, "D2": 1}
    assert partial_points({"D1": 10, "D2": 3}, {"D1": 7}, 0.5) == {"D1": 7, "D2": 2}   # 1.5 → 2
    assert partial_points({"D1": 10}, {}, 0.25) == {"D1": 3}                        # 2.5 → 3（偶数丸めにしない）
    for desc, partial, ratio in (({"D1": 2}, {"D1": 2}, None),       # 満点と同じ
                                 ({"D1": 2}, {"D1": 0}, None),       # 0 点
                                 ({"D1": 10}, {"D1": 2.5}, None),    # 整数でない
                                 ({"D1": 10}, {"D9": 1}, None),      # descriptive に無い
                                 ({"D1": 10}, {}, 0.01),             # 0 点に丸まる
                                 ({"D1": 2}, {}, 0.9)):              # 満点に丸まる
        with pytest.raises(ValueError):
            partial_points(desc, partial, ratio)


def _gate_desc_problems(root):
    from saiten_llm.cmd_finalize import gate
    project = Project(root)
    bridge, _ = open_bridge(project, write=False)
    return [p for p in gate(project, bridge) if p.startswith("記述")]


def _desc_points(tmp_path, desc_max, partial=None, bom=False):
    spec = {"marks": "keep", "descriptive": {"D1": desc_max}}
    if partial is not None:
        spec["partial"] = {"D1": partial}
    p = tmp_path / "points.json"
    text = json.dumps(spec, ensure_ascii=False)
    p.write_bytes((b"\xef\xbb\xbf" if bom else b"") + text.encode("utf-8"))
    return str(p)


def test_set_points_with_partial_keeps_confirmation(judged_project, tmp_path):
    root = str(judged_project)
    _confirm_all(root)
    # Windows の PowerShell が書く BOM 付きの JSON でも読める
    assert main(["set-points", root, "--input", _desc_points(tmp_path, 10, 4, bom=True),
                 "--teacher-approved"]) == 0
    project = Project(root)
    bridge, _ = open_bridge(project, write=False)
    scores = {f: s.get("D1") for f, s in bridge._desc_scores["scores"].items()}
    assert scores[PARTIAL_FILE] == 4
    assert sorted(set(scores.values())) == [0, 4, 10]
    record = jmod.load(project)["questions"]["D1"]
    assert record["converted"]["partials"] == [PARTIAL_FILE]
    assert record["converted"]["partial_score"] == 4
    assert record["confirmed"]["max_score"] == 10
    assert record["confirmed"]["by"] == "確定済みの正誤・部分を換算"
    assert record["provisional_confirmed"]["max_score"] == 2
    assert project.data["points"]["partial"] == {"D1": 4}
    assert _gate_desc_problems(root) == []
    # 部分の点は stats の正答率にもそのまま入る（4/10）
    from saiten_llm.cmd_points import build_matrix
    items, students, matrix = build_matrix(bridge)
    col = [it["id"] for it in items].index("D1")
    assert matrix[students.index(PARTIAL_FILE)][col] == pytest.approx(0.4)


def test_set_points_keeping_two_points(judged_project, tmp_path):
    """満点 2 点のまま（正 2・部分 1・誤 0）: 点数は変わらず、確定だけが保たれる"""
    root = str(judged_project)
    _confirm_all(root)
    project = Project(root)
    bridge, _ = open_bridge(project, write=False)
    before = {f: s.get("D1") for f, s in bridge._desc_scores["scores"].items()}
    first_confirm = jmod.load(project)["questions"]["D1"]["confirmed"]
    assert before[PARTIAL_FILE] == 1
    for _ in range(2):                    # 2回実行しても同じ
        assert main(["set-points", root, "--input", _desc_points(tmp_path, 2, 1),
                     "--teacher-approved"]) == 0
        bridge, _ = open_bridge(project, write=False)
        assert {f: s.get("D1") for f, s in bridge._desc_scores["scores"].items()} == before
        record = jmod.load(project)["questions"]["D1"]
        assert record["confirmed"]["hash"] == first_confirm["hash"]
        assert record["provisional_confirmed"] == first_confirm      # 最初の確定は上書きしない
        assert _gate_desc_problems(root) == []


def test_set_points_partial_ratio(judged_project, tmp_path):
    root = str(judged_project)
    _confirm_all(root)
    # 0 点に丸まる割合は断り、何も書き換えない
    assert main(["set-points", root, "--input", _desc_points(tmp_path, 10), "--partial-ratio", "0.01",
                 "--teacher-approved"]) == 1
    bridge, _ = open_bridge(Project(root), write=False)
    assert bridge.state["descriptive"]["questions"][0]["max_score"] == 2
    assert main(["set-points", root, "--input", _desc_points(tmp_path, 10), "--partial-ratio", "0.5",
                 "--teacher-approved"]) == 0
    bridge, _ = open_bridge(Project(root), write=False)
    assert bridge._desc_scores["scores"][PARTIAL_FILE]["D1"] == 5
    assert _gate_desc_problems(root) == []


def test_set_points_does_not_carry_stale_confirmation(judged_project, tmp_path):
    """確定の後で画面の得点が変わっていたら、確定は引き継がない"""
    root = str(judged_project)
    _confirm_all(root)
    bridge, _ = open_bridge(Project(root))
    assert bridge.set_descriptive_score("小テスト_p001.png", "D1", 0)["ok"]   # 確定の後の変更
    assert main(["set-points", root, "--input", _desc_points(tmp_path, 10, 5),
                 "--teacher-approved"]) == 0
    assert jmod.load(Project(root))["questions"]["D1"]["confirmed"] is None
    assert any("先生の確定がありません" in p for p in _gate_desc_problems(root))
    assert main(["confirm", root, "question", "--question", "D1"]) == 0
    assert _gate_desc_problems(root) == []


def test_set_points_without_partial_still_waits(judged_project, tmp_path):
    """partial を書かなければ、これまでどおり部分は未採点に戻り、finalize の前で止まる"""
    root = str(judged_project)
    _confirm_all(root)
    assert main(["set-points", root, "--input", _desc_points(tmp_path, 10),
                 "--teacher-approved"]) == 0
    problems = _gate_desc_problems(root)
    assert len(problems) == 1 and "未採点が 1 人分" in problems[0]
    assert main(["finalize", root]) == 2
