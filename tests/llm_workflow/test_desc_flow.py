"""記述問題の範囲・並べた画像・判定・確定（S4）"""
import json

import cv2
import pytest

from saiten_llm import judgments as jmod
from saiten_llm.cli import main
from saiten_llm.engine import open_bridge
from saiten_llm.workspace import Project


def _write(tmp_path, name, data):
    p = tmp_path / name
    p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return str(p)


@pytest.fixture
def desc_project(recognized_project, sample_pdf, tmp_path):
    """記述問題（問46）の範囲を設定し、並べた画像まで作った試験"""
    x1, y1, x2, y2 = sample_pdf[1]["desc_box_595"]
    regions = _write(tmp_path, "regions.json",
                     {"questions": [{"name": "問46", "region": [x1 * 2, y1 * 2, x2 * 2, y2 * 2]}]})
    assert main(["desc-set", str(recognized_project), "--input", regions]) == 0
    assert main(["crops", str(recognized_project), "--per-sheet", "4"]) == 0
    return recognized_project


def _index(project):
    data = json.loads((project.llm_view / "crops" / "D1" / "index.json").read_text(encoding="utf-8"))
    return {int(k): v for k, v in data["numbers"].items()}


def _truth_verdicts(sample_pdf, index):
    """合成答案の正解どおりの判定（無答は誤、ひらがなは保留）"""
    by_page = {f"小テスト_p{s['page']:03d}.png": s["descriptive"] for s in sample_pdf[1]["students"]}
    out = {}
    for n, f in index.items():
        d = by_page[f]
        if d["text"] == "こうごうせい":
            out[str(n)] = {"verdict": "保留", "reason": "ひらがな表記の扱いが採点基準にない"}
        elif d["truth"] == "部分":
            out[str(n)] = {"verdict": "部分", "reason": "誤字"}
        elif d["truth"] == "正":
            out[str(n)] = "正"
        else:
            out[str(n)] = "誤"
    return out


def test_regions_are_converted_to_saiten_coordinates(desc_project, sample_pdf):
    bridge, _ = open_bridge(Project(desc_project), write=False)
    q = bridge.state["descriptive"]["questions"][0]
    assert q["max_score"] == jmod.PROVISIONAL_MAX
    assert q["region"] == list(sample_pdf[1]["desc_box_595"])   # 写しの画素 ÷ 2
    assert (Project(desc_project).llm_view / "desc_regions.jpg").exists()


def test_montage_numbers_and_masked_source(desc_project):
    project = Project(desc_project)
    index = _index(project)
    assert sorted(index) == list(range(1, 7))
    sheets = sorted((project.llm_view / "crops" / "D1").glob("sheet_*.jpg"))
    assert [p.name for p in sheets] == ["sheet_01_001-004.jpg", "sheet_02_005-006.jpg"]
    img = cv2.imread(str(sheets[0]))
    assert img is not None and img.shape[1] > img.shape[0]   # 横長の解答欄を並べている


def test_judge_writes_provisional_scores(desc_project, sample_pdf, tmp_path):
    project = Project(desc_project)
    index = _index(project)
    verdicts = _write(tmp_path, "v.json", _truth_verdicts(sample_pdf, index))
    assert main(["judge", str(desc_project), "--question", "D1", "--input", verdicts]) == 0
    bridge, _ = open_bridge(project, write=False)
    scores = {f: s.get("D1") for f, s in bridge._desc_scores["scores"].items()}
    by_page = {f"小テスト_p{s['page']:03d}.png": s["descriptive"] for s in sample_pdf[1]["students"]}
    for f, d in by_page.items():
        expected = (None if d["text"] == "こうごうせい" else
                    {"正": 2, "部分": 1, "誤": 0, "無答": 0}[d["truth"]])
        assert scores.get(f) == expected, (f, d)
    data = jmod.load(project)
    holds = [f for f, i in data["questions"]["D1"]["items"].items() if i["status"] == "hold"]
    assert holds == ["小テスト_p003.png"]


def test_bad_verdicts_are_rejected(desc_project, tmp_path):
    for bad in ({"1": "まる"}, {"99": "正"}, {"1": {"verdict": "保留"}}):
        path = _write(tmp_path, "bad.json", bad)
        assert main(["judge", str(desc_project), "--question", "D1", "--input", path]) == 1
    assert not jmod.path(Project(desc_project)).exists()


def test_confirm_blocked_until_holds_are_scored(desc_project, sample_pdf, tmp_path):
    project = Project(desc_project)
    verdicts = _write(tmp_path, "v.json", _truth_verdicts(sample_pdf, _index(project)))
    assert main(["judge", str(desc_project), "--question", "D1", "--input", verdicts]) == 0
    assert main(["confirm", str(desc_project), "question", "--question", "D1"]) == 2
    assert jmod.load(project)["questions"]["D1"]["confirmed"] is None

    # 先生が画面で保留に点を入れ、LLM の誤を部分に直した
    bridge, _ = open_bridge(project)
    assert bridge.set_descriptive_score("小テスト_p003.png", "D1", 2)["ok"]
    assert bridge.set_descriptive_score("小テスト_p005.png", "D1", 1)["ok"]

    assert main(["confirm", str(desc_project), "question", "--question", "D1"]) == 0
    confirmed = jmod.load(project)["questions"]["D1"]["confirmed"]
    # 保留に点を入れたのは修正ではない。LLM の誤を部分にしたのが1件
    assert confirmed["teacher_changes"] == 1 and confirmed["hash"]

    # 確定した問題は LLM が書き換えない
    assert main(["judge", str(desc_project), "--question", "D1", "--input", verdicts]) == 1


def test_rejudge_keeps_teacher_changes(desc_project, sample_pdf, tmp_path):
    project = Project(desc_project)
    verdicts = _write(tmp_path, "v.json", _truth_verdicts(sample_pdf, _index(project)))
    assert main(["judge", str(desc_project), "--question", "D1", "--input", verdicts]) == 0
    bridge, _ = open_bridge(project)
    assert bridge.set_descriptive_score("小テスト_p005.png", "D1", 1)["ok"]   # 先生の修正
    # 保留だった答案を LLM が判定し直す（全員分を送っても先生の修正は守られる）
    again = _truth_verdicts(sample_pdf, _index(project))
    again["3"] = "正"
    path = _write(tmp_path, "v2.json", again)
    assert main(["judge", str(desc_project), "--question", "D1", "--input", path]) == 0
    bridge, _ = open_bridge(project, write=False)
    scores = {f: s.get("D1") for f, s in bridge._desc_scores["scores"].items()}
    assert scores["小テスト_p005.png"] == 1
    assert scores["小テスト_p003.png"] == 2


def test_review_lists_holds_and_changes(desc_project, sample_pdf, tmp_path, capsys):
    project = Project(desc_project)
    verdicts = _write(tmp_path, "v.json", _truth_verdicts(sample_pdf, _index(project)))
    assert main(["judge", str(desc_project), "--question", "D1", "--input", verdicts]) == 0
    bridge, _ = open_bridge(project)
    assert bridge.set_descriptive_score("小テスト_p001.png", "D1", 0)["ok"]
    capsys.readouterr()
    assert main(["review", str(desc_project)]) == 0
    out = capsys.readouterr().out
    assert "保留: 小テスト_p003.png" in out
    assert "先生の修正: 小テスト_p001.png  LLM 正 → 誤" in out


def test_accuracy_counts_teacher_corrections(desc_project, sample_pdf, tmp_path, capsys):
    project = Project(desc_project)
    verdicts = _write(tmp_path, "v.json", _truth_verdicts(sample_pdf, _index(project)))
    assert main(["judge", str(desc_project), "--question", "D1", "--input", verdicts]) == 0
    bridge, _ = open_bridge(project)
    assert bridge.set_descriptive_score("小テスト_p005.png", "D1", 1)["ok"]   # 誤 → 部分
    assert bridge.set_descriptive_score("小テスト_p003.png", "D1", 2)["ok"]   # 保留に点（数えない）
    capsys.readouterr()
    assert main(["accuracy", str(desc_project)]) == 0
    out = capsys.readouterr().out
    assert "一致 4/5（80%）・保留 1" in out
    assert "食い違い: 小テスト_p005.png  LLM 誤 → 先生 1点" in out
    report = json.loads((project.root / "精度.json").read_text(encoding="utf-8"))
    assert report["questions"][0]["disagreements"][0]["file"] == "小テスト_p005.png"


def test_agrees_rules():
    from saiten_llm.cmd_desc import agrees
    assert agrees("正", 10, 10) and not agrees("正", 6, 10)
    assert agrees("誤", 0, 10) and not agrees("誤", 3, 10)
    assert agrees("部分", 6, 10) and not agrees("部分", 10, 10) and not agrees("部分", 0, 10)
    assert agrees("部分", 1, 2) and agrees("保留", 2, 2) is None and agrees("正", None, 2) is None


def test_desc_set_refuses_to_replace_scored_questions(desc_project, sample_pdf, tmp_path):
    project = Project(desc_project)
    verdicts = _write(tmp_path, "v.json", _truth_verdicts(sample_pdf, _index(project)))
    assert main(["judge", str(desc_project), "--question", "D1", "--input", verdicts]) == 0
    regions = _write(tmp_path, "r.json", {"questions": [{"name": "x", "region": [1, 1, 50, 50]}]})
    assert main(["desc-set", str(desc_project), "--input", regions]) == 1
    assert main(["desc-set", str(desc_project), "--input", regions, "--replace"]) == 1


def test_confirm_answer_key_and_marks(recognized_project):
    # ファイルでないもの（フォルダ）を渡しても落ちずに断る
    assert main(["answer-key", str(recognized_project), "--from-xlsx",
                 str(Project(recognized_project).root)]) == 1
    from tests.llm_workflow.conftest import COORD
    key = COORD.parent / "answer_key_sample.xlsx"
    assert main(["answer-key", str(recognized_project), "--from-xlsx", str(key)]) == 0
    assert main(["confirm", str(recognized_project), "answer-key"]) == 0
    assert main(["confirm", str(recognized_project), "marks"]) == 0
    data = Project(recognized_project).data
    assert data["answer_key"]["teacher_confirmed_at"] and data["answer_key"]["hash"]
    assert data["marks_check"]["hash"]
