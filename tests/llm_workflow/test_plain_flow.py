"""マーカーの無い普通のプリント（記述のみ）を、PDF から個票まで通す。

最初の本番（手元の小テスト）がこの形式なので、合成の答案で道筋を確かめておく。
"""
import json
import shutil

import cv2
import pytest

from saiten_llm import mask as maskmod
from saiten_llm.cli import main
from saiten_llm.engine import open_bridge
from saiten_llm.imageio import read_image
from saiten_llm.tools.make_sample_exam import make_plain_exam
from saiten_llm.workspace import Project


@pytest.fixture(scope="module")
def plain_template(tmp_path_factory):
    root = tmp_path_factory.mktemp("plain") / "理科小テスト"
    root.mkdir()
    pdf, truth_path = make_plain_exam(root / "理科.pdf", students=8, seed=4)
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth_path.unlink()
    return root, truth


@pytest.fixture
def plain(tmp_path, plain_template):
    src, truth = plain_template
    root = tmp_path / "理科小テスト"
    shutil.copytree(src, root)
    return root, truth


def _write(tmp_path, name, data):
    p = tmp_path / name
    p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return str(p)


def _frac(box, size):
    w, h = size
    return [box[0] / w, box[1] / h, box[2] / w, box[3] / h]


def _confirmed_plain(root, truth, tmp_path):
    """PDF から、記述3問を判定して先生が確定するところまで。Returns: 記述問題の一覧"""
    r = str(root)
    assert main(["pdf", r]) == 0
    assert main(["orient", r, "--apply"]) == 0
    project = Project(root)
    rotated = project.data["orientation"]["rotated"]
    assert rotated == {"理科_p003.png": 180}
    assert project.data["orientation"]["with_markers"] == 0

    # 黒塗り: 氏名欄を少し広めに（ずれの分）
    x0, y0, x1, y1 = truth["name_box"]
    name = [x0 - 20, y0 - 20, x1 + 20, y1 + 20]
    maskmod.save_spec(project, {"regions": [_frac(name, truth["size"])], "extra": {}}, approved=True)
    assert main(["llm-view", r]) == 0
    assert Project(root).data["llm_view"]["unaligned"] == sorted(p.name for p in project.scan_files())

    # 全員の氏名欄が写しで真っ黒（傾き・ずれがあっても塗り漏れない）
    for page in sorted((project.llm_view / "pages").glob("*.jpg")):
        img = read_image(page, cv2.IMREAD_GRAYSCALE)
        inner = img[y0 + 2:y1 - 2, x0 + 2:x1 - 2]
        assert inner.max() < 40, page.name

    assert main(["mode", r, "--set", "descriptive_only", "--teacher-approved"]) == 0
    assert main(["prepare", r]) == 0
    regions = {"questions": [{"name": q["name"], "region": list(q["box"])} for q in truth["questions"]]}
    assert main(["desc-set", r, "--input", _write(tmp_path, "regions.json", regions)]) == 0
    bridge, _ = open_bridge(Project(root), write=False)
    qs = bridge.state["descriptive"]["questions"]
    assert [q["region"] for q in qs] == [list(q["box"]) for q in truth["questions"]]  # 倍率 1
    assert main(["confirm", r, "regions"]) == 0
    assert main(["crops", r, "--per-sheet", "8"]) == 0

    # 正解どおりに判定（ひらがなは保留）→ 先生が保留に点を入れて確定
    by_page = {f"理科_p{s['page']:03d}.png": s["answers"] for s in truth["students"]}
    for q in qs:
        index = json.loads((Project(root).llm_view / "crops" / q["id"] / "index.json")
                           .read_text(encoding="utf-8"))["numbers"]
        verdicts = {}
        for n, f in index.items():
            a = by_page[f][q["name"]]
            if a["text"] in ("こうごうせい", "たいよう"):
                verdicts[n] = {"verdict": "保留", "reason": "ひらがな"}
            elif a["truth"] == "部分":
                verdicts[n] = {"verdict": "部分", "reason": "誤字"}
            else:
                verdicts[n] = "正" if a["truth"] == "正" else "誤"
        assert main(["judge", r, "--question", q["id"], "--input",
                     _write(tmp_path, f"{q['id']}.json", verdicts)]) == 0
    bridge, _ = open_bridge(Project(root))
    for f, answers in by_page.items():
        for q in qs:
            if answers[q["name"]]["text"] in ("こうごうせい", "たいよう"):
                assert bridge.set_descriptive_score(f, q["id"], 2)["ok"]
    for q in qs:
        assert main(["confirm", r, "question", "--question", q["id"]]) == 0
    return qs


def test_plain_exam_end_to_end(plain, tmp_path):
    root, truth = plain
    r = str(root)
    qs = _confirmed_plain(root, truth, tmp_path)
    by_page = {f"理科_p{s['page']:03d}.png": s["answers"] for s in truth["students"]}

    assert main(["stats", r, "--total", "30"]) == 0
    points = {"descriptive": {q["id"]: 10 for q in qs}}
    assert main(["set-points", r, "--input", _write(tmp_path, "p.json", points),
                 "--teacher-approved"]) == 0
    assert main(["finalize", r]) == 2              # 部分点の入力待ち
    bridge, _ = open_bridge(Project(root))
    for f, answers in by_page.items():
        for q in qs:
            if answers[q["name"]]["truth"] == "部分":
                assert bridge.set_descriptive_score(f, q["id"], 5)["ok"]
    for q in qs:
        assert main(["confirm", r, "question", "--question", q["id"]]) == 0
    assert main(["finalize", r]) == 0

    from constants import RESULTS_FOLDER, SCORED_FOLDER
    out = Project(root).scans / RESULTS_FOLDER / SCORED_FOLDER
    assert len(list(out.glob("*"))) >= 8


def test_partial_given_for_some_questions(plain, tmp_path):
    """部分のある問題だけ partial を書く（部分の無い問題は書かなくてよい）。
    全問の確定が保たれ、確定し直さずに finalize まで進める"""
    from saiten_llm import judgments as jmod
    from saiten_llm.cmd_finalize import gate
    root, truth = plain
    r = str(root)
    qs = _confirmed_plain(root, truth, tmp_path)
    by_page = {f"理科_p{s['page']:03d}.png": s["answers"] for s in truth["students"]}
    q1, q2, q3 = qs                                   # 部分があるのは 問1 だけ
    spec = {"descriptive": {q1["id"]: 10, q2["id"]: 4, q3["id"]: 6}, "partial": {q1["id"]: 4}}
    assert main(["set-points", r, "--input", _write(tmp_path, "p.json", spec),
                 "--teacher-approved"]) == 0
    project = Project(root)
    bridge, _ = open_bridge(project, write=False)
    scores = bridge._desc_scores["scores"]
    partial_files = [f for f, a in by_page.items() if a["問1"]["truth"] == "部分"]
    assert partial_files                              # 検査が空振りしていない
    for f in partial_files:
        assert scores[f][q1["id"]] == 4
    data = jmod.load(project)
    for q, new_max in ((q1, 10), (q2, 4), (q3, 6)):
        confirmed = data["questions"][q["id"]]["confirmed"]
        assert confirmed and confirmed["max_score"] == new_max, q["id"]
    assert data["questions"][q1["id"]]["confirmed"]["by"] == "確定済みの正誤・部分を換算"
    assert gate(project, bridge) == []
    assert main(["finalize", r]) == 0


def test_crops_follow_the_answer_boxes(plain, tmp_path):
    """切り抜きが解答欄に当たっている（枠の線が切り抜きの端にある）"""
    root, truth = plain
    r = str(root)
    assert main(["pdf", r]) == 0 and main(["orient", r, "--apply"]) == 0
    project = Project(root)
    maskmod.save_spec(project, {"regions": [], "extra": {}}, approved=True)
    assert main(["llm-view", r]) == 0
    assert main(["mode", r, "--set", "descriptive_only", "--teacher-approved"]) == 0
    assert main(["prepare", r]) == 0
    q = truth["questions"][1]
    # マーカーの無い用紙は傾きを補正できないので、手順書どおり枠より20画素広く取る
    x0, y0, x1, y1 = q["box"]
    wide = [x0 - 20, y0 - 20, x1 + 20, y1 + 20]
    assert main(["desc-set", r, "--input",
                 _write(tmp_path, "r.json", {"questions": [{"name": q["name"], "region": wide}]})]) == 0
    from saiten_llm.montage import crop
    bridge, _ = open_bridge(Project(root), write=False)
    region = bridge.state["descriptive"]["questions"][0]["region"]
    for page in sorted((project.llm_view / "pages").glob("*.jpg")):
        c = cv2.cvtColor(crop(read_image(page), region, 1.0), cv2.COLOR_BGR2GRAY)
        # 枠の上の横線が、傾いていても全幅にわたって切り抜きの中に入っている
        top = (c[:45, 20:-20] < 110).any(axis=0).mean()
        bottom = (c[-45:, 20:-20] < 110).any(axis=0).mean()
        assert top > 0.97 and bottom > 0.97, (page.name, top, bottom)
