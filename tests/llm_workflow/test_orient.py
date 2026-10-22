"""向きの判定（orient.py）と orient コマンド"""
import cv2
import numpy as np

from omr_engine import find_sheet_markers, imread_unicode
from saiten_llm.cli import main
from saiten_llm.orient import check_folder, judge, rotate, thumbnail, to_gray
from saiten_llm.workspace import Project


def _page_rotations(truth):
    """合成答案で回したのと逆の角度（＝直すために時計回りに回す角度）"""
    undo = {0: 0, 180: 180, 90: 270}
    return {s["page"]: undo[s["rotation"]] for s in truth["students"]}


def test_markers_decide_every_page(project_dir, sample_pdf):
    assert main(["pdf", str(project_dir)]) == 0
    files = Project(project_dir).scan_files()
    report = check_folder(files)
    expected = _page_rotations(sample_pdf[1])
    assert len(files) == len(expected)
    for i, path in enumerate(files, 1):
        res = report["pages"][path.name]
        assert res["method"] == "marker", (path.name, res)
        assert res["rotation"] == expected[i], (path.name, res)
        assert not res["needs_review"]


def test_orient_command_reports_then_applies(project_dir):
    assert main(["pdf", str(project_dir)]) == 0
    # --apply なしでは回さず、先生の確認待ち（2）で終わる
    assert main(["orient", str(project_dir)]) == 2
    data = Project(project_dir).data["orientation"]
    assert data["applied"] is False and data["pending"]
    assert main(["orient", str(project_dir), "--apply"]) == 0
    # 回した後はすべて正しい向き・全員マーカーあり
    report = check_folder(Project(project_dir).scan_files())
    assert all(r["rotation"] == 0 for r in report["pages"].values())
    assert Project(project_dir).data["orientation"]["with_markers"] == 6


def test_rotation_history_is_kept(project_dir):
    """何度 orient を実行しても、どの答案を回したかの記録は消えない"""
    assert main(["pdf", str(project_dir)]) == 0
    assert main(["orient", str(project_dir), "--apply"]) == 0
    first = dict(Project(project_dir).data["orientation"]["rotated"])
    assert first
    assert main(["orient", str(project_dir), "--apply"]) == 0
    assert Project(project_dir).data["orientation"]["rotated"] == first


def _erase_markers(img):
    """マーカーを白く塗りつぶして「マーカーの無い用紙」を作る"""
    gray = to_gray(img)
    h, w = gray.shape
    out = img.copy()
    for band in (slice(0, int(h * 0.08)), slice(int(h * 0.92), h)):
        out[band][gray[band] < 128] = 255
    return out


def test_similarity_used_without_markers(oriented_project):
    files = Project(oriented_project).scan_files()
    upright = _erase_markers(imread_unicode(str(files[0])))
    other = _erase_markers(imread_unicode(str(files[1])))
    assert find_sheet_markers(upright) is None
    ref = thumbnail(to_gray(upright))
    res = judge(rotate(other, 180), ref)
    assert res["method"] == "similarity"
    assert res["rotation"] == 180 and not res["needs_review"], res
    res = judge(other, ref)
    assert res["rotation"] == 0


def test_blank_page_is_flagged():
    blank = np.full((1684, 1190, 3), 255, np.uint8)
    res = judge(blank)
    assert res["blank"] and res["needs_review"]


def test_landscape_page_needs_quarter_turn(oriented_project):
    files = Project(oriented_project).scan_files()
    img = imread_unicode(str(files[0]))
    res = judge(cv2.rotate(img, cv2.ROTATE_90_COUNTERCLOCKWISE))
    assert res["method"] == "marker" and res["rotation"] == 90
