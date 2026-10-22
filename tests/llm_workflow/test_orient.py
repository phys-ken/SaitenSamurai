"""向きの判定（orient.py）と orient コマンド"""
import cv2
import numpy as np
import pytest

from omr_engine import find_sheet_markers, imread_unicode
from saiten_llm.cli import main
from saiten_llm.imageio import read_image, write_image
from saiten_llm.orient import (candidate_rotations, check_folder, judge, rotate, thumbnail,
                               to_gray)
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


# --- 横長の用紙（B4横・マーカー無し） -----------------------------------

B4_LANDSCAPE = (1456, 1030)     # 幅x高さ（B4横を小さくしたもの）


def _b4_sheet(student):
    """マーカーの無い B4横の解答用紙（合成）。上下・左右が見分けられる配置にする"""
    w, h = B4_LANDSCAPE
    img = np.full((h, w, 3), 255, np.uint8)
    cv2.rectangle(img, (40, 30), (700, 110), (30, 30, 30), -1)          # 左上の見出し
    cv2.putText(img, f"student {student:02d}", (760, 90), cv2.FONT_HERSHEY_SIMPLEX,
                1.6, (0, 0, 0), 3)
    for i, (x0, y0, x1, y1) in enumerate([(40, 160, 980, 420), (40, 470, 980, 730),
                                          (40, 780, 620, 990)]):
        cv2.rectangle(img, (x0, y0), (x1, y1), (40, 40, 40), 3)
        cv2.putText(img, f"Q{i + 1} answer {student}", (x0 + 30, y0 + 120),
                    cv2.FONT_HERSHEY_SIMPLEX, 2.0, (0, 0, 0), 4)
    return img


# スキャンしたときの向き（時計回りに回った角度）。1枚目は縦に倒れている。
# 縦に倒れた基準は 90° 回して横長に置く決まりなので、1枚目はそれで正しく立つ向き（270°）にする
_SCANNED = {1: 270, 2: 0, 3: 180, 4: 90, 5: 270, 6: 0}


def _undo(scanned):
    return (360 - scanned) % 360


def _make_b4(root, scanned):
    (root / "scans").mkdir(parents=True)
    for student, rot in scanned.items():
        write_image(root / "scans" / f"b4_p{student:03d}.png", rotate(_b4_sheet(student), rot))
    return root


@pytest.fixture
def b4_project(tmp_path):
    return _make_b4(tmp_path / "B4横の小テスト", _SCANNED)


def _matches_upright(path, student):
    """ファイルの画像が、正しい向きの用紙と同じ向き・同じ形か"""
    img = read_image(path)
    return img.shape == _b4_sheet(student).shape and _matches_upright_after(img, student, 0)


def test_landscape_candidates():
    gray = to_gray(_b4_sheet(1))
    assert candidate_rotations(gray) == (90, 270)                  # 既定は縦長の用紙
    assert candidate_rotations(gray, landscape=True) == (0, 180)
    assert candidate_rotations(gray.T, landscape=True) == (90, 270)


def test_check_folder_landscape_reference_and_rotations(b4_project):
    files = Project(b4_project).scan_files()
    report = check_folder(files, landscape=True)
    assert report["reference"] == "b4_p001.png"
    pages = report["pages"]
    # 基準（1枚目）は縦長のスキャンなので、横長に置く 90° が選ばれる
    assert pages["b4_p001.png"]["method"] == "reference"
    assert pages["b4_p001.png"]["rotation"] == 90 == _undo(_SCANNED[1])
    for student, scanned in _SCANNED.items():
        res = pages[f"b4_p{student:03d}.png"]
        assert res["rotation"] == _undo(scanned), (student, res)
        assert not res["needs_review"], (student, res)


def test_orient_landscape_command(b4_project):
    r = str(b4_project)
    # 指定しないと、横長の答案を縦に倒そうとする（これまでの前提）
    assert main(["orient", r]) == 2
    pending = Project(b4_project).data["orientation"]["pending"]
    assert pending["b4_p002.png"] in (90, 270)

    assert main(["orient", r, "--landscape", "--apply"]) == 0
    project = Project(b4_project)
    assert project.data["paper"] == "landscape"
    assert project.data["orientation"]["paper"] == "landscape"
    for student in _SCANNED:
        assert _matches_upright(project.scans / f"b4_p{student:03d}.png", student), student
    assert project.data["orientation"]["rotated"] == {
        f"b4_p{s:03d}.png": _undo(rot) for s, rot in _SCANNED.items() if rot}

    # 2回目は --landscape を付けなくても横長のまま（何も回さない）
    assert main(["orient", r, "--apply"]) == 0
    project = Project(b4_project)
    assert project.data["orientation"]["pending"] == {}
    for student in _SCANNED:
        assert _matches_upright(project.scans / f"b4_p{student:03d}.png", student), student

    # 縦長に戻すと、また縦に倒す候補になる
    assert main(["orient", r, "--portrait"]) == 2
    assert Project(b4_project).data["paper"] == "portrait"


def test_landscape_upside_down_reference_is_fixed_by_flip_all(tmp_path):
    """縦に倒れた基準を 90° 置いたら逆さだった → 全員が同じく逆さに揃い、--flip-all で直る"""
    scanned = {1: 90, 2: 0, 3: 180, 4: 270}
    root = _make_b4(tmp_path / "B4横", scanned)
    assert main(["orient", str(root), "--landscape", "--apply"]) == 0
    scans = Project(root).scans
    for student in scanned:
        img = read_image(scans / f"b4_p{student:03d}.png")
        assert img.shape[1] > img.shape[0], student                  # 横長には置けている
        assert _matches_upright_after(img, student, 180), student    # 全員そろって逆さ
    assert main(["orient", str(root), "--flip-all", "--apply"]) == 0
    for student in scanned:
        assert _matches_upright(scans / f"b4_p{student:03d}.png", student), student
    # 回した角度の合計は、スキャンの傾きを戻す角度になっている
    assert Project(root).data["orientation"]["rotated"] == {
        f"b4_p{s:03d}.png": _undo(rot) for s, rot in scanned.items() if rot}


def _matches_upright_after(img, student, degrees):
    return float(abs(rotate(img, degrees).astype(int) - _b4_sheet(student)).mean()) < 2


def test_landscape_setting_survives_pdf_rebuild(b4_project):
    """pdf を作り直しても（orientation は消える）用紙の縦横の指定は残る"""
    project = Project(b4_project)
    project.data["paper"] = "landscape"
    project.data.pop("orientation", None)
    project.save()
    assert main(["orient", str(b4_project), "--apply"]) == 0
    assert Project(b4_project).data["orientation"]["paper"] == "landscape"
