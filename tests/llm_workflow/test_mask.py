"""黒塗り（mask.py）と LLM 用の写し"""
import hashlib

import numpy as np
import pytest

from saiten_llm import mask as maskmod
from saiten_llm.cli import main
from saiten_llm.workspace import Project

NAME_REGION = [0.24, 0.2, 0.57, 0.255]   # 合成答案の氏名欄（補正後の用紙に対する割合）


def _hashes(project):
    return {p.name: hashlib.sha256(p.read_bytes()).hexdigest()
            for p in project.scan_files()}


def test_normalize_rect_orders_and_clips():
    assert maskmod.normalize_rect([0.5, 0.6, 0.1, -0.2]) == [0.1, 0.0, 0.5, 0.6]
    with pytest.raises(ValueError):
        maskmod.normalize_rect([0.1, 0.1, 0.1005, 0.5])
    with pytest.raises(ValueError):
        maskmod.normalize_rect(["a", 0, 1, 1])


def test_apply_mask_is_opaque_and_copies():
    img = np.full((100, 200, 3), 200, np.uint8)
    out = maskmod.apply_mask(img, [[0.1, 0.1, 0.5, 0.5]])
    assert (out[10:50, 20:100] == 0).all()     # 完全な黒（半透明ではない）
    assert (img == 200).all()                   # 元の配列は変えない


def test_status_goes_none_draft_approved_stale(oriented_project):
    project = Project(oriented_project)
    assert maskmod.mask_status(project) == "none"
    spec = maskmod.validate_spec({"regions": [NAME_REGION]},
                                 {p.name for p in project.scan_files()})
    maskmod.save_spec(project, spec)
    assert maskmod.mask_status(project) == "draft"
    maskmod.save_spec(project, spec, approved=True)
    assert maskmod.mask_status(project) == "approved"
    # 承認の後で答案が変わった（ここでは1枚回した）ら、承認は無効
    from saiten_llm.orient import apply_rotation
    apply_rotation(project.scan_files()[0], 180)
    assert maskmod.mask_status(project) == "stale"


def test_llm_view_requires_approval(oriented_project):
    project = Project(oriented_project)
    with pytest.raises(PermissionError):
        maskmod.build_llm_view(project)
    maskmod.save_spec(project, {"regions": [NAME_REGION], "extra": {}})
    with pytest.raises(PermissionError):
        maskmod.build_llm_view(project)
    assert not (project.llm_view / "pages").exists()
    assert main(["llm-view", str(oriented_project)]) == 2


def test_llm_view_masks_copies_but_never_scans(oriented_project):
    project = Project(oriented_project)
    before = _hashes(project)
    names = [p.name for p in project.scan_files()]
    extra = [0.6, 0.6, 0.7, 0.7]
    maskmod.save_spec(project, {"regions": [NAME_REGION], "extra": {names[1]: [extra]}},
                      approved=True)
    assert maskmod.build_llm_view(project) == len(names)
    assert _hashes(project) == before            # 元画像には一切書き込まない

    import cv2
    for i, name in enumerate(names):
        view = cv2.imread(str(project.llm_view / "pages" / (name.rsplit(".", 1)[0] + ".jpg")))
        h, w = view.shape[:2]
        x1, y1, x2, y2 = NAME_REGION
        inner = view[int(y1 * h) + 3:int(y2 * h) - 3, int(x1 * w) + 3:int(x2 * w) - 3]
        assert inner.max() < 40, name              # 氏名欄は真っ黒（JPEG の誤差だけ許す）
        ex = view[int(0.6 * h) + 3:int(0.7 * h) - 3, int(0.6 * w) + 3:int(0.7 * w) - 3]
        assert (ex.max() < 40) == (i == 1), name   # 追加の範囲はその答案だけ
    assert project.data["llm_view"]["unaligned"] == []


def test_name_is_actually_hidden(oriented_project, sample_pdf):
    """黒塗りの前は氏名欄に文字があり、後は無い（検査が空振りしていないことの確認）"""
    import cv2
    project = Project(oriented_project)
    first = project.scan_files()[0]
    aligned, ok = maskmod.load_aligned(first)
    assert ok
    h, w = aligned.shape[:2]
    x1, y1, x2, y2 = NAME_REGION
    before = cv2.cvtColor(aligned[int(y1 * h):int(y2 * h), int(x1 * w):int(x2 * w)],
                          cv2.COLOR_BGR2GRAY)
    assert (before < 100).sum() > 200              # 名前の文字がある
    after = maskmod.apply_mask(aligned, [NAME_REGION])
    region = after[int(y1 * h) + 1:int(y2 * h) - 1, int(x1 * w) + 1:int(x2 * w) - 1]
    assert region.max() == 0
