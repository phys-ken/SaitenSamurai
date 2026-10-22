"""黒塗りの範囲を決める画面（saiten_llm/mask_page.html）をブラウザで操作する（L2）。

本物の MaskServer を合成答案で起動し、先生と同じ手順
（ドラッグで範囲 → 全員分を確かめる → 確認のチェック → 承認）をなぞる。
"""
import shutil
import sys
import threading
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(REPO))

import saiten_llm  # noqa: E402,F401
from saiten_llm.cli import main  # noqa: E402
from saiten_llm.mask import mask_status  # noqa: E402
from saiten_llm.mask_server import MaskServer  # noqa: E402
from saiten_llm.tools.make_sample_exam import make_exam  # noqa: E402
from saiten_llm.workspace import Project  # noqa: E402


@pytest.fixture(scope="module")
def exam_template(tmp_path_factory):
    root = tmp_path_factory.mktemp("mask") / "小テスト"
    root.mkdir()
    make_exam(root / "小テスト.pdf", students=3, seed=5, noise=False)
    (root / "小テスト_truth.json").unlink()
    assert main(["pdf", str(root)]) == 0
    assert main(["orient", str(root), "--apply"]) == 0
    return root


@pytest.fixture
def mask_server(tmp_path, exam_template):
    root = tmp_path / "小テスト"
    shutil.copytree(exam_template, root)
    srv = MaskServer(Project(root), timeout_minutes=2)
    thread = threading.Thread(target=srv.serve_until_done, kwargs={"poll": 0.05}, daemon=True)
    thread.start()
    yield srv
    srv.approved.set()
    thread.join(5)


def _drag(page, box, x1, y1, x2, y2):
    page.mouse.move(box["x"] + box["width"] * x1, box["y"] + box["height"] * y1)
    page.mouse.down()
    page.mouse.move(box["x"] + box["width"] * x2, box["y"] + box["height"] * y2, steps=5)
    page.mouse.up()


def test_draw_check_and_approve(page, mask_server):
    page.goto(mask_server.url)
    page.wait_for_function("document.querySelector('#page-img').naturalWidth > 0")
    box = page.locator("#draw").bounding_box()
    _drag(page, box, 0.24, 0.2, 0.57, 0.255)
    assert page.locator("#rect-list li").count() == 1

    page.click("#to-grid")
    page.wait_for_selector(".thumb")
    assert page.locator(".thumb").count() == 3
    # 確認のチェックを入れるまで承認できない
    assert page.locator("#approve").is_disabled()
    assert mask_status(Project(mask_server.project.root)) == "draft"

    # 1人だけ範囲を足す
    page.locator(".thumb").nth(1).click()
    page.wait_for_function("document.querySelector('#page-img').naturalWidth > 0")
    assert page.locator("#editor-title").inner_text() == "この答案だけの範囲"
    _drag(page, page.locator("#draw").bounding_box(), 0.6, 0.6, 0.7, 0.7)
    page.click("#to-grid")
    page.wait_for_selector(".thumb.extra")

    page.check("#confirm")
    page.click("#approve")
    page.wait_for_selector("#done:not([hidden])")
    assert mask_server.approved.wait(5)

    spec = Project(mask_server.project.root).data["mask"]
    assert mask_status(Project(mask_server.project.root)) == "approved"
    assert len(spec["regions"]) == 1
    x1, y1, x2, y2 = spec["regions"][0]
    assert abs(x1 - 0.24) < 0.02 and abs(y2 - 0.255) < 0.02
    assert list(spec["extra"]) == [mask_server.names[1]]


def test_zero_regions_warns_before_approval(page, mask_server):
    page.goto(mask_server.url)
    page.wait_for_function("document.querySelector('#page-img').naturalWidth > 0")
    page.click("#to-grid")
    page.wait_for_selector(".thumb")
    assert page.locator("#zero-note").is_visible()
