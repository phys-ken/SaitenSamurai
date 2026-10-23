"""返却答案（個票）の見た目: ○×△の大きさ・合計点の枠・プレビュー"""
import hashlib
import json
import shutil
import threading
import urllib.request

import cv2
import numpy as np
import pytest

from saiten_llm import mask as maskmod
from saiten_llm import render
from saiten_llm.cli import main
from saiten_llm.engine import open_bridge
from saiten_llm.imageio import read_image
from saiten_llm.tools.make_sample_exam import make_b4_exam
from saiten_llm.workspace import Project

VERDICT = {"正": 2, "部分": 1, "誤": 0}


def _red_bbox(image):
    """赤い（記号の）画素の外接矩形 (x1, y1, x2, y2)。無ければ None"""
    b, g, r = (image[..., i].astype(int) for i in range(3))
    ys, xs = np.nonzero((r - g > 60) & (r - b > 60))
    if not len(xs):
        return None
    return xs.min(), ys.min(), xs.max(), ys.max()


def _blue_mask(image):
    b, g, r = (image[..., i].astype(int) for i in range(3))
    return (b - r > 60) & (b - g > 60)


def _draw(region, score=2, fit=True, rs=None, size=(2024, 2866)):
    from descriptive_renderer import draw_descriptive_on_image
    img = np.full((size[1], size[0], 3), 255, np.uint8)
    config = {"questions": [{"id": "D1", "name": "問1", "region": list(region),
                             "aspect": 1, "max_score": 2}]}
    return draw_descriptive_on_image(img, config, {"D1": score}, output_scale=1.0,
                                     rendering_settings=rs, fit_to_region=fit)


SYMBOL_ONLY = {"descriptive_show_score": False, "descriptive_show_aspect": False}


# --- ○×△ の大きさ -------------------------------------------------------

def test_symbol_size_follows_the_answer_box():
    """記述のみモードの記号は解答欄の大きさに比例し、欄の中央 80% に収まる"""
    heights = {}
    for name, region in {"small": (100, 100, 530, 330), "large": (100, 100, 1000, 600)}.items():
        img = _draw(region, rs=SYMBOL_ONLY)
        x1, y1, x2, y2 = _red_bbox(img)
        rw, rh = region[2] - region[0], region[3] - region[1]
        # 欄の中央 80% の中にある
        assert x1 >= region[0] + 0.1 * rw - 2 and x2 <= region[2] - 0.1 * rw + 2, name
        assert y1 >= region[1] + 0.1 * rh - 2 and y2 <= region[3] - 0.1 * rh + 2, name
        # 記号だけなら欄の中央に来る
        assert abs((x1 + x2) / 2 - (region[0] + region[2]) / 2) < 0.05 * rw, name
        assert abs((y1 + y2) / 2 - (region[1] + region[3]) / 2) < 0.05 * rh, name
        heights[name] = y2 - y1
        assert heights[name] > 0.35 * min(rw, rh), (name, heights[name])   # 60 画素の上限で止まらない
    ratio = heights["large"] / heights["small"]
    assert 500 / 230 * 0.8 < ratio < 500 / 230 * 1.2, heights        # 欄の短辺に比例


def test_symbol_score_aspect_line_fits_in_box():
    """記号・得点・観点の1行も、大きくなって欄の中央 80% からはみ出さない"""
    region = (100, 100, 530, 330)
    img = _draw(region, score=1)
    dark = (img.min(axis=2) < 200)
    ys, xs = np.nonzero(dark)
    assert xs.min() >= 100 + 43 - 2 and xs.max() <= 530 - 43 + 2
    assert ys.min() >= 100 + 23 - 2 and ys.max() <= 330 - 23 + 2


def test_other_modes_keep_the_60px_limit():
    """fit_to_region を付けない呼び出し（マーク＋記述など）は従来どおり上限 60 画素"""
    from descriptive_renderer import draw_descriptive_on_image
    region = (100, 100, 1000, 600)
    old = _draw(region, rs=SYMBOL_ONLY, fit=False)
    x1, y1, x2, y2 = _red_bbox(old)
    assert y2 - y1 <= 62
    # 引数を省いたときも同じ（既定は従来の動き）
    img = np.full((2866, 2024, 3), 255, np.uint8)
    config = {"questions": [{"id": "D1", "name": "問1", "region": list(region),
                             "aspect": 1, "max_score": 2}]}
    default = draw_descriptive_on_image(img, config, {"D1": 2}, rendering_settings=SYMBOL_ONLY)
    assert np.array_equal(default, old)


def test_total_is_drawn_inside_the_region():
    from descriptive_scorer import render_descriptive_only_image
    img = np.full((2024, 2866, 3), 255, np.uint8)
    region = [2509, 143, 2767, 191]
    config = {"questions": [{"id": "D1", "name": "問1", "region": [100, 300, 1380, 730],
                             "aspect": 1, "max_score": 2}], "total_display_region": region}
    out = render_descriptive_only_image(img, config, {"D1": 1},
                                        {"total_show_aspects": False})
    ys, xs = np.nonzero(_blue_mask(out))
    assert len(xs) > 50                                   # 検査が空振りしていない
    assert xs.min() >= region[0] and xs.max() <= region[2]
    assert ys.min() >= region[1] and ys.max() <= region[3]


# --- 合成の B4 答案で、コマンドを通す ---------------------------------------

def _build(base, portrait, mask_touches=False):
    root = base / ("B4縦の小テスト" if portrait else "B4横の小テスト")
    root.mkdir(parents=True)
    _pdf, truth_path = make_b4_exam(root / "物理.pdf", students=5, seed=3, portrait=portrait)
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth_path.unlink()
    r = str(root)
    assert main(["pdf", r]) == 0
    assert main(["orient", r, "--apply"] + ([] if portrait else ["--landscape"])) == 0
    w, h = truth["size"]
    x0, y0, x1, y1 = truth["name_box"]
    if not mask_touches:
        x0, y0, x1, y1 = x0 + 5, y0 + 5, x1 - 5, y1 - 5
    maskmod.save_spec(Project(root), {"regions": [[x0 / w, y0 / h, x1 / w, y1 / h]], "extra": {}},
                      approved=True)
    assert main(["llm-view", r]) == 0
    assert main(["mode", r, "--set", "descriptive_only", "--teacher-approved"]) == 0
    assert main(["prepare", r]) == 0
    regions = base / "regions.json"
    regions.write_text(json.dumps({"questions": [{"name": q["name"], "region": q["box"]}
                                                 for q in truth["questions"]]},
                                  ensure_ascii=False), encoding="utf-8")
    assert main(["desc-set", r, "--input", str(regions)]) == 0
    bridge, _ = open_bridge(Project(root))
    qs = bridge.state["descriptive"]["questions"]
    for s in truth["students"]:
        for q in qs:
            assert bridge.set_descriptive_score(f"物理_p{s['page']:03d}.png", q["id"],
                                                VERDICT[s["answers"][q["name"]]])["ok"]
    bridge.save_session()
    return root, truth


@pytest.fixture(scope="module")
def b4_templates(tmp_path_factory):
    base = tmp_path_factory.mktemp("b4")
    return {"landscape": _build(base / "l", False), "portrait": _build(base / "p", True)}


@pytest.fixture(params=["landscape", "portrait"])
def b4(request, tmp_path, b4_templates):
    from saiten_llm.engine import session_path
    src, truth = b4_templates[request.param]
    root = tmp_path / src.name
    shutil.copytree(src, root)
    # 採点侍の設定は絶対パスを持つので、写した先に合わせて直す（conftest と同じ）
    path = session_path(Project(root))
    data = json.loads(path.read_text(encoding="utf-8"))
    data["image_folder"] = str(root / "scans")
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return root, truth


def _close(a, b, tol=6):
    return all(abs(x - y) <= tol for x, y in zip(a, b))


def test_detect_total_box(b4):
    """横長は右上の隅の欄、縦長は2つに分かれて見える欄を1つとして見つける"""
    root, truth = b4
    view = render.view_page(Project(root), "物理_p001.png")
    found = render.detect_total_box(view)
    assert found is not None and _close(found, truth["score_box"]), (found, truth["score_box"])


def test_detect_when_mask_touches_the_box(tmp_path):
    """黒塗りが「得点」欄の隣の線まで掛かっていても、欄の内側で見つける"""
    root, truth = _build(tmp_path, portrait=False, mask_touches=True)
    found = render.detect_total_box(render.view_page(Project(root), "物理_p001.png"))
    assert found is not None and _close(found, truth["score_box"]), found


def test_detect_returns_none_without_box():
    blank = np.full((2024, 2866, 3), 255, np.uint8)
    cv2.rectangle(blank, (100, 300), (1380, 730), (40, 40, 40), 3)     # 解答欄だけ
    assert render.detect_total_box(blank) is None


def test_total_region_command(b4, capsys):
    root, truth = b4
    r = str(root)
    assert main(["total-region", r, "--detect", "--apply"]) == 0
    region = open_bridge(Project(root), write=False)[0].get_total_display_region()["region"]
    x1, y1, x2, y2 = truth["score_box"]
    assert _close(region, [x1, (y1 + y2) // 2, x2, y2])                  # 欄の下半分
    assert (Project(root).llm_view / "total_region.jpg").exists()
    # 手で書く枠: 答案の外は断る
    assert main(["total-region", r, "--set", "10,10,99999,50"]) == 1
    assert main(["total-region", r, "--set", "100,100,300,160"]) == 0
    region = open_bridge(Project(root), write=False)[0].get_total_display_region()["region"]
    assert region == [100, 100, 300, 160]
    assert main(["total-region", r, "--clear"]) == 0
    assert open_bridge(Project(root), write=False)[0].get_total_display_region()["region"] is None


# --- 目で見て枠を決めるための画像（--crop・--set の重ね描き） -------------------

def _red_centroids(image):
    """赤い点の中心（原寸での目印の位置を、拡大した画像の中で探す）。

    目盛りの線が目印の上を通って赤を分けるので、少し太らせてから1つにまとめる
    """
    b, g, r = (image[..., i].astype(int) for i in range(3))
    mask = ((r - g > 120) & (r - b > 120)).astype(np.uint8)
    mask = cv2.dilate(mask, np.ones((7, 7), np.uint8))
    n, _labels, stats, centroids = cv2.connectedComponentsWithStats(mask)
    return [tuple(centroids[i]) for i in range(1, n) if stats[i, cv2.CC_STAT_AREA] > 20]


@pytest.mark.parametrize("size,rect", [((2866, 2024), (1719, 0, 2866, 404)),
                                       ((2024, 2866), (809, 0, 2024, 515)),
                                       ((2866, 2024), (2400, 60, 2800, 230))])   # 4倍に拡大
def test_grid_crop_labels_are_original_pixels(size, rect):
    """目盛りの数は原寸の画素: 原寸の (x, y) に置いた目印が、目盛り x・y の線の交点に来る"""
    w, h = size
    img = np.full((h, w, 3), 255, np.uint8)
    # 切り出しの端から 20 画素以上内側の、目盛りの線の交点に目印を置く
    marks = [(x, y) for x in range(-(-(rect[0] + 20) // 100) * 100, rect[2] - 20, 200)
             for y in range(-(-(rect[1] + 20) // 100) * 100, rect[3] - 20, 100)]
    assert len(marks) >= 2                                  # 検査が空振りしていない
    for x, y in marks:
        cv2.rectangle(img, (x - 3, y - 3), (x + 3, y + 3), (0, 0, 255), -1)
    out, ticks = render.grid_crop(img, rect)
    tx, ty = dict(ticks["x"]), dict(ticks["y"])
    assert all(v % 50 == 0 for v in tx) and all(v % 50 == 0 for v in ty)
    found = _red_centroids(out)
    assert len(found) == len(marks)
    scale = (ticks["x"][1][1] - ticks["x"][0][1]) / 50
    for x, y in marks:
        X, Y = tx[x], ty[y]
        assert any(abs(cx - X) <= scale + 1 and abs(cy - Y) <= scale + 1 for cx, cy in found), (x, y)
        # 目盛りの線が本当にその位置に引かれている（白い所で線の色になる）
        assert tuple(out[Y + 20 if Y + 20 < out.shape[0] else Y - 20, X]) == pytest.approx(
            render.GRID_COLOR, abs=40)


def test_grid_crop_rejects_outside():
    img = np.full((100, 100, 3), 255, np.uint8)
    with pytest.raises(ValueError):
        render.grid_crop(img, (200, 200, 300, 300))


def test_crop_command(b4):
    root, truth = b4
    r = str(root)
    out = Project(root).llm_view / "total_region_crop.jpg"
    assert main(["total-region", r, "--crop"]) == 0                       # 右上の隅
    assert read_image(out) is not None
    out.unlink()
    x1, y1, x2, y2 = truth["score_box"]
    assert main(["total-region", r, "--crop", f"{x1 - 100},{y1 - 50},{x2},{y2 + 50}"]) == 0
    assert read_image(out) is not None
    assert main(["total-region", r, "--crop", "0,0,99999,10"]) == 1       # 答案の外


def test_set_writes_overlay_on_the_box(b4):
    """--set のあと、設定した枠（赤）を重ねた画像が書かれ、枠が欄の位置に乗っている"""
    root, truth = b4
    x1, y1, x2, y2 = truth["score_box"]
    region = [x1 + 5, (y1 + y2) // 2, x2 - 5, y2 - 5]
    assert main(["total-region", str(root), "--set", ",".join(map(str, region))]) == 0
    view = Project(root).llm_view
    assert read_image(view / "total_region.jpg") is not None
    crop = read_image(view / "total_region_crop.jpg")
    # 同じ範囲で目盛りの位置を求め、赤い線がその枠の上にあることを確かめる
    from saiten_llm.cmd_render import _boxed_size
    size = _boxed_size(Project(root), open_bridge(Project(root), write=False)[0])
    _img, ticks = render.grid_crop(np.full((size[1], size[0], 3), 255, np.uint8),
                                   render.around(region, size))
    # 倍率は最初と最後の目盛りから求める（隣どうしだと丸めの誤差が大きい）
    (v0, p0), (v1, p1) = ticks["x"][0], ticks["x"][-1]
    scale = (p1 - p0) / (v1 - v0)
    ox = lambda v: p0 + (v - v0) * scale                       # noqa: E731
    (w0, q0) = ticks["y"][0]
    oy = lambda v: q0 + (v - w0) * scale                       # noqa: E731
    b, g, r = (crop[..., i].astype(int) for i in range(3))
    red = (r - g > 120) & (r - b > 120)
    # 赤い縦線・横線の中心（JPEG で赤が横ににじむので、端でなく線の中心で比べる）
    cols, rows = red.sum(axis=0), red.sum(axis=1)
    mid_x, mid_y = int(ox((region[0] + region[2]) / 2)), int(oy((region[1] + region[3]) / 2))
    left, right = int(np.argmax(cols[:mid_x])), mid_x + int(np.argmax(cols[mid_x:]))
    top, bottom = int(np.argmax(rows[:mid_y])), mid_y + int(np.argmax(rows[mid_y:]))
    assert cols[left] > 20 and rows[top] > 20                 # 検査が空振りしていない
    assert abs(left - ox(region[0])) <= 3 and abs(right - ox(region[2])) <= 3
    assert abs(top - oy(region[1])) <= 3 and abs(bottom - oy(region[3])) <= 3


def test_detect_not_found_is_only_a_hint(b4, monkeypatch, capsys):
    root, _truth = b4
    monkeypatch.setattr(render, "detect_total_box", lambda image: None)
    assert main(["total-region", str(root), "--detect", "--apply"]) == 0
    out = capsys.readouterr().out
    assert "目で見て決めてください" in out and "--crop" in out
    assert open_bridge(Project(root), write=False)[0].get_total_display_region()["region"] is None


def test_render_settings_command(b4, capsys):
    root, _truth = b4
    r = str(root)
    assert main(["render-settings", r, "--set", "descriptive_show_aspect=maybe"]) == 1
    assert main(["render-settings", r, "--set", "no_such_key=false"]) == 1
    assert main(["render-settings", r, "--set", "descriptive_opacity=2"]) == 1
    assert main(["render-settings", r, "--set", "descriptive_show_aspect=false",
                 "total_show_aspects=false", "descriptive_opacity=0.6"]) == 0
    rs = open_bridge(Project(root), write=False)[0].state["rendering_settings"]
    # 文字の "false" を True と取り違えない
    assert rs == {"descriptive_show_aspect": False, "total_show_aspects": False,
                  "descriptive_opacity": 0.6}
    capsys.readouterr()
    assert main(["render-settings", r]) == 0
    assert "* descriptive_show_aspect = False（既定 True）" in capsys.readouterr().out


def _scan_hashes(project):
    return {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in project.scan_files()}


def test_render_preview_matches_the_real_sheet(b4, tmp_path):
    """プレビューは、黒塗りの所を除けば、返却答案の実物と同じ絵になる"""
    root, truth = b4
    r = str(root)
    project = Project(root)
    before = _scan_hashes(project)
    assert main(["total-region", r, "--detect", "--apply"]) == 0
    assert main(["render-settings", r, "--set", "descriptive_show_aspect=false",
                 "total_show_aspects=false"]) == 0
    assert main(["render-preview", r, "--pages", "3"]) == 0
    previews = sorted((project.llm_view / "preview").glob("preview_*.jpg"))
    assert len(previews) == 3
    assert _scan_hashes(project) == before                 # 元の答案には書き込まない
    picked = Project(root).data["render_preview"]["pages"]

    # 実物（返却答案の生成と同じ関数・同じ設定）を作って比べる
    from constants import BOXED_FOLDER, RESULTS_FOLDER
    from descriptive_scorer import generate_descriptive_only_sheets
    bridge, _ = open_bridge(project, write=False)
    out = tmp_path / "real"
    generate_descriptive_only_sheets(str(project.scans / RESULTS_FOLDER / BOXED_FOLDER),
                                     bridge._desc_config_with_total_region(),
                                     bridge._desc_scores["scores"], str(out),
                                     rendering_settings=bridge.state["rendering_settings"])
    w, h = truth["size"]
    nx0, ny0, nx1, ny1 = truth["name_box"]
    for path, name in zip(previews, picked):
        preview, real = read_image(path), read_image(out / name)
        assert preview.shape == real.shape
        diff = np.abs(preview.astype(int) - real.astype(int)).max(axis=2)
        diff[ny0 - 10:ny1 + 10, nx0 - 10:nx1 + 10] = 0     # 黒塗りした氏名欄は違ってよい
        assert diff.mean() < 2 and (diff > 80).mean() < 0.002, (name, diff.mean())
        # ○×△ と合計点が描かれている（検査が空振りしていない）
        assert _red_bbox(preview) is not None and _blue_mask(preview).sum() > 50


def test_copied_project_does_not_write_to_the_original(tmp_path, b4_templates):
    """設定が元のフォルダを指したままの写しは開かない（元の採点を書き換えないため）"""
    src, _truth = b4_templates["landscape"]
    root = tmp_path / "写し"
    shutil.copytree(src, root)
    from saiten_llm.engine import session_path
    original = session_path(Project(src))
    before = original.read_bytes()
    assert main(["render-settings", str(root), "--set", "descriptive_show_score=false"]) == 1
    assert main(["total-region", str(root), "--set", "1,1,50,50"]) == 1
    assert original.read_bytes() == before


def test_render_preview_refuses_other_modes(recognized_project):
    assert main(["render-preview", str(recognized_project)]) == 1


def test_pick_pages_spreads_totals():
    totals = {f"p{i}": i for i in range(10)}
    assert render.pick_pages(totals, 3) == ["p0", "p4", "p9"] or \
        render.pick_pages(totals, 3) == ["p0", "p5", "p9"]
    assert render.pick_pages(totals, 1) == ["p5"]
    assert render.pick_pages({"a": 1}, 3) == ["a"]


# --- 合計点の枠を先生が決める画面（サーバー） --------------------------------

def _get(url):
    with urllib.request.urlopen(url, timeout=10) as res:
        return res.status, res.read()


def _post(url, body):
    req = urllib.request.Request(url, data=json.dumps(body).encode("utf-8"),
                                 headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=30) as res:
        return json.loads(res.read())


def test_total_region_server_preview_and_approve(b4):
    from saiten_llm.cmd_render import _boxed_size
    from saiten_llm.total_region_server import TotalRegionServer
    root, truth = b4
    project = Project(root)
    bridge, _ = open_bridge(project)
    size = _boxed_size(project, bridge)
    srv = TotalRegionServer(project, bridge, "物理_p001.png", size, timeout_minutes=1)
    thread = threading.Thread(target=srv.serve_until_done, kwargs={"poll": 0.05}, daemon=True)
    thread.start()
    try:
        base = f"http://127.0.0.1:{srv.port}"
        # 合言葉が無いと断る
        with pytest.raises(urllib.error.HTTPError):
            _get(f"{base}/api/state")
        status, body = _get(f"{base}/api/state?t={srv.token}")
        state = json.loads(body)
        assert status == 200 and state["detected"] is not None    # 自動で見つけた欄を示す
        status, body = _get(f"{base}/img?w=600&t={srv.token}")
        assert status == 200 and body[:2] == b"\xff\xd8"
        x1, y1, x2, y2 = truth["score_box"]
        w, h = size
        rect = [x1 / w, y1 / h, x2 / w, y2 / h]
        res = _post(f"{base}/api/preview?t={srv.token}", {"rect": rect, "lower_half": True})
        assert res["ok"] and res["image"].startswith("data:image/jpeg;base64,")
        assert _close(res["region"], [x1, (y1 + y2) // 2, x2, y2])
        res = _post(f"{base}/api/approve?t={srv.token}", {"rect": rect, "lower_half": True})
        assert res["ok"]
    finally:
        thread.join(5)
    assert srv.approved.is_set()
    region = open_bridge(Project(root), write=False)[0].get_total_display_region()["region"]
    assert _close(region, [x1, (y1 + y2) // 2, x2, y2])
