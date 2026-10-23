"""cmd_render.py — 返却答案（個票）の見た目を整えて確かめるコマンド。

  render-settings      描画の設定を見る・変える（観点を出すか など）
  total-region         合計点を書く枠を探す・設定する
  total-region-server  合計点を書く枠を先生がブラウザで決める
  render-preview       黒塗り済みの写しの上に、いまの設定で個票を描いてみる
"""
import shutil
import webbrowser

from saiten_llm import render
from saiten_llm.cmd_desc import _files, view_factor
from saiten_llm.engine import call, open_bridge
from saiten_llm.workspace import now_iso

PREVIEW_DIR = "preview"


def _say(*lines):
    for line in lines:
        print(line, flush=True)


def _boxed_size(project, bridge):
    """採点侍の画像（00_Processing）の大きさ（幅, 高さ）"""
    from constants import BOXED_FOLDER, RESULTS_FOLDER
    from PIL import Image
    files = _files(bridge)
    if not files:
        raise RuntimeError("採点侍の画像がありません。先に prepare（マーク式なら recognize）を実行してください")
    with Image.open(project.scans / RESULTS_FOLDER / BOXED_FOLDER / files[0]) as im:
        return im.size


def _region_text(region):
    return ",".join(str(v) for v in region)


# --- render-settings -----------------------------------------------------

def cmd_render_settings(project, args):
    try:
        changes = dict(render.parse_setting(t) for t in args.set or [])
    except ValueError as e:
        _say(str(e))
        return 1
    bridge, _ = open_bridge(project, write=bool(changes))
    if changes:
        call(bridge, "set_rendering_settings", changes)
        call(bridge, "save_session")
        _say("描画の設定を変えました: " + "、".join(f"{k}={v}" for k, v in changes.items()))
    res = call(bridge, "get_rendering_settings")
    _say("いまの描画の設定（* は既定から変えたもの）:")
    for key, value in res["settings"].items():
        mark = "*" if value != res["defaults"][key] else " "
        _say(f" {mark} {key} = {value}（既定 {res['defaults'][key]}）")
    if changes:
        _say("見た目は render-preview で確かめてください。")
    return 0


# --- total-region --------------------------------------------------------

def _parse_region(text, size):
    try:
        region = [int(round(float(v))) for v in text.split(",")]
    except ValueError:
        raise ValueError(f"枠は x1,y1,x2,y2 の数で書いてください: {text}") from None
    if len(region) != 4:
        raise ValueError(f"枠は x1,y1,x2,y2 の4つの数で書いてください: {text}")
    x1, y1, x2, y2 = region
    w, h = size
    if not (0 <= x1 < x2 <= w and 0 <= y1 < y2 <= h):
        raise ValueError(f"枠が答案（{w}x{h}）の外にはみ出しているか、幅・高さがありません: {text}")
    return region


RED, GREEN = (0, 0, 230), (0, 160, 0)
OVERLAY_FILE = "total_region.jpg"         # 1枚目の全体に枠を重ねた画像
CROP_FILE = "total_region_crop.jpg"       # 枠の周りを拡大し、目盛りを重ねた画像


def _first_page(project, bridge, size):
    """黒塗り済みの1枚目を、採点侍の答案と同じ大きさ（原寸の画素）にしたもの"""
    import cv2
    first = _files(bridge)[0]
    view = render.view_page(project, first)
    if (view.shape[1], view.shape[0]) != tuple(size):
        view = cv2.resize(view, tuple(size), interpolation=cv2.INTER_AREA)
    return first, view


def _write_overlays(project, page, region, detected, crop_rect):
    """全体に枠を重ねた画像と、範囲を拡大して目盛りを重ねた画像を書く"""
    overlay = render.save_jpeg(project.llm_view / OVERLAY_FILE,
                               render.draw_box_overlay(page, detected, region))
    boxes = [(detected, GREEN)] if detected else []
    if region:
        boxes.append((region, RED))
    image, _ticks = render.grid_crop(page, crop_rect, boxes=boxes)
    crop = render.save_jpeg(project.llm_view / CROP_FILE, image)
    return overlay, crop


def cmd_total_region(project, args):
    writing = bool(args.set or args.clear or args.apply)
    bridge, _ = open_bridge(project, write=writing)
    size = _boxed_size(project, bridge)
    current = call(bridge, "get_total_display_region")["region"]

    if args.clear:
        call(bridge, "set_total_display_region", None)
        call(bridge, "save_session")
        _say("合計点の枠を外しました（既定の位置に描きます）。")
        return 0
    if args.set:
        try:
            region = _parse_region(args.set, size)
        except ValueError as e:
            _say(str(e))
            return 1
        call(bridge, "set_total_display_region", region)
        call(bridge, "save_session")
        first, page = _first_page(project, bridge, size)
        overlay, crop = _write_overlays(project, page, region, None, render.around(region, size))
        _say(f"合計点の枠を設定しました: {_region_text(region)}（答案 {size[0]}x{size[1]}）",
             f"枠（赤）を重ねた画像: {crop}（拡大・目盛り付き）／ {overlay}（1枚目の全体。{first}）",
             "枠が欄の中（見出しの下）に乗っているか画像で確かめ、render-preview で合計点が読めるか確かめてください。")
        return 0

    _say(f"答案の大きさ: {size[0]}x{size[1]}（枠はこの画素で書きます）",
         f"いまの合計点の枠: {_region_text(current) if current else '未設定（既定の位置に描きます）'}")

    if args.crop:
        try:
            rect = (render.default_crop_rect(size) if args.crop == "auto"
                    else _parse_region(args.crop, size))
        except ValueError as e:
            _say(str(e))
            return 1
        first, page = _first_page(project, bridge, size)
        image, ticks = render.grid_crop(page, rect, boxes=[(current, RED)] if current else [])
        out = render.save_jpeg(project.llm_view / CROP_FILE, image)
        _say(f"切り出して目盛りを重ねた画像: {out}（{first}・範囲 {_region_text(rect)}）",
             "目盛りの数は原寸の画素です（50 画素ごと。太い線は 100 の倍数）。表示の倍率を掛ける必要はありません。"
             + ("赤の枠は、いまの合計点の枠です。" if current else ""),
             "欄を読んだら --set x1,y1,x2,y2 で設定します（「得点」などの見出しを避けて、欄の下半分にする）。")
        return 0

    if not args.detect:
        return 0

    # 検出はヒント。見つからなくても失敗にしない（枠は目で見て決める）
    factor = view_factor(project, bridge)
    first = _files(bridge)[0]
    view = render.view_page(project, first)
    found = render.detect_total_box(view)
    if found is None:
        _say(f"（ヒント）「得点」欄は見つかりませんでした（{first}）。目で見て決めてください。",
             "total-region <試験> --crop で、右上の隅を目盛り付きで拡大した画像が出ます。")
        return 0
    box = [int(round(v / factor)) for v in found]
    half = list(render.lower_half(box))
    _first, page = _first_page(project, bridge, size)
    overlay, crop = _write_overlays(project, page, half, box, render.around(box, size))
    _say(f"（ヒント）「得点」欄らしい四角: {_region_text(box)}",
         f"その下半分（見出しに重ならないように）: {_region_text(half)}",
         f"重ねた画像（緑=見つけた四角・赤=下半分）: {crop}（拡大・目盛り付き）／ {overlay}",
         "検出は取り違えることがあります。必ず画像で、本当に「得点」欄か確かめてください。")
    if args.apply:
        call(bridge, "set_total_display_region", half)
        call(bridge, "save_session")
        _say("下半分の枠を、合計点の枠に設定しました。見た目は render-preview で確かめてください。")
    else:
        _say(f"合っていれば: total-region <試験> --set {_region_text(half)}")
    return 0


def cmd_total_region_server(project, args):
    from saiten_llm.total_region_server import TotalRegionServer
    bridge, _ = open_bridge(project)
    if bridge.state["app_mode"] != "descriptive_only":
        _say("合計点の枠を画面で決めるのは、いまは記述のみモードだけです。")
        return 1
    size = _boxed_size(project, bridge)
    server = TotalRegionServer(project, bridge, _files(bridge)[0], size,
                               timeout_minutes=args.timeout)
    print("合計点を書く位置を決める画面を用意しました。先生にこのリンクを開いてもらってください:")
    print(f"  {server.url}")
    print(f"（このPCからしか開けません。承認されるか {args.timeout} 分たつと止まります）", flush=True)
    if not args.no_open:
        try:
            webbrowser.open(server.url)
        except Exception:
            pass
    if not server.serve_until_done():
        _say("承認されないまま時間切れになりました。合計点の枠は変えていません。")
        return 2
    _say(f"先生が合計点の枠を承認しました: {_region_text(server.region)}",
         "見た目は render-preview で確かめてください。")
    return 0


# --- render-preview ------------------------------------------------------

def cmd_render_preview(project, args):
    bridge, _ = open_bridge(project, write=False)
    if bridge.state["app_mode"] != "descriptive_only":
        _say("render-preview は、いまは記述のみモードだけに対応しています。",
             "マーク式の個票は、採点侍の画面の「位置プレビュー」で確かめてください。")
        return 1
    files = _files(bridge)
    if not files:
        _say("採点侍の画像がありません。先に prepare を実行してください。")
        return 1
    config = bridge._desc_config_with_total_region()
    if not config.get("questions"):
        _say("記述問題が設定されていません。先に desc-set を実行してください。")
        return 1
    size = _boxed_size(project, bridge)
    scores = bridge._desc_scores["scores"]
    totals = {f: sum(v for v in (scores.get(f) or {}).values() if v is not None) for f in files}
    picked = render.pick_pages(totals, args.pages)

    out_dir = project.llm_view / PREVIEW_DIR
    if out_dir.exists():
        shutil.rmtree(out_dir)
    rs = bridge.state["rendering_settings"]
    written = []
    for i, f in enumerate(picked, 1):
        image = render.render_preview(render.view_page(project, f), size, config,
                                      scores.get(f) or {}, rs)
        path = render.save_jpeg(out_dir / f"preview_{i:02d}_{totals[f]:g}点.jpg", image)
        written.append((path, f))
    unscored = sorted({q["name"] or q["id"] for q in config["questions"] for f in picked
                       if (scores.get(f) or {}).get(q["id"]) is None})
    full_max = sum(q["max_score"] for q in config["questions"])
    _say(f"いまの設定で個票を描いてみました（黒塗り済みの写しの上。{len(written)} 枚）:")
    for path, f in written:
        _say(f"  {path}（{f}・合計 {totals[f]:g} / {full_max} 点）")
    region = config.get("total_display_region")
    _say(f"合計点の枠: {_region_text(region) if region else '未設定（既定の位置。問題文に重なることがあります）'}")
    changed = {k: v for k, v in (rs or {}).items()}
    _say("描画の設定（既定から変えたもの）: "
         + ("、".join(f"{k}={v}" for k, v in changed.items()) if changed else "なし"))
    if unscored:
        _say(f"⚠ 未採点の問題があります（描かれません）: {', '.join(unscored)}")
    if any(q["max_score"] == 2 for q in config["questions"]) and not project.data.get("points"):
        _say("（配点を決める前なので、仮の2点満点で描いています）")
    _say("画像を見て、○×△の大きさ・合計点の位置を確かめ、先生にも見てもらってください。",
         "手書きのコメント（採点侍の画面で書いたもの）は、このプレビューには入りません。")
    project.data["render_preview"] = {"at": now_iso(), "pages": [f for _p, f in written]}
    project.save()
    return 0
