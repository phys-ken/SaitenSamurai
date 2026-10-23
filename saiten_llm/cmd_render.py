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
        _say(f"合計点の枠を設定しました: {_region_text(region)}（答案 {size[0]}x{size[1]}）",
             "見た目は render-preview で確かめてください。")
        return 0

    _say(f"答案の大きさ: {size[0]}x{size[1]}（枠はこの画素で書きます）",
         f"いまの合計点の枠: {_region_text(current) if current else '未設定（既定の位置に描きます）'}")
    if not args.detect:
        return 0

    factor = view_factor(project, bridge)
    first = _files(bridge)[0]
    view = render.view_page(project, first)
    found = render.detect_total_box(view)
    if found is None:
        _say(f"「得点」欄を見つけられませんでした（{first}）。",
             "先生に total-region-server で欄を囲んでもらうか、--set で枠を書いてください。")
        return 2
    box = [int(round(v / factor)) for v in found]
    half = list(render.lower_half(box))
    out = project.llm_view / "total_region.jpg"
    render.save_jpeg(out, render.draw_box_overlay(
        view, found, [int(round(v * factor)) for v in half]))
    _say(f"見つけた「得点」欄: {_region_text(box)}",
         f"合計点を書く枠（欄の下半分。見出しに重ならないように）: {_region_text(half)}",
         f"重ねた画像（緑=欄・赤=合計点の枠）: {out}")
    if args.apply:
        call(bridge, "set_total_display_region", half)
        call(bridge, "save_session")
        _say("下半分の枠を、合計点の枠に設定しました。見た目は render-preview で確かめてください。")
    else:
        _say(f"設定するときは: total-region <試験> --set {_region_text(half)}（または --detect --apply）")
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
