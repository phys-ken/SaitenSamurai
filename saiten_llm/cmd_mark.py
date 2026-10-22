"""cmd_mark.py — モード判定の材料・座標ファイル・読み取り・正答（S3）。

マークの読み取りはすべて採点侍のルールベースの処理で行う。LLM は
・黒塗りした1枚目と facts の数字からモードを提案する
・座標の枠を重ねた写しを見て、座標ファイルの取り違えに気づく
・先生から受け取った正答を answer_key.xlsx に書き写す
だけを受け持つ。
"""
import json
import shutil
from datetime import datetime
from pathlib import Path

from saiten_llm.engine import call, open_bridge, run_job
from saiten_llm.workspace import now_iso

MODES = {
    "mark_only": "マーク採点",
    "mark_and_descriptive": "マーク採点＋記述採点",
    "descriptive_only": "記述採点",
}
FORMATS = {"standard": "標準", "multi_digit": "数学（複数桁）"}
_BAD = ("ノーマーク", "複数マーク", "不正な値")


def _say(*lines):
    for line in lines:
        print(line, flush=True)


def templates_dir(project, given=None):
    """座標ファイルを控えておく場所（作業フォルダの templates/）"""
    if given:
        return Path(given)
    candidate = project.root.parent.parent / "templates"
    return candidate if candidate.is_dir() else None


# --- facts --------------------------------------------------------------

def cmd_facts(project, args):
    from saiten_llm.mask import load_aligned
    from saiten_llm.marksheet import count_bubbles
    files = project.scan_files()
    if not files:
        _say("答案の画像がありません。先に pdf を実行してください。")
        return 1
    orient = project.data.get("orientation") or {}
    aligned, has_markers = load_aligned(files[0])
    bubbles = count_bubbles(aligned) if has_markers else None
    view_pages = sorted((project.llm_view / "pages").glob("*.jpg"))
    _say(f"答案: {len(files)} 枚",
         f"四隅のマーカー: {orient.get('with_markers', '未確認')} / {len(files)} 枚"
         if orient.get("with_markers") is not None else
         f"四隅のマーカー（1枚目）: {'あり' if has_markers else 'なし'}",
         f"マーク欄らしき楕円（1枚目）: "
         + (f"約 {bubbles} 個（点線の楕円は数え漏れるので目安）" if bubbles is not None
            else "マーカーが無いので数えていません"),
         f"黒塗りした1枚目: {view_pages[0] if view_pages else 'まだありません（mask-server の後に作られます）'}")
    tdir = templates_dir(project, getattr(args, "templates", None))
    if tdir:
        coords = sorted(tdir.glob("*.xlsx"))
        _say(f"控えてある座標ファイル（{tdir}）: "
             + (", ".join(p.name for p in coords) if coords else "なし"))
    mode = project.data.get("mode")
    _say(f"モード: {MODES[mode['app_mode']] + '（' + FORMATS[mode['mark_format']] + '）・先生が承認済み' if mode else '未決定'}")
    _say("マーカーが無い用紙はマーク採点できません（記述採点のみ）。"
         "マーカーがあってもマーク欄が無ければ記述採点です。")
    return 0


# --- mode ---------------------------------------------------------------

def cmd_mode(project, args):
    if not args.teacher_approved:
        _say("モードは先生の承認を得てから決めます。提案を先生に伝え、"
             "承認されたら --teacher-approved を付けて実行してください。")
        return 2
    bridge, _ = open_bridge(project)
    call(bridge, "set_mode", args.set, args.format)
    if args.skip is not None:
        call(bridge, "set_skip_questions", args.skip)
    call(bridge, "save_session")
    project.data["mode"] = {"app_mode": args.set, "mark_format": args.format,
                            "skip_questions": bridge.state["skip_questions"],
                            "approved_at": now_iso()}
    project.save()
    _say(f"モードを「{MODES[args.set]}（{FORMATS[args.format]}）」にしました。",
         f"ID欄（学年・クラス・番号など）の行数: {bridge.state['skip_questions']}")
    return 0


# --- coord --------------------------------------------------------------

def cmd_coord(project, args):
    import cv2
    from omr_engine import parse_excel_coordinates
    from saiten_llm.marksheet import draw_coord_overlay, summarize_coords
    src = Path(args.file).resolve()
    if not src.exists():
        _say(f"座標ファイルが見つかりません: {src}")
        return 1
    if not project.data.get("mode"):
        _say("先にモードを決めてください（mode --set ... --teacher-approved）。")
        return 1
    # 採点結果フォルダの中に写しを置く。session_state.json は答案画像フォルダからの
    # 相対パスで記録するので、試験フォルダごと移しても辿れる
    from constants import RESULTS_FOLDER, RESULTS_DATA_FOLDER
    data_dir = project.scans / RESULTS_FOLDER / RESULTS_DATA_FOLDER
    data_dir.mkdir(parents=True, exist_ok=True)
    dst = data_dir / src.name
    if src != dst:
        shutil.copy2(src, dst)
    tdir = templates_dir(project, args.templates)
    if tdir and not (tdir / src.name).exists():
        shutil.copy2(src, tdir / src.name)
        _say(f"座標ファイルを控えました: {tdir / src.name}")

    bridge, adapter = open_bridge(project)
    if args.skip is not None:
        call(bridge, "set_skip_questions", args.skip)
    adapter.file_returns = [str(dst)]
    call(bridge, "select_coord_file")
    call(bridge, "save_session")
    skip = bridge.state["skip_questions"]
    coords, _ = parse_excel_coordinates(str(dst))
    summary = summarize_coords(coords, skip)
    cs = bridge.state["coord_summary"] or {}
    _say(f"座標ファイル: {dst.name}",
         f"  行の数: 全 {summary['rows_total']} 行 = ID欄 {summary['id_rows']} 行 + 解答 {summary['answer_rows']} 行",
         f"  1行の選択肢の数: {', '.join(map(str, summary['choices_per_row']))}")
    if cs.get("warning"):
        _say(f"  ⚠ {cs['warning']}")

    view = sorted((project.llm_view / "pages").glob("*.jpg"))
    if view:
        overlay = draw_coord_overlay(cv2.imread(str(view[0])), coords)
        out = project.llm_view / "coord_check.jpg"
        cv2.imwrite(str(out), overlay, [cv2.IMWRITE_JPEG_QUALITY, 88])
        _say(f"座標の枠を黒塗りした1枚目に重ねました: {out}",
             "  緑の枠が各マーク欄に、赤い枠が各設問の行に重なっているかを見てください。"
             f"赤い番号は座標ファイルの行番号で、ID欄も数えるため解答番号より {skip} 大きくなります"
             f"（行 {skip + 1} が問1）。")
    project.data["coord"] = {"file": dst.name, "skip_questions": skip, "set_at": now_iso(),
                             "answer_rows": summary["answer_rows"],
                             "choices_per_row": summary["choices_per_row"]}
    project.save()
    return 0


# --- recognize ----------------------------------------------------------

def _entries(bridge):
    from mark_checker import detect_all_entries_checker
    df = detect_all_entries_checker(bridge.state["omr_result"],
                                    mark_format=bridge.state["mark_format"])
    return df.to_dict("records")


def cmd_recognize(project, args):
    from saiten_llm.marksheet import single_mark_ratio
    bridge, adapter = open_bridge(project)
    if not bridge.state["coord_file"]:
        _say("先に coord で座標ファイルを指定してください。")
        return 1
    _say("答案を読み取っています…")
    done = run_job(bridge, adapter, "run_recognition")
    _say(done["message"])
    entries = _entries(bridge)
    ratio = single_mark_ratio(entries)
    counts = {}
    for e in entries:
        if e["category"] in _BAD:
            counts[e["category"]] = counts.get(e["category"], 0) + 1
    _say(f"1問にマークが1つだけの答え: {ratio:.0%}（{len(entries)} 件中）")
    if ratio < 0.6:
        _say("  ⚠ 少なすぎます。座標ファイルが用紙と合っていない可能性が高いので、"
             "coord_check.jpg を見直し、先生に確認してください。")
    for cat in _BAD:
        if counts.get(cat):
            examples = [f"{e['filename']} 問{int(e['question_no'])}"
                        for e in entries if e["category"] == cat][:8]
            _say(f"  {cat}: {counts[cat]} 件（例: {', '.join(examples)}）")
    if not counts:
        _say("  ノーマーク・複数マーク・不正な値はありません。")
    else:
        _say("これらは先生が採点侍の「マークの確認」で目視します（LLM は判断しません）。")
    project.data["recognition"] = {"done_at": now_iso(), "omr_result": Path(bridge.state["omr_result"]).name,
                                   "success": done.get("success_count"), "errors": done.get("error_count"),
                                   "single_mark_ratio": round(ratio, 3), "warnings": counts}
    project.save()
    return 2 if counts or ratio < 0.6 or done.get("error_count") else 0


# --- answer-key ---------------------------------------------------------

def _results_key_path(bridge):
    from constants import RESULTS_FOLDER, RESULTS_DATA_FOLDER, ANSWER_KEY_FILE
    return (Path(bridge.state["image_folder"]) / RESULTS_FOLDER / RESULTS_DATA_FOLDER
            / ANSWER_KEY_FILE)


def _backup(path):
    if path.exists():
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        backup = path.with_name(f"{path.stem}.backup-{stamp}{path.suffix}")
        shutil.copy2(path, backup)
        return backup
    return None


def _rows_from_json(path):
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    rows = data.get("rows", data) if isinstance(data, dict) else data
    if not isinstance(rows, list) or not rows:
        raise ValueError("JSON は {\"rows\": [{\"問題番号\": 1, \"正答\": \"3\"}, ...]} の形にしてください")
    out = []
    for r in rows:
        if "問題番号" not in r or "正答" not in r:
            raise ValueError(f"問題番号と正答は必須です: {r}")
        out.append({"問題番号": r["問題番号"], "正答": str(r["正答"]),
                    # 配点は仮に1点（最後の配点の検討で決める）
                    "配点": r.get("配点", 1), "観点": r.get("観点", 1),
                    "特例": r.get("特例", ""), "問題概要": r.get("問題概要", "")})
    return out


def _rows_from_sheet(bridge, filename):
    from scoring_engine import load_mark2_results
    students = load_mark2_results(bridge.state["omr_result"],
                                  skip_questions=bridge.state["skip_questions"])
    target = next((s for s in students if str(s.get("image")) == filename), None)
    if target is None:
        raise ValueError(f"読み取り結果に {filename} がありません（先に recognize を実行してください）")
    rows = []
    for q in sorted(target["answers"]):
        value = target["answers"][q]
        if value in (None, "", -1) or str(value).strip() in ("", "-1", "nan"):
            continue
        rows.append({"問題番号": q, "正答": str(value), "配点": 1, "観点": 1,
                     "特例": "", "問題概要": ""})
    if not rows:
        raise ValueError(f"{filename} からは正答が1つも読み取れませんでした")
    return rows


def cmd_answer_key(project, args):
    import pandas as pd
    from omr_engine import _write_styled_template
    bridge, adapter = open_bridge(project)
    if bridge.state["app_mode"] == "descriptive_only":
        _say("記述採点のみのモードでは、マークの正答は使いません。")
        return 1
    key_path = _results_key_path(bridge)
    key_path.parent.mkdir(parents=True, exist_ok=True)
    if args.from_xlsx:
        src = Path(args.from_xlsx).resolve()
        if not src.is_file():
            _say(f"正答の Excel が見つかりません: {src}")
            return 1
        if src != key_path.resolve():
            backup = _backup(key_path)
            shutil.copy2(src, key_path)
            if backup:
                _say(f"前の answer_key.xlsx を控えました: {backup.name}")
        source = f"先生の Excel（{src.name}）。中身は書き換えていません"
    else:
        rows = (_rows_from_json(args.from_json) if args.from_json
                else _rows_from_sheet(bridge, args.from_sheet))
        backup = _backup(key_path)
        if key_path.exists():
            key_path.unlink()
        _write_styled_template(pd.DataFrame(rows), key_path)
        if backup:
            _say(f"前の answer_key.xlsx を控えました: {backup.name}")
        source = (f"正答をマークした用紙（{args.from_sheet}）の読み取り結果" if args.from_sheet
                  else "先生から受け取った正答（配点は仮に1点）")
    adapter.file_returns = [str(key_path)]
    call(bridge, "select_answer_key")
    call(bridge, "save_session")
    ks = bridge.state["key_summary"] or {}
    _say(f"正答・配点: {key_path}", f"  元: {source}")
    if ks.get("stats_line"):
        _say(f"  {ks['stats_line']}")
    for e in ks.get("errors", []):
        _say(f"  ✗ {e}")
    for w in ks.get("warnings", []):
        _say(f"  ⚠ {w}")
    if ks.get("ok"):
        _say("先生に確かめてもらう一覧:")
        from scoring_engine import load_template
        t = load_template(str(key_path), mark_format=bridge.state["mark_format"])
        qs = sorted(t)
        for i in range(0, len(qs), 10):
            chunk = qs[i:i + 10]
            _say(f"  問{chunk[0]}〜{chunk[-1]}: " + ", ".join(str(t[q]["正答"]) for q in chunk))
    project.data["answer_key"] = {"set_at": now_iso(), "source": source, "ok": bool(ks.get("ok")),
                                  "teacher_confirmed_at": None}
    project.save()
    if args.from_sheet:
        _say(f"正答用の用紙 {args.from_sheet} は生徒ではないので、set-aside で集計から外してください。")
    return 0 if ks.get("ok") else 2


# --- set-aside ----------------------------------------------------------

def cmd_set_aside(project, args):
    from saiten_llm.engine import ensure_screen_closed
    ensure_screen_closed(project)
    src = project.scans / args.filename
    if not src.exists():
        _say(f"答案が見つかりません: {args.filename}")
        return 1
    dest_dir = project.root / "除外"
    dest_dir.mkdir(exist_ok=True)
    shutil.move(str(src), str(dest_dir / src.name))
    view = project.llm_view / "pages" / (src.stem + ".jpg")
    if view.exists():
        view.unlink()
    project.data.setdefault("set_aside", []).append(
        {"file": src.name, "reason": args.reason, "at": now_iso()})
    project.save()
    _say(f"{src.name} を集計から外しました（{dest_dir} に移しました）。",
         "マークを読み取り済みなら、recognize をもう一度実行してください。")
    return 0
