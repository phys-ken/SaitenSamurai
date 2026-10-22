"""cmd_desc.py — 記述問題の範囲・並べた画像・判定・確定（S4）。

流れ:
  desc-set  LLM が黒塗りした1枚目を見て問題と範囲の案を書く（仮の2点満点）
            → 先生が採点侍の「記述問題の設定」で微調整・承認
  crops     同じ問題の切り抜きを番号付きで並べた画像を作る（LLM が見る単位）
  judge     LLM の判定（正/部分/誤/保留）を書き込む
            → 先生が採点侍で確認・修正（画面を開いている間は先生の番）
  review    判定の状況・保留・先生の修正を一覧にする
  confirm   先生がチャットで「問3確認しました」と伝えたら、その時点の内容を記録する
"""
import json
from pathlib import Path

import cv2

from saiten_llm import judgments as jmod
from saiten_llm.engine import call, open_bridge, run_job
from saiten_llm.workspace import now_iso


def _say(*lines):
    for line in lines:
        print(line, flush=True)


def _questions(bridge):
    return (bridge.state.get("descriptive") or {}).get("questions") or []


def _files(bridge):
    listing = bridge.list_sheet_files()
    return listing["files"] if listing["ok"] else []


def view_factor(project, bridge):
    """採点侍の座標（00_Processing）→ LLM 用の写しの画素 の倍率"""
    from constants import RESULTS_FOLDER, BOXED_FOLDER
    from PIL import Image
    files = _files(bridge)
    if not files:
        raise RuntimeError("補正済みの答案がありません。先に recognize（記述のみなら prepare）を実行してください")
    boxed = project.scans / RESULTS_FOLDER / BOXED_FOLDER / files[0]
    view = project.llm_view / "pages" / (Path(files[0]).stem + ".jpg")
    if not view.exists():
        raise RuntimeError("LLM 用の写しがありません。先に黒塗りを承認してもらってください")
    with Image.open(boxed) as b, Image.open(view) as v:
        return v.size[0] / b.size[0]


# --- prepare（記述のみ） -------------------------------------------------

def cmd_prepare(project, args):
    bridge, adapter = open_bridge(project)
    if bridge.state["app_mode"] != "descriptive_only":
        _say("マーク採点のモードでは、recognize が答案の準備も行います。")
        return 1
    done = run_job(bridge, adapter, "run_prepare_images")
    _say(done["message"])
    return 2 if done.get("no_marker") else 0


# --- desc-set -----------------------------------------------------------

def _draw_regions(image, questions, factor):
    out = image.copy()
    for q in questions:
        x1, y1, x2, y2 = (int(round(v * factor)) for v in q["region"])
        cv2.rectangle(out, (x1, y1), (x2, y2), (0, 0, 220), 3)
        cv2.putText(out, q["id"], (x1 + 6, y1 + 30), cv2.FONT_HERSHEY_SIMPLEX, 0.9,
                    (0, 0, 220), 2, cv2.LINE_AA)
    return out


def cmd_desc_set(project, args):
    spec = json.loads(Path(args.input).read_text(encoding="utf-8"))
    items = spec.get("questions", spec) if isinstance(spec, dict) else spec
    if not isinstance(items, list) or not items:
        _say('JSON は {"questions": [{"name": "問1", "region": [x1, y1, x2, y2]}]} の形にしてください'
             "（範囲は _llm_view/pages の画像の画素で）")
        return 1
    bridge, _ = open_bridge(project)
    if bridge.state["app_mode"] == "mark_only":
        _say("マーク採点のみのモードです。記述問題を使うなら、先生の承認を得てモードを変えてください。")
        return 1
    factor = view_factor(project, bridge)
    existing = _questions(bridge)
    if existing:
        if not args.replace:
            _say(f"記述問題がすでに {len(existing)} 問あります。作り直すときは --replace を付けてください。")
            return 1
        scored = [q["id"] for q in existing if bridge.state["descriptive"]["scored_counts"].get(q["id"])]
        if scored:
            _say(f"得点が入っている問題があるので作り直せません: {', '.join(scored)}")
            return 1
        for q in list(existing):
            call(bridge, "delete_descriptive_question", q["id"])
    added = []
    for item in items:
        region = [v / factor for v in item["region"]]
        res = call(bridge, "add_descriptive_question", item.get("name", ""),
                   jmod.PROVISIONAL_MAX, item.get("aspect", 1), region)
        added.append(res["question_id"])
    questions = _questions(bridge)
    first = sorted((project.llm_view / "pages").glob("*.jpg"))[0]
    out = project.llm_view / "desc_regions.jpg"
    cv2.imwrite(str(out), _draw_regions(cv2.imread(str(first)), questions, factor),
                [cv2.IMWRITE_JPEG_QUALITY, 88])
    project.data["descriptive"] = {"proposed_at": now_iso(), "questions": added,
                                   "teacher_approved_at": None}
    project.save()
    _say(f"記述問題を {len(added)} 問設定しました（確認中は仮の2点満点）: "
         + ", ".join(f"{q['id']}={q['name']}" for q in questions),
         f"範囲を黒塗りした1枚目に重ねました: {out}",
         "先生に採点侍の「記述問題の設定」で範囲を確かめてもらい、承認をもらってください（open-ui）。")
    return 0


# --- crops --------------------------------------------------------------

def cmd_crops(project, args):
    from saiten_llm import montage
    bridge, _ = open_bridge(project, write=False)
    questions = _questions(bridge)
    if args.question:
        questions = [q for q in questions if q["id"] == args.question]
        if not questions:
            _say(f"問題が見つかりません: {args.question}")
            return 1
    if not questions:
        _say("記述問題がまだありません。先に desc-set を実行してください。")
        return 1
    factor = view_factor(project, bridge)
    files = _files(bridge)
    for q in questions:
        out_dir = project.llm_view / "crops" / q["id"]
        paths, index, missing = montage.build(q, files, project.llm_view / "pages", out_dir,
                                              factor, args.per_sheet)
        _say(f"{q['id']}（{q['name']}）: {len(index)} 人分を {len(paths)} 枚に並べました → {out_dir}")
        for p in paths:
            _say(f"  {p.name}")
        if missing:
            _say(f"  ⚠ LLM 用の写しが無い答案 {len(missing)} 枚: {', '.join(missing[:5])}")
    _say("番号は問題ごとの通し番号です。判定は番号で judge に渡してください。")
    return 0


# --- judge --------------------------------------------------------------

def cmd_judge(project, args):
    bridge, _ = open_bridge(project)
    q = next((q for q in _questions(bridge) if q["id"] == args.question), None)
    if q is None:
        _say(f"問題が見つかりません: {args.question}")
        return 1
    if q["max_score"] != jmod.PROVISIONAL_MAX:
        _say(f"{q['id']} は配点が確定済み（{q['max_score']}点満点）なので、判定は書き込めません。")
        return 1
    data = jmod.load(project)
    record = data["questions"].get(q["id"]) or {}
    if record.get("confirmed"):
        _say(f"{q['id']} は先生が確定済みです（{record['confirmed']['at']}）。判定は書き換えません。")
        return 1
    index_path = project.llm_view / "crops" / q["id"] / "index.json"
    if not index_path.exists():
        _say("並べた画像がありません。先に crops を実行してください。")
        return 1
    index = {str(k): v for k, v in json.loads(index_path.read_text(encoding="utf-8"))["numbers"].items()}
    verdicts = jmod.parse_verdicts(json.loads(Path(args.input).read_text(encoding="utf-8")), index)

    # 先生が画面で直した得点は上書きしない（LLM の前回の判定と今の得点が違うもの）
    current = {f: s.get(q["id"]) for f, s in bridge._desc_scores["scores"].items()}
    kept = []
    for filename in list(verdicts):
        prev = (record.get("items") or {}).get(filename)
        if prev and jmod.VERDICT_SCORE[prev["verdict"]] != current.get(filename):
            kept.append(filename)
            del verdicts[filename]
    call(bridge, "set_descriptive_scores_bulk", q["id"],
         {f: jmod.VERDICT_SCORE[v] for f, (v, _r) in verdicts.items()})
    jmod.record(data, q["id"], q["name"], verdicts)
    jmod.save(project, data)
    counts = {}
    for v, _r in verdicts.values():
        counts[v] = counts.get(v, 0) + 1
    _say(f"{q['id']}（{q['name']}）に {len(verdicts)} 人分の判定を書き込みました: "
         + "、".join(f"{k} {n}" for k, n in counts.items()))
    if kept:
        _say(f"先生が画面で直した {len(kept)} 人分は書き換えていません: {', '.join(kept[:5])}")
    return 0


# --- review -------------------------------------------------------------

def question_report(bridge, data, q):
    """1問分の状況: 件数・保留・先生の修正・未判定"""
    files = _files(bridge)
    record = data["questions"].get(q["id"]) or {"items": {}, "confirmed": None}
    items = record.get("items") or {}
    scores = bridge._desc_scores["scores"]
    report = {"id": q["id"], "name": q["name"], "total": len(files), "unscored": [],
              "holds": [], "partials": [], "changed": [], "unjudged": [], "counts": {},
              "confirmed": record.get("confirmed")}
    for f in files:
        score = (scores.get(f) or {}).get(q["id"])
        item = items.get(f)
        if score is None:
            report["unscored"].append(f)
        else:
            report["counts"][score] = report["counts"].get(score, 0) + 1
        if item is None:
            report["unjudged"].append(f)
            continue
        if item["verdict"] == "保留" and score is None:
            report["holds"].append((f, item["reason"]))
        if item["verdict"] == "部分" and score is None:
            report["partials"].append(f)      # 換算の後、先生の点の入力待ち
        expected = jmod.expected_score(item["verdict"], q["max_score"])
        if expected is not None and expected != score:
            report["changed"].append((f, item["verdict"], score))
    return report


def cmd_review(project, args):
    bridge, _ = open_bridge(project, write=False)
    data = jmod.load(project)
    questions = _questions(bridge)
    if args.question:
        questions = [q for q in questions if q["id"] == args.question]
    label = {2: "正", 1: "部分", 0: "誤"}
    for q in questions:
        r = question_report(bridge, data, q)
        provisional = q["max_score"] == jmod.PROVISIONAL_MAX
        counts = "、".join(f"{label.get(s, s) if provisional else str(s) + '点'} {n}"
                          for s, n in sorted(r["counts"].items(), reverse=True))
        state = (f"確定済み（{r['confirmed']['at']}）" if r["confirmed"] else "未確定")
        _say(f"{q['id']}（{q['name']}）: {state} / {r['total']} 人中 未採点 {len(r['unscored'])}"
             + (f" / {counts}" if counts else ""))
        for f, reason in r["holds"]:
            _say(f"  保留: {f} — {reason}")
        for f in r["partials"]:
            _say(f"  部分点の入力待ち: {f}")
        for f, verdict, score in r["changed"]:
            now = "未採点" if score is None else (label.get(score, score) if provisional else f"{score}点")
            _say(f"  先生の修正: {f}  LLM {verdict} → {now}")
        if r["unjudged"] and not r["confirmed"]:
            _say(f"  LLM がまだ判定していない答案: {len(r['unjudged'])} 人")
    return 0


# --- confirm ------------------------------------------------------------

def cmd_confirm(project, args):
    """先生がチャットで確認を伝えたときに、その時点の内容を記録する"""
    bridge, _ = open_bridge(project)
    if args.what == "question":
        q = next((q for q in _questions(bridge) if q["id"] == args.question), None)
        if q is None:
            _say(f"問題が見つかりません: {args.question}")
            return 1
        data = jmod.load(project)
        r = question_report(bridge, data, q)
        if r["unscored"]:
            _say(f"{q['id']} には未採点（保留を含む）の答案が {len(r['unscored'])} 人分残っているので、"
                 "まだ確定できません。先生に画面で点を入れてもらってください。")
            for f in r["unscored"][:10]:
                _say(f"  {f}")
            return 2
        scores = {f: (bridge._desc_scores["scores"].get(f) or {}).get(q["id"]) for f in _files(bridge)}
        record = data["questions"].setdefault(q["id"], {"name": q["name"], "items": {}, "confirmed": None})
        record["confirmed"] = {"at": now_iso(), "hash": jmod.scores_hash(scores),
                               "max_score": q["max_score"], "counts": r["counts"],
                               "teacher_changes": len(r["changed"])}
        jmod.save(project, data)
        _say(f"{q['id']}（{q['name']}）を先生の確定として記録しました（{len(scores)} 人分）。")
        return 0
    if args.what == "regions":
        project.data.setdefault("descriptive", {})["teacher_approved_at"] = now_iso()
        project.data["descriptive"]["regions"] = {q["id"]: q["region"] for q in _questions(bridge)}
        project.save()
        _say("記述問題の範囲を先生の承認として記録しました。")
        return 0
    if args.what == "answer-key":
        ks = bridge.state.get("key_summary") or {}
        if not ks.get("ok"):
            _say("正答チェックにエラーがあるので確定できません。")
            return 2
        key = Path(bridge.state["answer_key"])
        project.data.setdefault("answer_key", {})["teacher_confirmed_at"] = now_iso()
        project.data["answer_key"]["hash"] = jmod.file_hash(key)
        project.save()
        _say("正答を先生の確定として記録しました。")
        return 0
    if args.what == "marks":
        omr = Path(bridge.state["omr_result"]) if bridge.state.get("omr_result") else None
        if omr is None:
            _say("マークの読み取り結果がありません。")
            return 1
        corrections = omr.parent / "tmp_checking_dm_nm.csv"
        project.data["marks_check"] = {
            "teacher_confirmed_at": now_iso(), "omr_result": omr.name,
            "hash": jmod.file_hash(omr),
            "corrections_hash": jmod.file_hash(corrections) if corrections.exists() else None}
        project.save()
        _say("マークの確認（ノーマーク・複数マークなどの目視）を先生の確定として記録しました。")
        return 0
    _say(f"確定の種類が不明です: {args.what}")
    return 1
