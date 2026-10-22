"""cmd_finalize.py — 個票・集計を作る前の確認と、作成（S6）。

条件が一つでも欠けていれば作らない（生徒に返すものなので、確認の済んでいない
点数を渡さない）。試しに見たいときは stats を使う。
"""
from pathlib import Path

from saiten_llm import judgments as jmod
from saiten_llm.engine import open_bridge, run_job
from saiten_llm.workspace import now_iso


def _say(*lines):
    for line in lines:
        print(line, flush=True)


def gate(project, bridge):
    """作ってよいかを調べる。Returns: 足りないものの一覧（空なら作ってよい）"""
    problems = []
    mode = bridge.state["app_mode"]
    files = bridge.list_sheet_files()["files"] if bridge.list_sheet_files()["ok"] else []

    if mode != "mark_only":
        data = jmod.load(project)
        questions = (bridge.state.get("descriptive") or {}).get("questions") or []
        if not questions:
            problems.append("記述問題が設定されていません")
        for q in questions:
            label = f"記述 {q['id']}（{q['name']}）"
            if q["max_score"] == jmod.PROVISIONAL_MAX and not (project.data.get("points") or {}).get(
                    "descriptive", {}).get(q["id"]):
                problems.append(f"{label}: 配点が決まっていません（仮の2点満点のまま）")
                continue
            scores = {f: (bridge._desc_scores["scores"].get(f) or {}).get(q["id"]) for f in files}
            unscored = [f for f, v in scores.items() if v is None]
            if unscored:
                problems.append(f"{label}: 未採点が {len(unscored)} 人分あります（部分点の入力待ちを含む）")
                continue
            confirmed = (data["questions"].get(q["id"]) or {}).get("confirmed")
            if not confirmed:
                problems.append(f"{label}: 先生の確定がありません")
            elif confirmed.get("max_score") != q["max_score"]:
                problems.append(f"{label}: 確定の後で配点が変わりました（確定し直しが必要）")
            elif confirmed.get("hash") != jmod.scores_hash(scores):
                problems.append(f"{label}: 確定の後で得点が書き換わりました（確定し直しが必要）")

    if mode != "descriptive_only":
        ks = bridge.state.get("key_summary") or {}
        key_info = project.data.get("answer_key") or {}
        if not ks.get("ok"):
            problems.append("正答チェックが通っていません")
        if not key_info.get("teacher_confirmed_at"):
            problems.append("正答を先生が確定していません")
        elif bridge.state.get("answer_key") and key_info.get("hash") != jmod.file_hash(
                Path(bridge.state["answer_key"])):
            problems.append("確定の後で answer_key.xlsx が書き換わりました（確定し直しが必要）")
        if not (project.data.get("points") or {}).get("marks") and not key_info.get("points_set_at"):
            problems.append("マークの配点を先生が決めていません（set-points）")
        marks = project.data.get("marks_check") or {}
        if not marks.get("teacher_confirmed_at"):
            problems.append("マークの確認（ノーマーク・複数マークの目視）を先生が確定していません")
        elif bridge.state.get("omr_result"):
            omr = Path(bridge.state["omr_result"])
            if omr.name != marks.get("omr_result") or marks.get("hash") != jmod.file_hash(omr):
                problems.append("確定の後で読み取り結果が変わりました（マークの確認をし直す必要あり）")
    return problems


def cmd_finalize(project, args):
    bridge, adapter = open_bridge(project)
    problems = gate(project, bridge)
    if problems:
        _say("個票・集計はまだ作れません。残っていること:")
        for p in problems:
            _say(f"  ・{p}")
        _say("試しに全体の様子を見たいときは stats を使ってください（個票は作りません）。")
        return 2
    _say("確認はすべてそろっています。採点済みの答案（個票）を作っています…")
    done = run_job(bridge, adapter, "run_scoring")
    _say(done["message"], "集計を作っています…")
    done2 = run_job(bridge, adapter, "run_summary")
    _say(done2["message"])
    project.data["finalized"] = {"at": now_iso(), "scored": done["message"],
                                 "summary": done2.get("output_folder")}
    project.save()
    return 0
