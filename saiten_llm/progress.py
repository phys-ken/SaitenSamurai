"""progress.py — 試験の進み具合（進捗.md）を project.json などから書き出す。

「メモ」の見出しより下は先生とエージェントが自由に書く欄なので、書き直しても残す。
"""
from saiten_llm import judgments as jmod
from saiten_llm.mask import mask_status

PROGRESS_FILE = "進捗.md"
MEMO_HEADING = "## メモ"


def _box(done):
    return "[x]" if done else "[ ]"


def _question_names(project):
    """記述問題の名前（採点侍の descriptive_config.json から読むだけ）。

    judgments.json には判定を書き込んだ問題の名前しか無いので、判定前の問題名はこちらから取る
    """
    from constants import RESULTS_DATA_FOLDER, RESULTS_FOLDER
    from descriptive_scorer import DESCRIPTIVE_CONFIG_FILE, load_descriptive_config
    path = project.scans / RESULTS_FOLDER / RESULTS_DATA_FOLDER / DESCRIPTIVE_CONFIG_FILE
    config = load_descriptive_config(str(path)) if path.exists() else None
    return {q.get("id"): q.get("name", "") for q in (config or {}).get("questions", [])}


def checklist(project):
    d = project.data
    lines = []
    pdf = d.get("pdf") or {}
    lines.append(f"- {_box(pdf)} PDF を画像にする" + (f"（{pdf.get('pages')} 枚）" if pdf else ""))
    orient = d.get("orientation") or {}
    lines.append(f"- {_box(orient.get('applied'))} 向きを揃える"
                 + (f"（要確認 {len(orient.get('needs_review', []))} 枚）" if orient.get("needs_review") else ""))
    lines.append(f"- {_box(mask_status(project) == 'approved')} 黒塗りの範囲を先生が承認")
    mode = d.get("mode")
    lines.append(f"- {_box(mode)} モードを先生が承認" + (f"（{mode['app_mode']} / {mode['mark_format']}）" if mode else ""))
    app_mode = (mode or {}).get("app_mode")
    if app_mode in (None, "mark_only", "mark_and_descriptive"):
        rec = d.get("recognition") or {}
        lines += [
            f"- {_box(d.get('coord'))} 座標ファイル" + (f"（{d['coord']['file']}）" if d.get("coord") else ""),
            f"- {_box(rec)} マークの読み取り"
            + (f"（確認の要る答え: {sum(rec.get('warnings', {}).values())} 件）" if rec else ""),
            f"- {_box((d.get('marks_check') or {}).get('teacher_confirmed_at'))} マークの確認を先生が確定",
            f"- {_box((d.get('answer_key') or {}).get('teacher_confirmed_at'))} 正答を先生が確定",
        ]
    if app_mode in (None, "mark_and_descriptive", "descriptive_only"):
        desc = d.get("descriptive") or {}
        lines.append(f"- {_box(desc.get('teacher_approved_at'))} 記述問題の範囲を先生が承認")
        data = jmod.load(project) if jmod.path(project).exists() else {"questions": {}}
        names = _question_names(project) if desc.get("questions") else {}
        for qid in desc.get("questions", []):
            q = data["questions"].get(qid) or {}
            items = q.get("items") or {}
            holds = sum(1 for i in items.values() if i["status"] == "hold")
            state = ("先生が確定" if q.get("confirmed") else
                     f"LLM 判定 {len(items)} 人・保留 {holds}" if items else "未判定")
            name = names.get(qid) or q.get("name", "")
            lines.append(f"  - {_box(q.get('confirmed'))} {qid}（{name}）: {state}")
    points = d.get("points")
    lines.append(f"- {_box(points)} 配点を先生が決定")
    lines.append(f"- {_box(d.get('finalized'))} 個票・集計を作成")
    return lines


def write(project):
    path = project.root / PROGRESS_FILE
    memo = f"{MEMO_HEADING}\n\n（先生と決めたこと・気をつけることを書き足す）\n"
    if path.exists():
        text = path.read_text(encoding="utf-8")
        if MEMO_HEADING in text:
            memo = text[text.index(MEMO_HEADING):]
    body = [f"# 進捗 — {project.root.name}", "",
            "このファイルの上半分は `status --write` が書き直します。", "",
            "## 進み具合", "", *checklist(project), "", memo.rstrip("\n"), ""]
    path.write_text("\n".join(body), encoding="utf-8")
    return path
