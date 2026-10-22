"""judgments.py — 記述問題の判定の記録（<試験フォルダ>/judgments.json）。

点数の正本は採点侍の descriptive_scores.json（先生が画面で直すのもこちら）。
このファイルには「LLM がどう判定し、なぜそうしたか」と「先生がいつ確定したか」を残す。

確認中は仮の2点満点で表す: 正=2 / 部分=1 / 誤=0 / 保留=未採点
状態: 未判定（記録なし）/ llm（LLM判定済み・未確認）/ hold（判断保留・先生の確認待ち）
      / 問題ごとの confirmed（先生が確定）
"""
import hashlib
import json
import os
import tempfile

from saiten_llm.workspace import now_iso

JUDGMENTS_FILE = "judgments.json"
PROVISIONAL_MAX = 2
VERDICT_SCORE = {"正": 2, "部分": 1, "誤": 0, "保留": None}


def path(project):
    return project.root / JUDGMENTS_FILE


def load(project):
    p = path(project)
    if p.exists():
        return json.loads(p.read_text(encoding="utf-8"))
    return {"version": 1, "questions": {}}


def save(project, data):
    fd, tmp = tempfile.mkstemp(dir=str(project.root), suffix=".tmp")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path(project))


def scores_hash(scores_by_file):
    """1問分の得点 {ファイル名: 得点} の指紋（確定の後で書き換わったら気づくため）"""
    text = json.dumps(sorted(scores_by_file.items()), ensure_ascii=False)
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def file_hash(file_path):
    return hashlib.sha256(file_path.read_bytes()).hexdigest()


def parse_verdicts(raw, index):
    """LLM の判定（番号→判定）を検査し、{ファイル名: (判定, 理由)} にする。

    raw: {"番号": {"verdict": "正", "reason": "..."}} または {"番号": "正"}
    index: 並べた画像の番号 → ファイル名
    """
    out = {}
    for number, value in raw.items():
        key = str(int(number))
        if key not in index:
            raise ValueError(f"番号 {number} の答案は、並べた画像にありません")
        if isinstance(value, str):
            value = {"verdict": value}
        verdict = value.get("verdict")
        if verdict not in VERDICT_SCORE:
            raise ValueError(f"番号 {number}: 判定は 正 / 部分 / 誤 / 保留 のどれかにしてください（{verdict}）")
        reason = str(value.get("reason", "")).strip()
        if verdict in ("保留", "部分") and not reason:
            raise ValueError(f"番号 {number}: {verdict} には理由を書いてください")
        out[index[key]] = (verdict, reason)
    return out


def record(data, qid, name, verdicts):
    q = data["questions"].setdefault(qid, {"name": name, "items": {}, "confirmed": None})
    q["name"] = name
    at = now_iso()
    for filename, (verdict, reason) in verdicts.items():
        q["items"][filename] = {"status": "hold" if verdict == "保留" else "llm",
                                "verdict": verdict, "reason": reason, "at": at}
    return q
