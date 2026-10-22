"""marksheet.py — マークシートかどうかの手がかりと、座標ファイルの照合（すべてルール）。

モード（マーク／マーク＋記述／記述のみ）は LLM が黒塗りした1枚目を見て
根拠付きで提案し、先生が承認して決める。ここはその材料になる事実を集める。
"""
import cv2

from saiten_llm.mask import ALIGNED_SCALE

# 楕円のマーク欄らしい形（補正後 1190x1684 の画素）
_BUBBLE_W = (6, 34)
_BUBBLE_H = (12, 56)
_BUBBLE_RATIO = (1.2, 3.8)


def count_bubbles(aligned):
    """補正済みの用紙にある「縦長の楕円」の数（マーク欄の格子の手がかり）。

    点線の楕円は輪郭が途切れて数え漏れるため、実数より少なく出る。
    数十個以上あればマーク欄がある、ほぼ0なら無い、という目安にだけ使う
    """
    gray = cv2.cvtColor(aligned, cv2.COLOR_BGR2GRAY) if aligned.ndim == 3 else aligned
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    contours, _ = cv2.findContours(binary, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    count = 0
    for c in contours:
        x, y, w, h = cv2.boundingRect(c)
        if not (_BUBBLE_W[0] <= w <= _BUBBLE_W[1] and _BUBBLE_H[0] <= h <= _BUBBLE_H[1]):
            continue
        if not (_BUBBLE_RATIO[0] <= h / w <= _BUBBLE_RATIO[1]):
            continue
        # 楕円は外接矩形の 6〜9 割ほどを占める（文字や罫線の切れ端を除く）
        area = cv2.contourArea(c)
        if area < 0.5 * w * h:
            continue
        count += 1
    return count


def _cell_box(c, scale, pad=2):
    x0 = int((c["x"] - pad) * scale)
    y0 = int((c["y"] - pad) * scale)
    x1 = int((c["x"] + c["width"] + pad) * scale)
    y1 = int((c["y"] + c["height"] + pad) * scale)
    return x0, y0, x1, y1


def single_mark_ratio(entries):
    """読み取り結果のうち「1問にマークが1つだけ」だった答えの割合。

    座標ファイルが用紙と合っていないと、ほとんどがノーマークか複数マークになる。
    （インクの量で座標の当否を見る方法は、罫線の上にずれた座標の方が
    インクが多くなり区別できなかったので使わない）
    """
    if not entries:
        return 0.0
    bad = {"ノーマーク", "複数マーク", "不正な値"}
    return sum(1 for e in entries if e["category"] not in bad) / len(entries)


def draw_coord_overlay(image, coords, scale=ALIGNED_SCALE):
    """座標ファイルのマーク位置を枠で描いた写しを返す（LLM が目で照合する用）"""
    out = image.copy()
    groups = {}
    for c in coords:
        x0, y0, x1, y1 = _cell_box(c, scale, pad=0)
        cv2.rectangle(out, (x0, y0), (x1, y1), (0, 160, 0), 1)
        g = groups.setdefault(c["question_no"], [x0, y0, x1, y1])
        g[0], g[1] = min(g[0], x0), min(g[1], y0)
        g[2], g[3] = max(g[2], x1), max(g[3], y1)
    for q, (x0, y0, x1, y1) in groups.items():
        cv2.rectangle(out, (x0 - 3, y0 - 3), (x1 + 3, y1 + 3), (0, 0, 220), 1)
        cv2.putText(out, str(q), (max(0, x0 - 34), y1), cv2.FONT_HERSHEY_SIMPLEX,
                    0.45, (0, 0, 220), 1, cv2.LINE_AA)
    return out


def summarize_coords(coords, skip):
    """座標ファイルの要約: 行（設問）数と、行ごとの選択肢数"""
    per_q = {}
    for c in coords:
        per_q[c["question_no"]] = per_q.get(c["question_no"], 0) + 1
    answer_rows = {q: n for q, n in per_q.items() if q > skip}
    choices = sorted(set(answer_rows.values()))
    return {"rows_total": len(per_q), "id_rows": len(per_q) - len(answer_rows),
            "answer_rows": len(answer_rows), "choices_per_row": choices}
