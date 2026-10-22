"""allocate.py — 配点を考えるための統計と、配点案（数値計算だけ。ファイルは扱わない）。

配点は先生が個別に判断する。ここは材料を出すだけで、決まった規則で配点を決めない。

項目ごとの値は 0〜1:
  マーク: 正答 1 / 誤答 0
  記述: 得点 ÷ 満点（確認中の仮の2点満点なら 正 1 / 部分 0.5 / 誤 0）
"""
import numpy as np


def item_stats(matrix):
    """項目ごとの正答率・識別力。

    matrix: 生徒×項目（0〜1）
    Returns: [{p, r, ul}]
      p  : 正答率（平均）
      r  : 識別力（その項目を除いた合計点との相関。点双列相関の修正版）
      ul : 上位27%と下位27%の正答率の差
    """
    x = np.asarray(matrix, dtype=float)
    n, k = x.shape
    total = x.sum(axis=1)
    order = np.argsort(total)
    g = max(1, int(round(n * 0.27)))
    low, high = order[:g], order[-g:]
    out = []
    for j in range(k):
        item = x[:, j]
        rest = total - item
        if item.std() == 0 or rest.std() == 0:
            r = None
        else:
            r = float(np.corrcoef(item, rest)[0, 1])
        out.append({"p": float(item.mean()), "r": r,
                    "ul": float(item[high].mean() - item[low].mean())})
    return out


def cronbach_alpha(weighted):
    x = np.asarray(weighted, dtype=float)
    k = x.shape[1]
    if k < 2:
        return None
    total_var = x.sum(axis=1).var(ddof=1)
    if total_var == 0:
        return None
    return float(k / (k - 1) * (1 - x.var(axis=0, ddof=1).sum() / total_var))


def integer_points(weights, total):
    """重みに比例した整数の配点（各項目1点以上・合計は total ちょうど）"""
    w = np.asarray(weights, dtype=float)
    k = len(w)
    if total < k:
        raise ValueError(f"満点 {total} 点では {k} 項目に1点ずつも配れません")
    raw = (total - k) * w / w.sum()
    pts = np.floor(raw).astype(int)
    rest = total - k - pts.sum()
    for j in np.argsort(-(raw - pts))[:rest]:
        pts[j] += 1
    return [int(v) + 1 for v in pts]


def plan_stats(matrix, points):
    x = np.asarray(matrix, dtype=float)
    pts = np.asarray(points, dtype=float)
    weighted = x * pts
    totals = weighted.sum(axis=1)
    full = pts.sum()
    bins = [0] * 10
    for t in totals:
        bins[min(9, int(t / full * 10))] += 1
    return {"mean": float(totals.mean()), "sd": float(totals.std(ddof=1)) if len(totals) > 1 else 0.0,
            "alpha": cronbach_alpha(weighted), "min": float(totals.min()),
            "max": float(totals.max()), "histogram_10": bins}


def plans(matrix, stats, total):
    """配点案（均等・難しい問題を重く・識別力の高い問題を重く）"""
    equal = [1.0] * len(stats)
    difficulty = [1.0 + (1.0 - s["p"]) for s in stats]
    discrimination = [1.0 + 2.0 * max(0.0, s["r"] or 0.0) for s in stats]
    out = []
    for key, label, w in (("equal", "全問同じ重み", equal),
                          ("difficulty", "難しい問題を重く", difficulty),
                          ("discrimination", "識別力の高い問題を重く", discrimination)):
        pts = integer_points(w, total)
        out.append({"key": key, "label": label, "points": pts, **plan_stats(matrix, pts)})
    return out
