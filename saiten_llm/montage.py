"""montage.py — 同じ問題の切り抜きを番号付きで1枚に並べる（LLM に見せる単位）。

切り抜きは黒塗り済みの写し（_llm_view/pages）から作るので、氏名欄の黒塗りは
そのまま残る。番号は問題ごとの通し番号で、どの答案かは index.json で引く
（画像にはファイル名を書かない）。
"""
import json

import cv2
import numpy as np

from saiten_llm.imageio import read_image, write_image

TILE_MAX_W = 900          # 1枚の切り抜きの最大の幅
SHEET_MAX_W = 1800        # 並べた画像の最大の幅
LABEL_W = 56              # 番号を書く欄の幅
GAP = 8


def crop(view_image, region, factor):
    """補正済み用紙の座標（00_Processing）の範囲を、LLM 用の写しから切り抜く"""
    h, w = view_image.shape[:2]
    x1, y1, x2, y2 = (int(round(v * factor)) for v in region)
    x1, x2 = max(0, x1), min(w, x2)
    y1, y2 = max(0, y1), min(h, y2)
    return view_image[y1:y2, x1:x2]


def _fit(tile):
    h, w = tile.shape[:2]
    if w > TILE_MAX_W:
        tile = cv2.resize(tile, (TILE_MAX_W, int(h * TILE_MAX_W / w)), interpolation=cv2.INTER_AREA)
    return tile


def _labeled(tile, number):
    h, w = tile.shape[:2]
    out = np.full((h + 2, w + LABEL_W + 2, 3), 255, np.uint8)
    out[1:h + 1, LABEL_W + 1:LABEL_W + 1 + w] = tile
    cv2.rectangle(out, (LABEL_W, 0), (LABEL_W + w + 1, h + 1), (120, 120, 120), 1)
    cv2.putText(out, str(number), (6, min(h, 40)), cv2.FONT_HERSHEY_SIMPLEX, 0.9,
                (0, 0, 200), 2, cv2.LINE_AA)
    return out


def compose(tiles):
    """[(番号, 切り抜き)] を列に詰めて1枚にする"""
    labeled = [_labeled(_fit(t), n) for n, t in tiles]
    tw = max(t.shape[1] for t in labeled)
    cols = max(1, (SHEET_MAX_W + GAP) // (tw + GAP))
    rows = [labeled[i:i + cols] for i in range(0, len(labeled), cols)]
    heights = [max(t.shape[0] for t in r) for r in rows]
    out = np.full((sum(heights) + GAP * (len(rows) + 1),
                   cols * (tw + GAP) + GAP, 3), 245, np.uint8)
    y = GAP
    for r, rh in zip(rows, heights):
        x = GAP
        for t in r:
            out[y:y + t.shape[0], x:x + t.shape[1]] = t
            x += tw + GAP
        y += rh + GAP
    return out


def build(question, filenames, view_dir, out_dir, factor, per_sheet):
    """1問分の並べた画像を作る。

    Args:
        question: descriptive_config の問題（id, name, region）
        filenames: 答案のファイル名（採点侍の順）
        view_dir: _llm_view/pages
        out_dir: 書き出し先（_llm_view/crops/<問題ID>）
        factor: 00_Processing の座標 → 写しの画素 の倍率
        per_sheet: 1枚に並べる人数

    Returns:
        (書き出した画像のパスの一覧, {番号: ファイル名})
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    for old in out_dir.glob("*.jpg"):
        old.unlink()
    index = {}
    sheets = []
    batch = []
    missing = []
    for number, name in enumerate(filenames, 1):
        stem = name.rsplit(".", 1)[0]
        view = read_image(view_dir / f"{stem}.jpg")
        if view is None:
            missing.append(name)
            continue
        index[number] = name
        batch.append((number, crop(view, question["region"], factor)))
        if len(batch) == per_sheet:
            sheets.append(batch)
            batch = []
    if batch:
        sheets.append(batch)
    paths = []
    for i, tiles in enumerate(sheets, 1):
        path = out_dir / f"sheet_{i:02d}_{tiles[0][0]:03d}-{tiles[-1][0]:03d}.jpg"
        write_image(path, compose(tiles))
        paths.append(path)
    (out_dir / "index.json").write_text(
        json.dumps({"question": question["id"], "name": question["name"],
                    "per_sheet": per_sheet, "numbers": index, "missing": missing},
                   ensure_ascii=False, indent=1), encoding="utf-8")
    return paths, index, missing
