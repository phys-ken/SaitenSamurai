"""orient.py — 答案の向きをルールで確かめ、揃える。

判定の順番（どれも画像を LLM に見せずに済む方法）:
1. 縦横比: 横長のページは 90°/270° のどちらかに回っている
2. Mark2 のマーカー配置: 上辺に3つ（中央寄りに1つ多い）、下辺に2つ。
   これが逆なら 180° 回っている
3. 1枚目との似かた: マーカーが無い用紙は、基準の答案と縮小画像を比べ、
   よく似ている向きを選ぶ。差が小さいときは「要確認」にする

基準の答案（1枚目）が正しい向きかどうかは、黒塗りの画面で先生が見て確かめる。
"""
import cv2
import numpy as np

from omr_engine import find_marker_blobs, imread_unicode

# 似かたの差がこれより小さければ、向きを決めきれないとして先生に確認を頼む
SIMILARITY_MARGIN = 0.08
# 黒い画素の割合がこれ未満なら白紙とみなす
BLANK_INK_RATIO = 0.003

_ROTATE = {
    90: cv2.ROTATE_90_CLOCKWISE,
    180: cv2.ROTATE_180,
    270: cv2.ROTATE_90_COUNTERCLOCKWISE,
}


def rotate(image, degrees):
    """時計回りに degrees 度回す（0/90/180/270）"""
    degrees %= 360
    return image if degrees == 0 else cv2.rotate(image, _ROTATE[degrees])


def to_gray(image):
    return cv2.cvtColor(image, cv2.COLOR_BGR2GRAY) if image.ndim == 3 else image


def marker_counts(gray):
    """上端・下端の帯（高さの8%）にある正方形マーカーの数"""
    h, w = gray.shape
    band = int(h * 0.08)
    min_side, max_side = max(4, int(w * 0.01)), int(w * 0.05)
    top = find_marker_blobs(gray, 0, 0, w, band, min_side, max_side)
    bottom = find_marker_blobs(gray, 0, h - band, w, band, min_side, max_side)
    return len(top), len(bottom)


def thumbnail(gray, size=(48, 68)):
    """向きを比べるための小さな画像（明るさを揃えて平均0・分散1に）"""
    small = cv2.resize(gray, size, interpolation=cv2.INTER_AREA).astype(np.float32)
    small -= small.mean()
    std = small.std()
    return small / std if std > 0 else small


def similarity(a, b):
    return float((a * b).mean())


def ink_ratio(gray):
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    # 大津の二値化は白紙でも境目を作るので、はっきり暗い画素だけ数える
    return float(((gray < 128) & (binary > 0)).mean())


def judge(image, reference=None):
    """1枚の向きを判定する。

    Args:
        image: BGR 画像（ファイルのまま）
        reference: 基準の答案の thumbnail（無ければ似かたは使わない）

    Returns:
        dict(rotation=時計回りに回す角度, method=marker|similarity|none,
             margin=似かたの差, blank=白紙か, needs_review=要確認か)
    """
    gray = to_gray(image)
    h, w = gray.shape
    candidates = (90, 270) if w > h else (0, 180)
    result = {"rotation": candidates[0], "method": "none", "margin": None,
              "blank": ink_ratio(gray) < BLANK_INK_RATIO, "needs_review": True}
    if result["blank"]:
        return result

    by_marker = []
    for rot in candidates:
        top, bottom = marker_counts(rotate(gray, rot))
        if top >= 3 and bottom >= 2 and top > bottom:
            by_marker.append(rot)
    if len(by_marker) == 1:
        result.update(rotation=by_marker[0], method="marker", needs_review=False)
        return result

    if reference is not None:
        scores = sorted(((similarity(thumbnail(rotate(gray, rot)), reference), rot)
                         for rot in candidates), reverse=True)
        margin = scores[0][0] - scores[1][0]
        result.update(rotation=scores[0][1], method="similarity",
                      margin=round(margin, 3),
                      needs_review=margin < SIMILARITY_MARGIN)
    return result


def check_folder(files):
    """全ページの向きを判定する（ファイルは書き換えない）。

    基準は「マーカーで向きが確定した最初のページ」、無ければ白紙でない最初のページ。

    Returns:
        {"reference": ファイル名, "pages": {ファイル名: judge の結果}}
    """
    images = {}
    reference_name = None
    reference = None
    pages = {}
    for path in files:
        img = imread_unicode(str(path))
        if img is None:
            pages[path.name] = {"error": "画像を読めません"}
            continue
        images[path.name] = img
        res = judge(img)
        pages[path.name] = res
        if reference is None and res["method"] == "marker":
            reference_name = path.name
            reference = thumbnail(rotate(to_gray(img), res["rotation"]))

    if reference is None:
        for name, res in pages.items():
            if name in images and not res.get("blank"):
                reference_name = name
                gray = to_gray(images[name])
                rot = 90 if gray.shape[1] > gray.shape[0] else 0
                reference = thumbnail(rotate(gray, rot))
                res.update(rotation=rot, method="reference", needs_review=False)
                break

    if reference is not None:
        for name, res in pages.items():
            if name in images and res.get("method") == "none" and not res["blank"]:
                pages[name] = judge(images[name], reference)
    return {"reference": reference_name, "pages": pages}


def apply_rotation(path, degrees):
    """画像ファイルを回して上書きする（元の PDF は残っているので戻せる）"""
    img = imread_unicode(str(path))
    if img is None:
        raise ValueError(f"画像を読めません: {path.name}")
    ok, encoded = cv2.imencode(path.suffix or ".png", rotate(img, degrees))
    if not ok:
        raise ValueError(f"画像を書き出せません: {path.name}")
    path.write_bytes(encoded.tobytes())
