"""imageio.py — 日本語を含むパスでも画像を読み書きする。

cv2.imread / cv2.imwrite は Windows で日本語のパス（「期末テスト」など）を扱えず、
黙って None / False を返す。先生のフォルダ名はほぼ日本語なので、必ずこちらを使う。
"""
from pathlib import Path

import cv2
import numpy as np


def read_image(path, flags=cv2.IMREAD_COLOR):
    """読めなければ None（cv2.imread と同じ）"""
    try:
        data = np.fromfile(str(path), dtype=np.uint8)
    except OSError:
        return None
    if data.size == 0:
        return None
    return cv2.imdecode(data, flags)


def write_image(path, image, jpeg_quality=88):
    """拡張子の形式で書き出す。失敗したら ValueError"""
    path = Path(path)
    params = [cv2.IMWRITE_JPEG_QUALITY, jpeg_quality] if path.suffix.lower() in (".jpg", ".jpeg") else []
    ok, buf = cv2.imencode(path.suffix or ".png", image, params)
    if not ok:
        raise ValueError(f"画像を書き出せません: {path.name}")
    path.write_bytes(buf.tobytes())
