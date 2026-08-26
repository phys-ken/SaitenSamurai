"""手書きコメント（webui/docs/handwriting-plan.md）。

筆跡はベクターとして 01_Results/handwriting.json に保存し、
「採点済み答案を生成」の最終段で出力画像へ高解像度で焼き込む。
元画像・補正画像には一切書き込まない（main_src ドライバも無改修）。
"""
from __future__ import annotations

import logging
from pathlib import Path

logger = logging.getLogger(__name__)

HANDWRITING_FILE = "handwriting.json"

# 保存を受け付けるペン色（UI と一致させる）。任意色は受けない
ALLOWED_COLORS = ("#c73e2e", "#1d5fa8", "#2b2825")


def _ok(**data):
    return {"ok": True, **data}


def _err(message):
    return {"ok": False, "error": message}


class HandwritingMixin:
    def _init_handwriting(self):
        self._hw_cache = None   # 読み込み済み handwriting.json（dict）

    def _hw_path(self):
        return self._results_data_folder() / HANDWRITING_FILE

    def _hw_load(self):
        from constants import load_json_safe
        if self._hw_cache is None:
            data = None
            if self.state["image_folder"]:
                data = load_json_safe(self._hw_path(), required_keys=["version"])
            self._hw_cache = data or {"version": 1, "sheets": {}}
        return self._hw_cache

    def _load_handwriting_state(self):
        """image_folder 選択/復元時にキャッシュを読み直す"""
        self._hw_cache = None
        self._hw_load()

    # ----------------------------------------------------------------

    def get_handwriting(self, filename):
        """1枚ぶんの筆跡を返す（無ければ空）"""
        if not self.state["image_folder"]:
            return _err("画像フォルダを選択してください")
        sheet = self._hw_load()["sheets"].get(str(filename))
        return _ok(strokes=(sheet or {}).get("strokes", []))

    def set_handwriting(self, filename, natural_w, natural_h, strokes):
        """1枚ぶんの筆跡を保存する（ストローク確定・undo のたびに呼ばれる）。

        strokes: [{color, width, points: [[x, y, pressure], ...]}, ...]
        座標は 00_Processing 画像の実ピクセル。
        """
        from constants import atomic_json_save
        if not self.state["image_folder"]:
            return _err("画像フォルダを選択してください")
        try:
            natural_w = int(natural_w)
            natural_h = int(natural_h)
            assert natural_w > 0 and natural_h > 0
            clean = []
            for s in strokes:
                color = str(s["color"]).lower()
                if color not in ALLOWED_COLORS:
                    return _err(f"使用できない色です: {color}")
                width = float(s["width"])
                assert 0.2 <= width <= 30
                points = [[float(p[0]), float(p[1]),
                           min(1.0, max(0.0, float(p[2]) if len(p) > 2 else 0.5))]
                          for p in s["points"]]
                assert points
                clean.append({"color": color, "width": width, "points": points})
        except (KeyError, TypeError, ValueError, AssertionError, IndexError):
            return _err("筆跡データの形式が不正です")

        data = self._hw_load()
        if clean:
            data["sheets"][str(filename)] = {
                "w": natural_w, "h": natural_h, "strokes": clean}
        else:
            data["sheets"].pop(str(filename), None)   # 全消去はエントリごと消す
        try:
            self._hw_path().parent.mkdir(parents=True, exist_ok=True)
            atomic_json_save(self._hw_path(), data)
        except Exception as e:
            return _err(f"筆跡の保存に失敗しました: {e}")
        return _ok(stroke_count=len(clean))

    # ----------------------------------------------------------------

    def _bake_handwriting(self, output_folder):
        """採点済み答案画像に筆跡を焼き込む（生成ジョブの最終段から呼ぶ）。

        筆跡座標は 00_Processing 基準なので、出力画像とのサイズ比で拡大する。
        失敗しても生成自体は成功扱い（ログのみ）。
        """
        import numpy as np
        try:
            import cv2
            from PIL import Image, ImageDraw
        except Exception:
            logger.exception("焼き込みに必要なライブラリがありません")
            return 0

        data = self._hw_load()
        sheets = data.get("sheets", {})
        if not sheets:
            return 0
        out = Path(output_folder)
        baked = 0
        for filename, sheet in sheets.items():
            target = out / filename
            if not target.exists():
                # 出力の拡張子が変わる場合（png→jpg）に備えて stem 一致も探す
                cands = list(out.glob(Path(filename).stem + ".*"))
                if not cands:
                    continue
                target = cands[0]
            try:
                img = cv2.imdecode(np.fromfile(str(target), dtype=np.uint8),
                                   cv2.IMREAD_COLOR)
                if img is None:
                    continue
                scale = img.shape[1] / float(sheet["w"])
                pil = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
                draw = ImageDraw.Draw(pil)
                for stroke in sheet["strokes"]:
                    color = stroke["color"].lstrip("#")
                    rgb = tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))
                    base_w = float(stroke["width"]) * scale
                    pts = stroke["points"]
                    if len(pts) == 1:
                        x, y, p = pts[0]
                        r = max(0.5, base_w * (0.5 + p * 0.8) / 2)
                        draw.ellipse([x * scale - r, y * scale - r,
                                      x * scale + r, y * scale + r], fill=rgb)
                        continue
                    for a, b in zip(pts, pts[1:]):
                        # 筆圧は区間の平均で線幅に反映（0.5 が標準の太さ）
                        p = (a[2] + b[2]) / 2
                        w = max(1, round(base_w * (0.5 + p * 0.8)))
                        draw.line([a[0] * scale, a[1] * scale,
                                   b[0] * scale, b[1] * scale],
                                  fill=rgb, width=w)
                        r = w / 2
                        for x, y, _ in (a, b):   # 継ぎ目を丸める
                            draw.ellipse([x * scale - r, y * scale - r,
                                          x * scale + r, y * scale + r],
                                         fill=rgb)
                result = cv2.cvtColor(np.array(pil), cv2.COLOR_RGB2BGR)
                suffix = target.suffix.lower()
                encode_as = ".jpg" if suffix in (".jpg", ".jpeg") else ".png"
                params = ([cv2.IMWRITE_JPEG_QUALITY, 85]
                          if encode_as == ".jpg" else [])
                ok, buf = cv2.imencode(encode_as, result, params)
                if ok:
                    buf.tofile(str(target))
                    baked += 1
            except Exception:
                logger.exception("筆跡の焼き込みに失敗: %s", filename)
        if baked:
            logger.info("手書きコメントを %d 枚に焼き込みました", baked)
        return baked
