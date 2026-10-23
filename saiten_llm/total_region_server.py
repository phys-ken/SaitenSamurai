"""total_region_server.py — 合計点を書く枠（「得点」欄）を、先生がブラウザで決める一時サーバー。

total-region --detect で見つからない答案（縦長の用紙など）のための救済。
黒塗りの画面（mask_server）と同じ作りで、自分の PC からしか繋がらず、合言葉付きの URL を使う。

見せるのは黒塗り済みの写し（_llm_view/pages の1枚目）。先生が囲んだ枠で、
採点侍の描画の関数を使って試し描きし、先生が承認したら total_display_config.json に保存する。
"""
import base64
from pathlib import Path

import cv2

from saiten_llm import mask as maskmod
from saiten_llm import render
from saiten_llm.engine import call
from saiten_llm.mask_server import LocalPageServer

PAGE = Path(__file__).with_name("total_region_page.html")


def normalized_to_region(rect, size):
    """[0〜1 の x1, y1, x2, y2] → 採点侍の画素の枠（整数）"""
    x1, y1, x2, y2 = maskmod.normalize_rect(rect)
    w, h = size
    return [int(round(x1 * w)), int(round(y1 * h)), int(round(x2 * w)), int(round(y2 * h))]


class TotalRegionServer(LocalPageServer):
    page_html = PAGE

    def __init__(self, project, bridge, filename, boxed_size, timeout_minutes=60):
        self.project = project
        self.bridge = bridge
        self.filename = filename
        self.boxed_size = tuple(boxed_size)
        view = render.view_page(project, filename)
        # 採点侍の画像と同じ大きさにそろえておく（枠は採点侍の座標で持つ）
        if (view.shape[1], view.shape[0]) != self.boxed_size:
            view = cv2.resize(view, self.boxed_size, interpolation=cv2.INTER_AREA)
        self.view = view
        box = render.detect_total_box(view)
        self.detected = list(box) if box else None
        self.region = None
        super().__init__(timeout_minutes)

    def _normalized(self, region):
        if not region:
            return None
        w, h = self.boxed_size
        return [region[0] / w, region[1] / h, region[2] / w, region[3] / h]

    def state(self):
        current = self.bridge.get_total_display_region().get("region")
        return {"ok": True, "page": self.filename, "size": list(self.boxed_size),
                "detected": self._normalized(self.detected), "current": self._normalized(current)}

    def preview(self, body):
        region = normalized_to_region(body["rect"], self.boxed_size)
        if body.get("lower_half"):
            region = list(render.lower_half(region))
        config = dict(self.bridge._desc_config_with_total_region())
        config["total_display_region"] = region
        scores = self.bridge._desc_scores["scores"].get(self.filename, {})
        image = render.render_preview(self.view, self.boxed_size, config, scores,
                                      self.bridge.state["rendering_settings"])
        image = cv2.rectangle(image, tuple(region[:2]), tuple(region[2:]), (0, 160, 0), 2)
        jpeg = maskmod.encode_jpeg(image, max_width=1400)
        return {"ok": True, "region": region,
                "image": "data:image/jpeg;base64," + base64.b64encode(jpeg).decode("ascii")}

    def approve(self, body):
        region = normalized_to_region(body["rect"], self.boxed_size)
        if body.get("lower_half"):
            region = list(render.lower_half(region))
        call(self.bridge, "set_total_display_region", region)
        call(self.bridge, "save_session")
        self.region = region
        self.approved.set()
        return {"ok": True, "region": region}

    def handle_get(self, path, query):
        if path == "/api/state":
            return 200, self.state(), "json"
        if path == "/img":
            width = max(100, min(2000, int(query.get("w", ["1400"])[0])))
            return 200, maskmod.encode_jpeg(self.view, max_width=width), "image/jpeg"
        return None

    def handle_post(self, path, body):
        if path == "/api/preview":
            return self.preview(body)
        if path == "/api/approve":
            return self.approve(body)
        return None
