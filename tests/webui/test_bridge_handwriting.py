"""手書きコメント（L1）: 保存往復・焼き込みの画素と倍率・全モードフック。"""
import sys
from pathlib import Path

import numpy as np
import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "webui"))
sys.path.insert(0, str(PROJECT_ROOT / "main_src"))
sys.path.insert(0, str(PROJECT_ROOT / "tests"))
sys.path.insert(0, str(PROJECT_ROOT / "tests" / "webui"))

from api.bridge import Bridge  # noqa: E402
from test_bridge_jobs import RecordingAdapter, _wait_job_done  # noqa: E402
from test_bridge_extras import (  # noqa: E402
    _build_marked_inputs, _bridge_through_recognition,
)

STROKE = {"color": "#c73e2e", "width": 2.0,
          "points": [[100, 100, 0.5], [200, 100, 0.5], [200, 200, 0.9]]}


class TestPersistence:
    def test_roundtrip_and_reload(self, tmp_path):
        scans, coord, key = _build_marked_inputs(tmp_path)
        b, _ = _bridge_through_recognition(scans, coord, key)
        assert b.get_handwriting("s1.png")["strokes"] == []
        res = b.set_handwriting("s1.png", 595, 842, [STROKE])
        assert res["ok"] and res["stroke_count"] == 1

        # 別ブリッジで読み直しても残っている（ファイル保存の確認）
        b2 = Bridge(window_adapter=RecordingAdapter())
        b2.set_mode("mark_only", "multi_digit")
        b2._win.folder_returns = [str(scans)]
        assert b2.select_image_folder()["ok"]
        strokes = b2.get_handwriting("s1.png")["strokes"]
        assert len(strokes) == 1
        assert strokes[0]["points"][2] == [200.0, 200.0, 0.9]

    def test_clear_removes_entry(self, tmp_path):
        scans, coord, key = _build_marked_inputs(tmp_path)
        b, _ = _bridge_through_recognition(scans, coord, key)
        b.set_handwriting("s1.png", 595, 842, [STROKE])
        assert b.set_handwriting("s1.png", 595, 842, [])["ok"]
        from api.handwriting import HANDWRITING_FILE
        import json
        data = json.loads((b._results_data_folder() / HANDWRITING_FILE)
                          .read_text(encoding="utf-8"))
        assert data["sheets"] == {}

    def test_validation(self, tmp_path):
        scans, coord, key = _build_marked_inputs(tmp_path)
        b, _ = _bridge_through_recognition(scans, coord, key)
        bad_color = dict(STROKE, color="#00ff00")
        assert not b.set_handwriting("s1.png", 595, 842, [bad_color])["ok"]
        assert not b.set_handwriting("s1.png", 595, 842,
                                     [{"color": "#c73e2e", "width": 2,
                                       "points": []}])["ok"]
        assert not b.set_handwriting("s1.png", 0, 842, [STROKE])["ok"]


class TestBaking:
    def test_scoring_bakes_strokes_at_scaled_position(self, tmp_path):
        """焼き込み: 出力画像の（スケール換算した）筆跡位置に朱の画素が乗る"""
        import cv2
        scans, coord, key = _build_marked_inputs(tmp_path)
        b, adapter = _bridge_through_recognition(scans, coord, key)
        # 下半分の何もない場所に横線を書く
        b.set_handwriting("s1.png", 595, 842,
                          [{"color": "#c73e2e", "width": 3.0,
                            "points": [[100, 600, 0.5], [400, 600, 0.5]]}])
        assert b.run_scoring()["ok"]
        _wait_job_done(b)
        done = [e for e in adapter.events if e["type"] == "job_done"][-1]
        assert done["ok"] and "手書きコメント 1 枚" in done["message"]

        from constants import RESULTS_FOLDER, SCORED_FOLDER
        img = cv2.imdecode(np.fromfile(
            str(scans / RESULTS_FOLDER / SCORED_FOLDER / "s1.png"),
            dtype=np.uint8), cv2.IMREAD_COLOR)
        scale = img.shape[1] / 595.0
        y = int(600 * scale)
        band = img[y - 6:y + 6, int(120 * scale):int(380 * scale)]
        b_, g, r = (band[:, :, 0].astype(int), band[:, :, 1].astype(int),
                    band[:, :, 2].astype(int))
        reddish = int(((r > 120) & (r - g > 40) & (r - b_ > 40)).sum())
        assert reddish > 50, "スケール換算した位置に朱の筆跡が無い"
        # 書いていない答案は焼き込み対象外
        img2 = cv2.imdecode(np.fromfile(
            str(scans / RESULTS_FOLDER / SCORED_FOLDER / "s2.png"),
            dtype=np.uint8), cv2.IMREAD_COLOR)
        y2 = int(600 * (img2.shape[1] / 595.0))
        band2 = img2[y2 - 6:y2 + 6, :]
        b2_, g2, r2 = (band2[:, :, 0].astype(int), band2[:, :, 1].astype(int),
                       band2[:, :, 2].astype(int))
        assert int(((r2 > 120) & (r2 - g2 > 40) & (r2 - b2_ > 40)).sum()) == 0

    def test_rescoring_does_not_double_bake(self, tmp_path):
        """再生成してもドライバ出力から作り直すため二重焼き込みにならない"""
        import cv2
        scans, coord, key = _build_marked_inputs(tmp_path)
        b, _ = _bridge_through_recognition(scans, coord, key)
        b.set_handwriting("s1.png", 595, 842,
                          [{"color": "#2b2825", "width": 2.0,
                            "points": [[100, 650, 0.5], [300, 650, 0.5]]}])
        from constants import RESULTS_FOLDER, SCORED_FOLDER
        out = scans / RESULTS_FOLDER / SCORED_FOLDER / "s1.png"
        assert b.run_scoring()["ok"]; _wait_job_done(b)
        first = cv2.imdecode(np.fromfile(str(out), dtype=np.uint8),
                             cv2.IMREAD_COLOR)
        assert b.run_scoring()["ok"]; _wait_job_done(b)
        second = cv2.imdecode(np.fromfile(str(out), dtype=np.uint8),
                              cv2.IMREAD_COLOR)
        assert np.array_equal(first, second)
