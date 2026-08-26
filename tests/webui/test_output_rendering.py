"""採点済み答案の出力: 合計点表示位置と表示項目設定（L1）。

- 合計点の満点/観点別トグル（描画層のレイアウト単体）
- 記述系モードでの total_display_region 注入（位置設定が無視されるバグの回帰）
- set_rendering_settings の検証と永続化
"""
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

SCORING_RESULT = {
    "total_score": 82,
    "max_score": 100,
    "aspect_scores": {1: 40, 2: 42},
    "aspect_max_scores": {1: 50, 2: 50},
}


class TestTotalLayoutToggles:
    def test_default_shows_max_and_aspects(self):
        from image_renderer import _prepare_total_score_layout
        img = np.full((842, 595, 3), 255, dtype=np.uint8)
        line1, line2, *_ = _prepare_total_score_layout(img, SCORING_RESULT, None, 1.0)
        assert line1 == "得点：82 / 100"
        assert "観点①：40/50" in line2 and "観点②：42/50" in line2

    def test_hide_max(self):
        from image_renderer import _prepare_total_score_layout
        img = np.full((842, 595, 3), 255, dtype=np.uint8)
        line1, line2, *_ = _prepare_total_score_layout(
            img, SCORING_RESULT, None, 1.0,
            rendering_settings={"total_show_max": False})
        assert line1 == "得点：82"
        assert line2 != ""   # 観点行は既定のまま

    def test_hide_aspects(self):
        from image_renderer import _prepare_total_score_layout
        img = np.full((842, 595, 3), 255, dtype=np.uint8)
        line1, line2, *_ = _prepare_total_score_layout(
            img, SCORING_RESULT, None, 1.0,
            rendering_settings={"total_show_aspects": False})
        assert line1 == "得点：82 / 100"
        assert line2 == ""

    def test_combined_total_toggles_and_region(self):
        """draw_combined_total: 指定領域の中に描き、トグルで描画量が減る"""
        from descriptive_renderer import draw_combined_total
        config = {"questions": [{"id": "D1", "name": "問1", "max_score": 10,
                                 "aspect": 1, "region": [10, 10, 100, 60]}],
                  "total_display_region": [300, 700, 560, 800]}
        white = np.full((842, 595, 3), 255, dtype=np.uint8)

        def drawn_pixels(img, y1, y2, x1, x2):
            return int((img[y1:y2, x1:x2] != 255).any(axis=2).sum())

        full = draw_combined_total(white, SCORING_RESULT, config, {"D1": 7})
        # 指定領域の中に描かれ、外（左上）は白いまま
        assert drawn_pixels(full, 700, 800, 300, 560) > 50
        assert drawn_pixels(full, 0, 400, 0, 595) == 0

        minimal = draw_combined_total(
            white, SCORING_RESULT, config, {"D1": 7},
            rendering_settings={"total_show_max": False,
                                "total_show_aspects": False})
        # トグルで出力が変わる。ピクセル数の大小はフォント自動拡大で
        # 環境により逆転する（Windows CI で顕在化）ため、内容の差だけを固定し、
        # テキスト自体は _combined_total_lines を直接検証する
        assert not np.array_equal(full, minimal)
        from descriptive_renderer import _combined_total_lines
        line1, line2 = _combined_total_lines(
            89, 110, SCORING_RESULT["aspect_scores"],
            SCORING_RESULT["aspect_max_scores"],
            {"total_show_max": False, "total_show_aspects": False})
        assert line1 == "得点：89" and line2 == ""
        line1, line2 = _combined_total_lines(
            89, 110, SCORING_RESULT["aspect_scores"],
            SCORING_RESULT["aspect_max_scores"])
        assert line1 == "得点：89 / 110" and "観点①:40/50" in line2


class TestRenderingSettingsApi:
    def test_validation_and_diff_storage(self):
        b = Bridge(window_adapter=RecordingAdapter())
        assert not b.set_rendering_settings({"unknown_key": 1})["ok"]
        assert not b.set_rendering_settings("not a dict")["ok"]
        assert b.set_rendering_settings({"total_show_max": False,
                                         "show_score": True})["ok"]
        # 既定値と同じ show_score は差分に残らない
        assert b.state["rendering_settings"] == {"total_show_max": False}
        full = b.get_rendering_settings()
        assert full["settings"]["total_show_max"] is False
        assert full["settings"]["show_score"] is True
        # 既定値へ戻すと差分が空 → None
        assert b.set_rendering_settings({"total_show_max": True})["ok"]
        assert b.state["rendering_settings"] is None

    def test_opacity_number_validation(self):
        b = Bridge(window_adapter=RecordingAdapter())
        assert not b.set_rendering_settings({"descriptive_opacity": "abc"})["ok"]
        assert b.set_rendering_settings({"descriptive_opacity": "0.3"})["ok"]
        assert b.state["rendering_settings"] == {"descriptive_opacity": 0.3}


def _make_desc_bridge(tmp_path):
    """記述のみモード・画像準備済みの bridge"""
    import cv2
    adapter = RecordingAdapter()
    b = Bridge(window_adapter=adapter)
    b.set_mode("descriptive_only", "standard")
    scans = tmp_path / "scans"
    scans.mkdir()
    for name in ("s1.png", "s2.png"):
        img = np.full((842, 595, 3), 255, dtype=np.uint8)
        cv2.rectangle(img, (60, 60), (240, 140), (40, 40, 40), 2)
        cv2.imwrite(str(scans / name), img)
    adapter.folder_returns = [str(scans)]
    assert b.select_image_folder()["ok"]
    assert b.run_prepare_images()["ok"]
    _wait_job_done(b)
    return b, scans


class TestTotalRegionInjection:
    def test_config_injection_helper(self, tmp_path):
        b, _ = _make_desc_bridge(tmp_path)
        b.add_descriptive_question("問1", 5, 1, [50, 50, 250, 150])
        assert "total_display_region" not in b._desc_config
        assert b.set_total_display_region([300, 700, 560, 800])["ok"]
        cfg = b._desc_config_with_total_region()
        assert cfg["total_display_region"] == [300, 700, 560, 800]
        # 元の config は汚さない
        assert "total_display_region" not in b._desc_config

    def test_state_reflects_saved_region_after_reselect(self, tmp_path):
        """フォルダ選択で保存済み位置が state に載る（ヒント表示の回帰）"""
        b, scans = _make_desc_bridge(tmp_path)
        assert b.set_total_display_region([300, 700, 560, 800])["ok"]
        assert b.state["total_display_region"] == [300, 700, 560, 800]

        b2 = Bridge(window_adapter=RecordingAdapter())
        b2.set_mode("descriptive_only", "standard")
        b2._win.folder_returns = [str(scans)]
        assert b2.select_image_folder()["ok"]
        assert b2.state["total_display_region"] == [300, 700, 560, 800]

    def test_desc_only_scoring_draws_total_at_region(self, tmp_path):
        """回帰: 記述のみ生成が位置設定を反映する（従来は無視されていた）"""
        import cv2
        b, scans = _make_desc_bridge(tmp_path)
        b.add_descriptive_question("問1", 5, 1, [50, 50, 250, 150])
        b.start_descriptive_scoring()
        b.set_descriptive_score("s1.png", "D1", 4)
        assert b.set_total_display_region([300, 650, 560, 750])["ok"]
        assert b.run_scoring()["ok"]
        _wait_job_done(b)

        from constants import RESULTS_FOLDER, SCORED_FOLDER
        out = scans / RESULTS_FOLDER / SCORED_FOLDER / "s1.png"
        assert out.exists()
        img = cv2.imdecode(np.fromfile(str(out), dtype=np.uint8),
                           cv2.IMREAD_COLOR)
        region = img[650:750, 300:560]
        drawn = int((region < 200).any(axis=2).sum())
        assert drawn > 50, "指定領域に合計点が描かれていない（注入漏れ）"


class TestRenderPreview:
    def test_preview_requires_inputs(self):
        b = Bridge(window_adapter=RecordingAdapter())
        res = b.get_render_preview()
        assert not res["ok"] and "座標ファイル" in res["error"]

    def test_preview_reflects_offset(self, tmp_path):
        """プレビューは実描画関数の出力で、オフセット変更で画像が変わる"""
        import base64
        import cv2
        sys.path.insert(0, str(PROJECT_ROOT / "tests"))
        from test_multi_digit_mode import _create_answer_key
        from test_multidigit_image_e2e import (make_coord_xlsx, make_sheet,
                                               sym_positions)
        coord = make_coord_xlsx(tmp_path / "coord.xlsx", 4)
        key = tmp_path / "key.xlsx"
        _create_answer_key(key, [
            {'問題番号': 1, '正答': '-24', '配点': 3, '観点': 1},
        ])
        scans = tmp_path / "scans"
        scans.mkdir()
        filled = {q: p for q, p in zip([1, 2, 3], sym_positions('-24'))}
        cv2.imwrite(str(scans / "s1.png"),
                    make_sheet(filled, with_markers=True))

        adapter = RecordingAdapter()
        b = Bridge(window_adapter=adapter)
        b.set_mode("mark_only", "multi_digit")
        adapter.folder_returns = [str(scans)]
        adapter.file_returns = [str(coord), str(key)]
        assert b.select_image_folder()["ok"]
        b.set_skip_questions(0)   # フォルダ選択は完全クリアするので後から設定
        assert b.select_coord_file()["ok"]
        assert b.select_answer_key()["ok"]
        assert b.run_recognition()["ok"]
        _wait_job_done(b)

        def decode(res):
            raw = base64.b64decode(res["data_url"].split(",", 1)[1])
            return cv2.imdecode(np.frombuffer(raw, np.uint8), cv2.IMREAD_COLOR)

        res0 = b.get_render_preview()
        assert res0["ok"], res0
        img0 = decode(res0)
        assert img0 is not None and img0.size > 0
        # 白紙ではなく何かが描かれている
        assert int((img0 < 200).any(axis=2).sum()) > 30

        assert b.set_rendering_settings({"mark_result_offset": 2.0})["ok"]
        img1 = decode(b.get_render_preview())
        assert img0.shape == img1.shape
        assert not np.array_equal(img0, img1), "オフセットがプレビューに反映されていない"
