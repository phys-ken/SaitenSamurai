"""webui 記述採点API（L1: 実 descriptive_scorer・記述のみモードの通し）。"""
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "webui"))
sys.path.insert(0, str(PROJECT_ROOT / "main_src"))
sys.path.insert(0, str(PROJECT_ROOT / "tests"))
sys.path.insert(0, str(PROJECT_ROOT / "tests" / "webui"))

from api.bridge import Bridge  # noqa: E402
from test_bridge_jobs import RecordingAdapter, _wait_job_done  # noqa: E402


@pytest.fixture
def desc_bridge(tmp_path):
    """記述のみモード・答案2枚（文字入りの合成画像）・画像準備済み"""
    import cv2
    import numpy as np
    adapter = RecordingAdapter()
    b = Bridge(window_adapter=adapter)
    b.set_mode("descriptive_only", "standard")
    scans = tmp_path / "scans"
    scans.mkdir()
    for name, shade in (("s1.png", 30), ("s2.png", 90)):
        img = np.full((400, 300, 3), 255, dtype=np.uint8)
        cv2.rectangle(img, (60, 60), (240, 140), (shade,) * 3, -1)  # 解答らしき塗り
        cv2.imwrite(str(scans / name), img)
    adapter.folder_returns = [str(scans)]
    assert b.select_image_folder()["ok"]
    assert b.run_prepare_images()["ok"]
    _wait_job_done(b)
    return b, adapter, scans


class TestConfigEditing:
    def test_add_update_delete_roundtrip(self, desc_bridge):
        b, _, scans = desc_bridge
        res = b.add_descriptive_question("問1", 5, 1, [50, 50, 250, 150])
        assert res["ok"] and res["question_id"] == "D1"
        res = b.add_descriptive_question("問2", 3, 2, [50, 200, 250, 300])
        assert res["question_id"] == "D2"

        # tk 互換のファイルに保存されている
        from descriptive_scorer import load_descriptive_config
        cfg = load_descriptive_config(str(b._desc_config_path()))
        assert [q["id"] for q in cfg["questions"]] == ["D1", "D2"]
        assert cfg["questions"][0]["region"] == [50, 50, 250, 150]

        assert b.update_descriptive_region("D2", [10, 10, 60, 60])["ok"]
        assert b.delete_descriptive_question("D1")["ok"]
        cfg = load_descriptive_config(str(b._desc_config_path()))
        assert [q["id"] for q in cfg["questions"]] == ["D2"]

    def test_delete_cleans_scores(self, desc_bridge):
        b, _, _ = desc_bridge
        b.add_descriptive_question("問1", 5, 1, [50, 50, 250, 150])
        b.start_descriptive_scoring()
        b.set_descriptive_score("s1.png", "D1", 4)
        b.delete_descriptive_question("D1")
        from descriptive_scorer import load_descriptive_scores
        data = load_descriptive_scores(str(b._desc_scores_path()))
        assert data["scores"]["s1.png"] == {}

    def test_id_never_reused_after_delete(self, desc_bridge):
        """削除後の追加で ID が再利用されない（過去スコアと衝突しない）"""
        b, _, _ = desc_bridge
        b.add_descriptive_question("A", 5, 1, [0, 0, 50, 50])
        b.add_descriptive_question("B", 5, 1, [0, 60, 50, 110])
        b.delete_descriptive_question("D1")
        res = b.add_descriptive_question("C", 5, 1, [0, 120, 50, 170])
        assert res["question_id"] == "D3"


class TestScoringFlow:
    def test_crop_score_persist(self, desc_bridge):
        b, _, _ = desc_bridge
        b.add_descriptive_question("問1", 5, 1, [50, 50, 250, 150])
        assert b.start_descriptive_scoring()["ok"]

        targets = b.list_descriptive_targets("D1")
        assert [t["filename"] for t in targets["items"]] == ["s1.png", "s2.png"]
        assert all(t["score"] is None for t in targets["items"])

        crop = b.get_descriptive_crop("D1", "s1.png")
        assert crop["ok"] and crop["data_url"].startswith("data:image/png")

        assert b.set_descriptive_score("s1.png", "D1", 5)["ok"]
        assert not b.set_descriptive_score("s1.png", "D1", 6)["ok"]   # 配点超過
        assert b.set_descriptive_score("s2.png", "D1", 0)["ok"]
        st = b.state["descriptive"]
        assert st["scored_counts"]["D1"] == 2

        # tk 互換形式（version + scores）で保存されている
        import json
        data = json.loads(b._desc_scores_path().read_text(encoding="utf-8"))
        assert data["version"] == 1
        assert data["scores"]["s1.png"]["D1"] == 5

    def test_generate_sheets_and_summary(self, desc_bridge):
        import cv2
        import numpy as np
        from constants import (RESULTS_FOLDER, SCORED_FOLDER,
                               FINAL_REPORT_FOLDER, STUDENT_SUMMARY_FILE)
        b, adapter, scans = desc_bridge
        b.add_descriptive_question("問1", 5, 1, [50, 50, 250, 150])
        b.start_descriptive_scoring()
        b.set_descriptive_score("s1.png", "D1", 5)
        b.set_descriptive_score("s2.png", "D1", 2)

        assert b.run_scoring()["ok"]
        _wait_job_done(b)
        done = [e for e in adapter.events if e["type"] == "job_done"][-1]
        assert done["ok"], done
        scored = scans / RESULTS_FOLDER / SCORED_FOLDER
        out_img = cv2.imread(str(scored / "s1.png"))
        src_img = cv2.imread(str(scans / "s1.png"))
        assert out_img is not None
        assert not np.array_equal(out_img, src_img), "得点が描画されていない"

        assert b.run_summary()["ok"]
        _wait_job_done(b)
        done = [e for e in adapter.events if e["type"] == "job_done"][-1]
        assert done["ok"], done
        summary = scans / RESULTS_FOLDER / FINAL_REPORT_FOLDER / STUDENT_SUMMARY_FILE
        assert summary.exists()
        import openpyxl
        ws = openpyxl.load_workbook(summary, data_only=True).active
        values = [c.value for row in ws.iter_rows() for c in row]
        assert 5 in values and 2 in values
