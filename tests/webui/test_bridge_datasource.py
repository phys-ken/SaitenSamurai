"""webui Bridge のデータソース選択（L1: pywebview なしで直接検証）。

tk 版と同じ「選んだ瞬間に検証する」振る舞いを固定する。
CI(windows-latest) でも実行される — bridge は pywebview を import しない設計。
"""
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "webui"))
sys.path.insert(0, str(PROJECT_ROOT / "main_src"))
sys.path.insert(0, str(PROJECT_ROOT / "tests"))

from api.bridge import Bridge  # noqa: E402
from test_multi_digit_mode import (  # noqa: E402
    _create_answer_key, _create_coord_xlsx,
    MULTI_DIGIT_HEADERS, STANDARD_HEADERS,
)


class FakeAdapter:
    """ネイティブダイアログの偽物 — 返すパスをキューで指定"""

    def __init__(self):
        self.folder_returns = []
        self.file_returns = []

    def open_folder_dialog(self, directory=""):
        return self.folder_returns.pop(0) if self.folder_returns else None

    def open_file_dialog(self, file_types=None, directory=""):
        return self.file_returns.pop(0) if self.file_returns else None


@pytest.fixture
def bridge():
    return Bridge(window_adapter=FakeAdapter())


class TestImageFolder:
    def test_counts_images(self, bridge, tmp_path):
        for name in ("a.png", "b.jpg", "c.txt"):
            (tmp_path / name).write_bytes(b"x")
        bridge._win.folder_returns = [str(tmp_path)]
        res = bridge.select_image_folder()
        assert res["ok"] and res["state"]["image_count"] == 2

    def test_empty_folder_rejected(self, bridge, tmp_path):
        bridge._win.folder_returns = [str(tmp_path)]
        res = bridge.select_image_folder()
        assert not res["ok"] and "画像" in res["error"]
        assert bridge.state["image_folder"] is None

    def test_cancel_keeps_state(self, bridge):
        res = bridge.select_image_folder()  # ダイアログでキャンセル
        assert res["ok"] and res.get("cancelled") is True


class TestCoordFile:
    def test_summary_on_select(self, bridge, tmp_path):
        coord = tmp_path / "coord.xlsx"
        _create_coord_xlsx(coord, STANDARD_HEADERS, num_questions=8)
        bridge.set_skip_questions(0)
        bridge._win.file_returns = [str(coord)]
        res = bridge.select_coord_file()
        assert res["ok"]
        s = res["state"]["coord_summary"]
        assert s["answer_rows"] == 8 and s["marks_per_row"] == 10
        assert s["warning"] is None

    def test_mode_mismatch_warning(self, bridge, tmp_path):
        """数学モード × 標準座標 → 選択時点で警告（tk 版と同じ）"""
        coord = tmp_path / "coord.xlsx"
        _create_coord_xlsx(coord, STANDARD_HEADERS, num_questions=8)
        bridge.set_mode("mark_only", "multi_digit")
        bridge.set_skip_questions(0)
        bridge._win.file_returns = [str(coord)]
        res = bridge.select_coord_file()
        assert res["ok"]
        assert "数学マーク採点" in res["state"]["coord_summary"]["warning"]

    def test_mode_change_revalidates(self, bridge, tmp_path):
        """選択後にモードを変えると警告も追従する"""
        coord = tmp_path / "coord.xlsx"
        _create_coord_xlsx(coord, MULTI_DIGIT_HEADERS, num_questions=4)
        bridge.set_skip_questions(0)
        bridge._win.file_returns = [str(coord)]
        bridge.select_coord_file()
        assert "標準マーク採点モード" in bridge.state["coord_summary"]["warning"]
        bridge.set_mode("mark_only", "multi_digit")
        assert bridge.state["coord_summary"]["warning"] is None

    def test_unreadable_file_rolls_back(self, bridge, tmp_path):
        bogus = tmp_path / "not_excel.xlsx"
        bogus.write_text("これはExcelではない", encoding="utf-8")
        bridge._win.file_returns = [str(bogus)]
        res = bridge.select_coord_file()
        assert not res["ok"] and "読み込めません" in res["error"]
        assert bridge.state["coord_file"] is None

    def test_skip_change_revalidates(self, bridge, tmp_path):
        coord = tmp_path / "coord.xlsx"
        _create_coord_xlsx(coord, STANDARD_HEADERS, num_questions=8)
        bridge.set_skip_questions(0)
        bridge._win.file_returns = [str(coord)]
        bridge.select_coord_file()
        res = bridge.set_skip_questions(4)
        assert res["state"]["coord_summary"]["answer_rows"] == 4


class TestAnswerKey:
    def test_auto_check_on_select(self, bridge, tmp_path):
        key = tmp_path / "key.xlsx"
        _create_answer_key(key, [
            {'問題番号': 1, '正答': '3', '配点': 2, '観点': 1},
            {'問題番号': 2, '正答': '5', '配点': 3, '観点': 1},
        ])
        bridge._win.file_returns = [str(key)]
        res = bridge.select_answer_key()
        summary = res["state"]["key_summary"]
        assert summary["ok"] is True
        assert "2問" in summary["stats_line"] and "満点5点" in summary["stats_line"]
        assert summary["check_md"] and Path(summary["check_md"]).exists()

    def test_validation_errors_surface(self, bridge, tmp_path):
        """複数桁モードの不正正答は errors として UI に届く"""
        key = tmp_path / "key.xlsx"
        _create_answer_key(key, [
            {'問題番号': '1-3', '正答': '-2', '配点': 3, '観点': 1},  # 文字数不一致
        ])
        bridge.set_mode("mark_only", "multi_digit")
        bridge._win.file_returns = [str(key)]
        res = bridge.select_answer_key()
        summary = res["state"]["key_summary"]
        assert summary["ok"] is False and len(summary["errors"]) >= 1
