"""
test_answer_key_checker.py — answer_key事前チェックとMarkdown書き出しのテスト

- check_answer_key: 正常/エラー/警告(中抜け・未登録行)/座標超過
- write_check_report / write_model_answer: Markdown内容
- run_answer_key_check: 2ファイル書き出しの統括
"""
import sys
from pathlib import Path

import pandas as pd
import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "main_src"))

from constants import MARK_FORMAT_STANDARD, MARK_FORMAT_MULTI_DIGIT
from answer_key_checker import (
    check_answer_key,
    write_check_report,
    write_model_answer,
    run_answer_key_check,
    _format_row_list,
)

sys.path.insert(0, str(PROJECT_ROOT / "tests"))
from test_multi_digit_mode import _create_answer_key, _create_coord_xlsx, MULTI_DIGIT_HEADERS


class TestFormatRowList:
    def test_consecutive_grouping(self):
        assert _format_row_list([2, 3, 4, 7]) == '2〜4, 7'
        assert _format_row_list([5]) == '5'
        assert _format_row_list([]) == ''


class TestCheckAnswerKey:
    def test_ok_multi_digit(self, tmp_path):
        path = tmp_path / "key.xlsx"
        _create_answer_key(path, [
            {'問題番号': 1, '正答': '-24', '配点': 3, '観点': 1},
            {'問題番号': 4, '正答': 'a', '配点': 2, '観点': 2},
        ])
        res = check_answer_key(path, mark_format=MARK_FORMAT_MULTI_DIGIT)
        assert res['ok'] is True
        assert res['errors'] == []
        assert [r['label'] for r in res['rows']] == ['1-3', '4']
        assert res['stats']['満点'] == 5
        assert res['stats']['観点別配点'] == {1: 3, 2: 2}
        assert res['stats']['最終使用行'] == 4

    def test_validation_errors_collected(self, tmp_path):
        path = tmp_path / "key.xlsx"
        _create_answer_key(path, [
            {'問題番号': '1-3', '正答': '-2', '配点': 3, '観点': 1},   # 長さ不一致
            {'問題番号': 4, '正答': 'xy', '配点': 2, '観点': 1},       # 不正記号
        ])
        res = check_answer_key(path, mark_format=MARK_FORMAT_MULTI_DIGIT)
        assert res['ok'] is False
        assert len(res['errors']) == 2
        assert any('文字列書式' in e for e in res['errors'])
        assert any('使用できない文字' in e for e in res['errors'])

    def test_gap_warning(self, tmp_path):
        """使用行の中抜けは警告(エラーではない)"""
        path = tmp_path / "key.xlsx"
        _create_answer_key(path, [
            {'問題番号': 1, '正答': '-24', '配点': 3, '観点': 1},  # 1〜3
            {'問題番号': 6, '正答': '5', '配点': 2, '観点': 1},    # 4,5が中抜け
        ])
        res = check_answer_key(path, mark_format=MARK_FORMAT_MULTI_DIGIT)
        assert res['ok'] is True
        assert any('4〜5' in w for w in res['warnings'])

    def test_incomplete_row_warning(self, tmp_path):
        """正答は書いたが配点を忘れた行（部分入力）は警告"""
        path = tmp_path / "key.xlsx"
        _create_answer_key(path, [
            {'問題番号': 1, '正答': '3', '配点': 2, '観点': 1},
            {'問題番号': 2, '正答': '5', '配点': '', '観点': ''},  # 配点忘れ
        ])
        res = check_answer_key(path, mark_format=MARK_FORMAT_STANDARD)
        assert res['ok'] is True
        assert any('入力が不完全' in w and '2' in w for w in res['warnings'])

    def test_empty_template_rows_no_noise(self, tmp_path):
        """自動生成テンプレの完全な空行（末尾未使用行）は警告しない"""
        rows = [{'問題番号': 1, '正答': '-24', '配点': 3, '観点': 1}]
        for q in range(2, 51):
            rows.append({'問題番号': q, '正答': '', '配点': '', '観点': ''})
        path = tmp_path / "key.xlsx"
        _create_answer_key(path, rows)
        res = check_answer_key(path, mark_format=MARK_FORMAT_MULTI_DIGIT)
        assert res['ok'] is True
        # 1-3自動割付＋4〜50は空 → 中抜けなし・部分入力なし → 警告ゼロ
        assert res['warnings'] == []

    def test_coordinate_overrun_error(self, tmp_path):
        """座標のマーク行数を超える問はエラー"""
        key_path = tmp_path / "key.xlsx"
        _create_answer_key(key_path, [
            {'問題番号': 3, '正答': '-24', '配点': 3, '観点': 1},  # 3〜5行を消費
        ])
        coord_path = tmp_path / "coord.xlsx"
        _create_coord_xlsx(coord_path, MULTI_DIGIT_HEADERS, num_questions=4)  # 行1〜4のみ
        res = check_answer_key(key_path, mark_format=MARK_FORMAT_MULTI_DIGIT,
                               coord_excel_path=coord_path, skip_questions=0)
        assert res['ok'] is False
        assert any('超える' in e for e in res['errors'])

    def test_missing_file(self, tmp_path):
        res = check_answer_key(tmp_path / "nothing.xlsx")
        assert res['ok'] is False
        assert any('見つかりません' in e for e in res['errors'])

    def test_empty_template(self, tmp_path):
        path = tmp_path / "key.xlsx"
        _create_answer_key(path, [
            {'問題番号': 1, '正答': '', '配点': '', '観点': ''},
        ])
        res = check_answer_key(path)
        assert res['ok'] is False
        assert any('1問もありません' in e for e in res['errors'])


class TestMarkdownOutputs:
    def _make_result(self, tmp_path):
        path = tmp_path / "answer_key.xlsx"
        _create_answer_key(path, [
            {'問題番号': 1, '正答': '-24', '配点': 3, '観点': 1},
            {'問題番号': 4, '正答': '', '配点': 1, '観点': 2, '特例': '全員正解'},
        ])
        return path

    def test_run_writes_two_files(self, tmp_path):
        path = self._make_result(tmp_path)
        res, check_md, model_md = run_answer_key_check(
            path, mark_format=MARK_FORMAT_MULTI_DIGIT)
        assert res['ok'] is True
        check_text = Path(check_md).read_text(encoding='utf-8')
        assert '✅ エラーはありません' in check_text
        assert '| 1-3 | 1〜3 | -24 | 3 |' in check_text
        assert '満点: 4点' in check_text
        model_text = Path(model_md).read_text(encoding='utf-8')
        assert '# 模範解答' in model_text
        assert '※全員正解' in model_text
        assert '**満点: 4点**' in model_text
        # 模範解答には問題概要列を載せない
        assert '問題概要' not in model_text

    def test_model_answer_not_written_on_error(self, tmp_path):
        path = tmp_path / "answer_key.xlsx"
        _create_answer_key(path, [
            {'問題番号': '1-3', '正答': '-2', '配点': 3, '観点': 1},
        ])
        res, check_md, model_md = run_answer_key_check(
            path, mark_format=MARK_FORMAT_MULTI_DIGIT)
        assert res['ok'] is False
        assert model_md is None
        check_text = Path(check_md).read_text(encoding='utf-8')
        assert '❌' in check_text

    def test_standard_mode_report(self, tmp_path):
        path = tmp_path / "answer_key.xlsx"
        _create_answer_key(path, [
            {'問題番号': 1, '正答': '3', '配点': 2, '観点': 1},
            {'問題番号': 2, '正答': '10', '配点': 2, '観点': 1},
        ])
        res, check_md, model_md = run_answer_key_check(path)
        assert res['ok'] is True
        assert Path(model_md).exists()
        text = Path(check_md).read_text(encoding='utf-8')
        assert '標準マーク' in text


class TestGuiIntegration:
    def _create_gui(self, mark_format=MARK_FORMAT_MULTI_DIGIT):
        import tkinter as tk
        from conftest import get_shared_tk_root
        from main_gui import SaitenSamuraiGUI
        from constants import MODE_MARK_ONLY
        root = get_shared_tk_root()
        top = tk.Toplevel(root)
        top.withdraw()
        app = SaitenSamuraiGUI(top, mode=MODE_MARK_ONLY, mark_format=mark_format)
        return top, app

    def test_check_button_writes_markdowns(self, tmp_path):
        """📋ボタン相当の呼び出しで2つのMarkdownが書き出され、ログに要約が出る"""
        key_path = tmp_path / "answer_key.xlsx"
        _create_answer_key(key_path, [
            {'問題番号': 1, '正答': '-24', '配点': 3, '観点': 1},
        ])
        top, app = self._create_gui()
        try:
            app.template_path.set(str(key_path))
            app.run_answer_key_check_gui(auto=True)
            assert (tmp_path / "answer_key_check.md").exists()
            assert (tmp_path / "answer_key_模範解答.md").exists()
            log = app.log_text.get("1.0", "end")
            assert '✅ エラーなし' in log
        finally:
            top.destroy()

    def test_auto_check_quiet_on_empty_key(self, tmp_path):
        """生成直後の空answer_keyでは自動チェックが騒がない（未入力の案内のみ）"""
        key_path = tmp_path / "answer_key.xlsx"
        _create_answer_key(key_path, [
            {'問題番号': 1, '正答': '', '配点': '', '観点': ''},
        ])
        top, app = self._create_gui()
        try:
            app.template_path.set(str(key_path))
            app.run_answer_key_check_gui(auto=True)
            log = app.log_text.get("1.0", "end")
            assert '正答が未入力です' in log
            assert '❌' not in log
        finally:
            top.destroy()
