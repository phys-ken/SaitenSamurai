#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
answer_key_checker.py — answer_key(正答データ)の事前チェックとMarkdown書き出し

採点実行前に登録ミスに気づけるよう、answer_key.xlsx を検証して
  1. <stem>_check.md   … 検証結果・行割当表・集計(自分の確認用)
  2. <stem>_模範解答.md … 問/解答番号/正答/配点/観点の表(配布用)
の2ファイルを書き出す。標準モード・複数桁モードの両方に対応する。

GUIからは正答データの選択/自動検出時と「📋 正答チェック」ボタンから呼ばれる。
"""

import datetime
import logging
from pathlib import Path

import pandas as pd

from constants import (
    MARK_FORMAT_STANDARD,
    MARK_FORMAT_MULTI_DIGIT,
)
from scoring_engine import load_template, number_to_circled, SPECIAL_ALL_CORRECT

logger = logging.getLogger(__name__)


def check_answer_key(template_path, mark_format=MARK_FORMAT_STANDARD,
                     coord_excel_path=None, skip_questions=0):
    """answer_keyを検証し、チェック結果と行割当・集計を返す。

    load_template のバリデーション(範囲重複・範囲長≠正答文字数・不正記号等)に加え、
    以下を検出する:
      - 警告: 使用行の中抜け(意図しない行飛ばしの気づき用)
      - 警告: 問題番号があるのに未登録(正答/配点空欄でスキップ)の行
      - エラー: 座標ファイルのマーク行数超過(coord_excel_path指定時)

    Args:
        template_path: answer_key.xlsxのパス
        mark_format: MARK_FORMAT_STANDARD / MARK_FORMAT_MULTI_DIGIT
        coord_excel_path: 座標ファイルのパス(任意。指定時は行数超過チェック)
        skip_questions: スキップ問題数(座標チェック用)

    Returns:
        dict: {
            'ok': bool,            # エラーが1件もない
            'errors': [str],
            'warnings': [str],
            'infos': [str],
            'template_dict': dict or None,  # 検証成功時のみ
            'rows': [dict],        # 行割当表 [{label, first, last, answer, ...}]
            'stats': dict or None, # 集計(検証成功時のみ)
        }
    """
    result = {
        'ok': False, 'errors': [], 'warnings': [], 'infos': [],
        'template_dict': None, 'rows': [], 'stats': None,
    }

    template_path = Path(template_path)
    if not template_path.exists():
        result['errors'].append(f"ファイルが見つかりません: {template_path}")
        return result

    # --- load_template による検証(エラーは集約メッセージで返る) ---
    try:
        template_dict = load_template(str(template_path), mark_format=mark_format)
    except ValueError as e:
        msg = str(e)
        # "answer_keyの検証エラー:\n..." のヘッダを外して行ごとに分解
        lines = [ln.strip() for ln in msg.split('\n') if ln.strip()]
        if lines and lines[0].startswith('answer_keyの検証エラー'):
            lines = lines[1:]
        result['errors'].extend(lines if lines else [msg])
        return result
    except Exception as e:
        result['errors'].append(f"読み込みエラー: {e}")
        return result

    if not template_dict:
        result['errors'].append("採点対象の問題が1問もありません（正答・配点が未入力です）")
        return result

    result['template_dict'] = template_dict

    # --- 行割当表 ---
    question_numbers = sorted(template_dict.keys())
    used_rows = set()
    rows = []
    for q in question_numbers:
        t = template_dict[q]
        span = t.get('span', 1)
        rows.append({
            'label': str(t.get('group_label', q)),
            'first': q,
            'last': q + span - 1,
            'span': span,
            'answer': t['正答'],
            'points': t['配点'],
            'aspect': t['観点'],
            'special': t.get('特例', ''),
            'summary': t.get('問題概要', ''),
        })
        used_rows.update(range(q, q + span))
    result['rows'] = rows

    # --- 警告: 使用行の中抜け ---
    first_used, last_used = min(used_rows), max(used_rows)
    gaps = [r for r in range(first_used, last_used + 1) if r not in used_rows]
    if gaps:
        result['warnings'].append(
            f"使用行の途中に未使用の解答番号があります: {_format_row_list(gaps)}"
            "（意図した行飛ばしなら問題ありません）")

    # --- 警告: 入力が不完全なため未登録(スキップ)になった行 ---
    # 完全な空行は自動生成テンプレの正常状態なので対象外(中抜けは上の警告が指摘する)
    skipped = _find_incomplete_rows(template_path, used_rows)
    if skipped:
        result['warnings'].append(
            f"入力が不完全（正答または配点が空欄）なため採点対象外の行があります: {_format_row_list(skipped)}")

    # --- 座標ファイルとの整合(任意) ---
    total_mark_rows = None
    if coord_excel_path and Path(str(coord_excel_path)).exists():
        try:
            from omr_engine import parse_excel_coordinates
            coordinates, _ = parse_excel_coordinates(str(coord_excel_path))
            coord_q_nos = {c['question_no'] for c in coordinates
                           if isinstance(c['question_no'], (int, float))}
            answer_rows = sorted(int(q) - skip_questions for q in coord_q_nos
                                 if q > skip_questions)
            if answer_rows:
                total_mark_rows = len(answer_rows)
                max_row = max(answer_rows)
                overruns = [r['label'] for r in rows if r['last'] > max_row]
                if overruns:
                    result['errors'].append(
                        f"座標定義のマーク行数({max_row}行)を超える問があります: {', '.join(overruns)}")
        except Exception as e:
            result['infos'].append(f"座標ファイルの整合チェックをスキップしました: {e}")

    # --- 集計 ---
    total_points = sum(r['points'] for r in rows)
    aspect_points = {}
    for r in rows:
        aspect_points[r['aspect']] = aspect_points.get(r['aspect'], 0) + r['points']
    special_count = sum(1 for r in rows if r['special'])
    stats = {
        '問題数': len(rows),
        '満点': total_points,
        '観点別配点': aspect_points,
        '使用マーク行数': len(used_rows),
        '最終使用行': last_used,
        '特例(全員正解)': special_count,
    }
    if total_mark_rows is not None:
        stats['座標のマーク行数'] = total_mark_rows
        stats['残り行数'] = total_mark_rows - last_used
    result['stats'] = stats

    result['ok'] = not result['errors']
    return result


def _find_incomplete_rows(template_path, used_rows):
    """入力が不完全なため未登録になった行番号を返す。

    「正答・配点・特例のいずれかに入力があるのに登録されなかった行」のみを対象とし、
    完全な空行(自動生成テンプレの未使用行)は正常として無視する。
    例: 正答は書いたが配点を忘れた行、特例だけ書いて配点を忘れた行。
    """
    skipped = []
    try:
        df = pd.read_excel(template_path)
        if '問題番号' not in df.columns:
            return []

        def _has_value(row, col):
            if col not in df.columns:
                return False
            v = row[col]
            return not (pd.isna(v) or str(v).strip() == '')

        for _, row in df.iterrows():
            raw = row['問題番号']
            if pd.isna(raw) or str(raw).strip() == '':
                continue
            try:
                # 単独intの行のみ対象(範囲表記の行はload_templateが登録/エラー済み)
                q = int(float(str(raw).strip()))
            except (ValueError, TypeError):
                continue
            if q in used_rows:
                continue
            if _has_value(row, '正答') or _has_value(row, '配点') or _has_value(row, '特例'):
                skipped.append(q)
    except Exception:
        return []
    return sorted(set(skipped))


def _format_row_list(nums):
    """[2,3,4,7] → '2〜4, 7' のように連続をまとめて表示する。"""
    if not nums:
        return ''
    nums = sorted(nums)
    parts = []
    start = prev = nums[0]
    for n in nums[1:]:
        if n == prev + 1:
            prev = n
            continue
        parts.append(f"{start}〜{prev}" if prev > start else str(start))
        start = prev = n
    parts.append(f"{start}〜{prev}" if prev > start else str(start))
    return ', '.join(parts)


# ========================================
# Markdown 書き出し
# ========================================


def _answer_display(row):
    """正答の表示文字列(特例で空欄なら「—」)"""
    return row['answer'] if row['answer'] else '—'


def _rows_note(row):
    return '※全員正解' if row['special'] == SPECIAL_ALL_CORRECT else ''


def write_check_report(check_result, out_path, template_name='', mark_format=MARK_FORMAT_STANDARD):
    """チェック報告Markdown(検証結果・行割当表・集計)を書き出す。"""
    fmt_name = '数学マーク（複数桁）' if mark_format == MARK_FORMAT_MULTI_DIGIT else '標準マーク'
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    lines = [
        '# answer_key チェック報告',
        '',
        f'- ファイル: `{template_name}`',
        f'- モード: {fmt_name}',
        f'- チェック日時: {now}',
        '',
        '## 検証結果',
        '',
    ]
    if check_result['errors']:
        lines.append(f"❌ **エラー {len(check_result['errors'])}件** — 修正するまで採点できません")
        lines.append('')
        for e in check_result['errors']:
            lines.append(f'- ❌ {e}')
    else:
        lines.append('✅ エラーはありません')
    if check_result['warnings']:
        lines.append('')
        for w in check_result['warnings']:
            lines.append(f'- ⚠ {w}')
    if check_result['infos']:
        lines.append('')
        for i in check_result['infos']:
            lines.append(f'- ℹ {i}')
    lines.append('')

    if check_result['rows']:
        lines += [
            '## 行割当・登録内容',
            '',
            '| 問 | 解答番号 | 正答 | 配点 | 観点 | 特例 | 問題概要 |',
            '|---|---|---|---|---|---|---|',
        ]
        for r in check_result['rows']:
            row_range = f"{r['first']}" if r['span'] == 1 else f"{r['first']}〜{r['last']}"
            lines.append(
                f"| {r['label']} | {row_range} | {_answer_display(r)} | {r['points']} "
                f"| {number_to_circled(r['aspect'])} | {r['special']} | {r['summary']} |")
        lines.append('')

    stats = check_result['stats']
    if stats:
        lines += ['## 集計', '']
        lines.append(f"- 問題数: {stats['問題数']}問 / 満点: {stats['満点']}点")
        aspect_str = ' / '.join(
            f"観点{number_to_circled(a)}: {p}点"
            for a, p in sorted(stats['観点別配点'].items()))
        lines.append(f"- {aspect_str}")
        lines.append(f"- 使用マーク行: {stats['使用マーク行数']}行（最終使用行: 解答番号{stats['最終使用行']}）")
        if '座標のマーク行数' in stats:
            lines.append(f"- 座標のマーク行数: {stats['座標のマーク行数']}行（残り {stats['残り行数']}行）")
        if stats['特例(全員正解)']:
            lines.append(f"- 特例（全員正解）: {stats['特例(全員正解)']}問")
        lines.append('')

    Path(out_path).write_text('\n'.join(lines), encoding='utf-8')
    return str(out_path)


def write_model_answer(check_result, out_path, title='模範解答'):
    """配布用の模範解答Markdownを書き出す(エラーがある場合は書き出さずNoneを返す)。"""
    if not check_result['ok'] or not check_result['rows']:
        return None

    stats = check_result['stats']
    lines = [
        f'# {title}',
        '',
        '| 問 | 解答番号 | 正答 | 配点 | 観点 | 備考 |',
        '|---|---|---|---|---|---|',
    ]
    for r in check_result['rows']:
        row_range = f"{r['first']}" if r['span'] == 1 else f"{r['first']}〜{r['last']}"
        lines.append(
            f"| {r['label']} | {row_range} | {_answer_display(r)} | {r['points']}点 "
            f"| {number_to_circled(r['aspect'])} | {_rows_note(r)} |")
    aspect_str = ' / '.join(
        f"観点{number_to_circled(a)}: {p}点"
        for a, p in sorted(stats['観点別配点'].items()))
    lines += [
        '',
        f"**満点: {stats['満点']}点**（{aspect_str}）",
        '',
    ]
    Path(out_path).write_text('\n'.join(lines), encoding='utf-8')
    return str(out_path)


def run_answer_key_check(template_path, mark_format=MARK_FORMAT_STANDARD,
                         coord_excel_path=None, skip_questions=0):
    """チェックを実行し、answer_keyと同じフォルダにMarkdown 2ファイルを書き出す。

    Returns:
        (check_result, check_md_path, model_md_path)
        model_md_path はエラー時 None。
    """
    template_path = Path(template_path)
    check_result = check_answer_key(template_path, mark_format=mark_format,
                                    coord_excel_path=coord_excel_path,
                                    skip_questions=skip_questions)
    stem = template_path.stem
    check_md = template_path.parent / f"{stem}_check.md"
    model_md = template_path.parent / f"{stem}_模範解答.md"
    write_check_report(check_result, check_md, template_name=template_path.name,
                       mark_format=mark_format)
    model_written = write_model_answer(check_result, model_md)
    return check_result, str(check_md), model_written
