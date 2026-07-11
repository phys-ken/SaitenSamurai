#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scoring_engine.py — 採点コアロジック

テンプレート読込、OMR結果読込、採点判定の純粋ロジック。
画像・ファイルIOに依存せず、データ構造のみで入出力を行う。

将来の拡張:
    - 複数正答パターンの追加
    - 配点変更ルールの追加
    - 記述式問題との統合採点
"""

import datetime
import logging
import re

import pandas as pd

from constants import (
    MARK_FORMAT_STANDARD,
    MARK_FORMAT_MULTI_DIGIT,
    MULTI_DIGIT_VALUE_TO_SYMBOL,
    MULTI_DIGIT_SYMBOL_TO_VALUE,
    MULTI_DIGIT_VALID_SYMBOLS,
)

logger = logging.getLogger(__name__)


def number_to_circled(num):
    """数字を丸数字に変換（1→①, 2→②, ...）"""
    if 1 <= num <= 50:
        return chr(0x2460 + num - 1)  # ①-㊿
    return str(num)


def normalize_value(value):
    """
    値を正規化して文字列に変換
    
    Args:
        value: 入力値（数値、文字列、NaNなど）
    
    Returns:
        正規化された文字列
    """
    if pd.isna(value):
        return ''
    
    # 浮動小数点数の場合
    if isinstance(value, float):
        # 整数値と等しい場合は整数に変換（例: 1.0 -> 1）
        if value == int(value):
            value = int(value)
    
    # 文字列化して空白除去
    return str(value).strip()


def normalize_zero_ten(value):
    """マークシート選択肢の "10" を "0" に正規化する。

    マークシートは最大10択(1,2,...,9,0)で、10番目のマーク位置は選択肢"0"。
    旧データ形式では同じ位置が"10"と記録されていたため、
    正誤比較の前に必ずこの関数で両辺を正規化する(後方互換)。
    採点(score_answers)とCTT分析(ctt_analyzer)は必ずこのヘルパーを
    経由し、0⇔10の扱いが食い違わないようにする。
    """
    s = str(value).strip()
    return '0' if s == '10' else s


def normalize_answer_set(answer_str):
    """複数正答文字列("3;10" / "3|10")を0⇔10正規化済みの集合に変換する。"""
    if not answer_str:
        return set()
    return {normalize_zero_ten(p) for p in str(answer_str).replace('|', ';').split(';')}


def choice_to_position_index(choice, num_choices, mark_format=MARK_FORMAT_STANDARD):
    """選択肢の値をマーク位置のインデックス(0始まり)に変換する。

    standard: マークシートの並びは 1,2,...,9,0 で、選択肢"0"は10番目の位置
    ("10"はレガシー表記で"0"と同義)。数値でない・位置が存在しない
    場合は None を返す。

    multi_digit(複数桁設問モード): 並びは -,0,1,...,9,a,b,c,d で
    位置 = ヘッダ値+1("-"のヘッダ値は-1)。記号('-','a'等)と
    数値表記('-1','10'等)の両方を受理する。0⇔10の正規化は行わない。

    正答位置の赤字表示(image_renderer)・正答枠オーバーレイ(gui_components)
    など「選択肢の値→物理的なマーク位置」の解決は必ずこの関数を経由し、
    箇所ごとに変換ルールがズレないようにする。

    Args:
        choice: 選択肢の値(str/int。"3", "0", "10", 3.0、multi_digitでは"-","a"など)
        num_choices: その設問のマーク位置の数
        mark_format: MARK_FORMAT_STANDARD / MARK_FORMAT_MULTI_DIGIT

    Returns:
        0始まりの位置インデックス、または None
    """
    if mark_format == MARK_FORMAT_MULTI_DIGIT:
        s = str(choice).strip().lower()
        if s in MULTI_DIGIT_SYMBOL_TO_VALUE:
            v = MULTI_DIGIT_SYMBOL_TO_VALUE[s]
        else:
            try:
                f = float(s)
            except (ValueError, TypeError):
                return None
            if not f.is_integer():
                return None
            v = int(f)
            if v not in MULTI_DIGIT_VALUE_TO_SYMBOL:
                return None
        index = v + 1  # ヘッダ値-1("-")が位置0
        return index if 0 <= index < num_choices else None

    try:
        f = float(normalize_zero_ten(choice))
    except (ValueError, TypeError):
        return None
    if not f.is_integer():
        return None
    v = int(f)
    if v == 0:
        index = 9  # 選択肢"0" = 10番目のマーク位置
    elif 1 <= v <= 9:
        index = v - 1
    else:
        return None
    return index if index < num_choices else None


SPECIAL_ALL_CORRECT = '全員正解'
# 特例列に入力可能な値。これ以外の値は警告を出して通常採点にフォールバックする
VALID_SPECIAL_VALUES = (SPECIAL_ALL_CORRECT,)

# 複数桁モード: 問題番号セルの範囲表記 "1-3" (全角ハイフン類は事前に半角へ正規化)
_RANGE_PATTERN = re.compile(r'^(\d+)\s*-\s*(\d+)$')
_FULLWIDTH_HYPHENS = str.maketrans({'－': '-', 'ー': '-', '−': '-', '‐': '-', '～': '-', '〜': '-'})


def _check_excel_date_cell(q_no_raw):
    """問題番号セルがExcelの日付自動変換に化けていないか検査する。

    Excelは「1-3」のような入力を日付(1月3日)へ自動変換するため、
    ユーザーは範囲表記を書いたつもりでもセルはdatetimeになっている。
    日付型ならヒント付きのValueErrorを送出する。
    """
    if isinstance(q_no_raw, (datetime.datetime, datetime.date)):
        raise ValueError(
            f"問題番号セルが日付({q_no_raw})になっています。"
            "Excelが「1-3」等の入力を日付に自動変換した可能性があります。"
            "セルの書式を「文字列」にしてから入力し直してください")


def _parse_question_range(q_no_raw):
    """複数桁モードの問題番号セルを (先頭行番号, span, ラベル, 明示範囲か) に解析する。

    "1-3" → (1, 3, "1-3", True) / "5" → (5, 1, "5", False)。解析不能なら ValueError。
    明示範囲でない行は、正答の文字数から消費行数を自動割付できる
    (load_template側で span を上書きする)。
    """
    _check_excel_date_cell(q_no_raw)
    s = normalize_value(q_no_raw).translate(_FULLWIDTH_HYPHENS)
    m = _RANGE_PATTERN.match(s)
    if m:
        start, end = int(m.group(1)), int(m.group(2))
        if start >= end:
            raise ValueError(f"問題番号『{s}』が不正です（範囲は 開始-終了 の昇順で記入してください）")
        return start, end - start + 1, f"{start}-{end}", True
    try:
        q_no = int(float(s))
    except (ValueError, TypeError):
        raise ValueError(f"問題番号『{s}』を解釈できません（単独行は「5」、範囲は「1-3」の形式で記入してください）")
    return q_no, 1, s, False


def load_template(template_path, mark_format=MARK_FORMAT_STANDARD):
    """
    採点用テンプレートを読み込み

    正答が未登録（空欄・NaN）の問題は採点対象外としてスキップする。
    ただし特例列が「全員正解」の問題は、正答が空欄でも採点対象に含める
    （不適切問題の救済措置。配点は必須のまま）。

    複数桁モード(mark_format=MARK_FORMAT_MULTI_DIGIT)では問題番号列の
    範囲表記「1-3」を受理し、先頭行番号をキーに 'span'(消費するマーク行数) と
    'group_label'(表示用ラベル) を付与する。正答は記号列("-24"等)。
    - 明示範囲「1-3」: 範囲長=正答文字数をバリデーションする
    - 単独表記「1」+複数文字正答: 正答の文字数ぶん連続行を自動割付する
      (例: 問題番号1に正答"-24" → 1〜3行を消費、group_label="1-3")
    - 特例「全員正解」で正答空欄の複数行グループは範囲表記が必須
      (単独表記では消費行数を推論できないため span=1 になる)

    Args:
        template_path: 正答データExcelファイルのパス
        mark_format: MARK_FORMAT_STANDARD / MARK_FORMAT_MULTI_DIGIT

    Returns:
        問題番号をキーとした辞書（正答登録済みの問題のみ）
    """
    df = pd.read_excel(template_path)

    # 必要な列をチェック
    required_columns = ['問題番号', '正答', '配点', '観点']
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"テンプレートに列'{col}'が見つかりません")

    # 特例は任意列(古いテンプレートには存在しない)
    has_special_col = '特例' in df.columns

    # 辞書形式に変換（正答未登録の行はスキップ）
    template_dict = {}
    skipped_questions = []
    multi_digit = (mark_format == MARK_FORMAT_MULTI_DIGIT)
    errors = []          # multi_digit: バリデーションエラーを集約して一括報告
    occupied_rows = {}   # multi_digit: マーク行番号 -> group_label（範囲重複の検出）
    for _, row in df.iterrows():
        # 問題番号が空欄・NaNの場合はスキップ（Excelの末尾空行対策）
        q_no_raw = row['問題番号']
        if pd.isna(q_no_raw) or str(q_no_raw).strip() == '':  # type: ignore[arg-type]
            continue
        if multi_digit:
            try:
                q_no, span, group_label, is_explicit_range = _parse_question_range(q_no_raw)
            except ValueError as e:
                errors.append(str(e))
                continue
        else:
            _check_excel_date_cell(q_no_raw)
            try:
                q_no = int(float(q_no_raw))
            except (ValueError, TypeError):
                s = normalize_value(q_no_raw).translate(_FULLWIDTH_HYPHENS)
                if _RANGE_PATTERN.match(s):
                    # 範囲表記は複数桁モード専用 — モード違いを案内する
                    raise ValueError(
                        f"問題番号『{s}』の範囲表記は数学マーク採点（複数桁）モード用です。"
                        "起動画面で「数学マーク採点」を選んでから読み込んでください")
                raise ValueError(
                    f"問題番号『{s}』を数値として解釈できません。"
                    "問題番号列には数値を入力してください")
            span = 1
            group_label = str(q_no)
            is_explicit_range = False
        # 特例区分の読み取り（想定外の値は警告して通常扱い）
        special = ''
        if has_special_col:
            special_raw = normalize_value(row['特例'])
            if special_raw:
                if special_raw in VALID_SPECIAL_VALUES:
                    special = special_raw
                else:
                    logger.warning("問題%s: 特例列の値'%s'は未対応のため通常採点します（対応値: %s）",
                                   group_label, special_raw, '/'.join(VALID_SPECIAL_VALUES))
        # 正答が空欄・NaNの場合は採点対象外（全員正解の特例問題は正答空欄を許容）
        answer_val = row['正答']
        answer_empty = bool(pd.isna(answer_val)) or str(answer_val).strip() == ''  # type: ignore[arg-type]
        if answer_empty and special != SPECIAL_ALL_CORRECT:
            skipped_questions.append(q_no)
            continue
        # 配点が空欄・NaNの場合も採点対象外
        points_val = row['配点']
        if pd.isna(points_val) or str(points_val).strip() == '':  # type: ignore[arg-type]
            skipped_questions.append(q_no)
            continue

        if multi_digit:
            # 正答は記号列("-24"等)。大文字A-Dは小文字に正規化
            answer_str = '' if answer_empty else normalize_value(answer_val).lower()
            if ';' in answer_str or '|' in answer_str:
                errors.append(f"問題{group_label}: 複数正答（;区切り）は複数桁設問モードでは未対応です")
                continue
            bad_chars = sorted({ch for ch in answer_str if ch not in MULTI_DIGIT_VALID_SYMBOLS})
            if bad_chars:
                errors.append(f"問題{group_label}: 正答『{answer_str}』に使用できない文字『{''.join(bad_chars)}』があります"
                              f"（使用可能: - 0〜9 a〜d）")
                continue
            if is_explicit_range:
                # 明示範囲: 範囲長=正答文字数を検証(全員正解の正答空欄は除く)
                if answer_str and len(answer_str) != span:
                    errors.append(f"問題{group_label}: 範囲{span}行に対し正答『{answer_str}』が{len(answer_str)}文字です"
                                  f"（先頭ゼロを含む正答はセルを文字列書式で入力してください）")
                    continue
            elif len(answer_str) > 1:
                # 自動割付: 単独表記の行は正答の文字数ぶん連続行を消費する
                # (行番号は各行に固定されているためズレは起きず、
                #  消費先に別の登録があれば下の重複チェックで検出される)
                span = len(answer_str)
                group_label = f"{q_no}-{q_no + span - 1}"
            conflict = next((r for r in range(q_no, q_no + span) if r in occupied_rows), None)
            if conflict is not None:
                errors.append(f"問題番号{group_label}と{occupied_rows[conflict]}が重複しています（マーク行{conflict}）")
                continue
            for r in range(q_no, q_no + span):
                occupied_rows[r] = group_label
        else:
            answer_str = normalize_value(answer_val)

        template_dict[q_no] = {
            '正答': answer_str,
            '配点': int(float(points_val)),
            '観点': int(float(row['観点'])) if not pd.isna(row['観点']) else 1,  # type: ignore[arg-type]
            # 問題概要は任意列(古いテンプレートには存在しない)
            '問題概要': normalize_value(row['問題概要']) if '問題概要' in df.columns else '',
            '特例': special
        }
        if multi_digit:
            template_dict[q_no]['span'] = span
            template_dict[q_no]['group_label'] = group_label

    if errors:
        raise ValueError("answer_keyの検証エラー:\n" + "\n".join(errors))

    if skipped_questions:
        total = len(df)
        registered = len(template_dict)
        logger.info("テンプレート: %d問中%d問が採点対象、%d問スキップ（正答未登録）", total, registered, len(skipped_questions))
        logger.info("スキップされた問題: %s", skipped_questions)

    special_questions = [q for q, t in template_dict.items() if t.get('特例')]
    if special_questions:
        logger.info("特例（%s）が設定された問題: %s", SPECIAL_ALL_CORRECT, special_questions)

    return template_dict


def load_mark2_results(mark2_result_path, skip_questions=0):
    """
    Mark2読取結果Excelを読み込み
    
    Args:
        mark2_result_path: Mark2結果Excelのパス
        skip_questions: スキップする問題数（後方互換性のため残すが、ヘッダー行の値を優先する）
    
    Returns:
        学生データのリスト
    """
    # ヘッダーなしで読み込み、構造を解析する
    # Row 0: No, File, 1, 2, 3... (Original Indices)
    # Row 1: NaN, NaN, 学年, クラス, ..., 1, 2, 3... (Question Names / Scored Indices)
    # Row 2+: Data
    df = pd.read_excel(mark2_result_path, header=None)
    
    # File列のインデックスを探す（通常は1列目）
    # Row 0を検索
    header_row = df.iloc[0]
    file_col_idx = -1
    for idx, val in enumerate(header_row):
        if str(val).lower() == 'file':
            file_col_idx = idx
            break
            
    if file_col_idx == -1:
        # 見つからない場合、1列目と仮定
        file_col_idx = 1
    
    # 列マッピングを作成 (Column Index -> Scored Question Number)
    # Row 1 (設問名行) を使用する
    col_mapping = {}
    if len(df) > 1:
        name_row = df.iloc[1]
        
        # 1. まず数値ヘッダーを収集して、オフセットが必要か判定する
        # skip_questions分の列（ID列など）はスキップして判定する
        start_check_idx = file_col_idx + 1 + skip_questions
        numeric_headers = []
        
        for idx in range(start_check_idx, len(name_row)):
            val = name_row[idx]
            if pd.notna(val):
                s_val = str(val).strip()
                try:
                    f_val = float(s_val)
                    if f_val == int(f_val):
                        numeric_headers.append(int(f_val))
                except ValueError:
                    pass
        
        # オフセット判定: 最小値が skip_questions より大きい場合は、
        # 元の番号(Original Index)が使われていると判断し、オフセットを適用する
        needs_offset = False
        if numeric_headers and skip_questions > 0:
            min_val = min(numeric_headers)
            if min_val > skip_questions:
                needs_offset = True
        
        # 2. マッピング作成
        # ID列も含めてスキャンするが、マッピング時に調整する
        for idx in range(file_col_idx + 1, len(name_row)):
            # skip_questions以内の列は、明示的に除外する（ID列として扱う）
            # ただし、needs_offset=False（名前ベース）の場合は、名前が"1"ならQ1として扱うべきか？
            # 安全のため、skip_questions分は常にスキップする
            if idx < file_col_idx + 1 + skip_questions:
                continue
                
            val = name_row[idx]
            if pd.notna(val):
                s_val = str(val).strip()
                try:
                    f_val = float(s_val)
                    if f_val == int(f_val):
                        q_num = int(f_val)
                        
                        if needs_offset:
                            q_num -= skip_questions
                            
                        col_mapping[idx] = q_num
                except ValueError:
                    pass
    
    # マッピングが見つからない場合（Row 1が空など）、Row 0を使ってskip_questionsで計算する（バックアップ）
    if not col_mapping:
        for idx in range(file_col_idx + 1, len(header_row)):
            val = header_row[idx]
            if pd.notna(val) and str(val).isdigit():
                original_q = int(val)
                if original_q > skip_questions:
                    col_mapping[idx] = original_q - skip_questions

    # 両方のヘッダー解析が失敗した場合、空のcol_mappingのまま処理を続けると
    # 全生徒が空白・0点として扱われてしまうため、ここで明示的に停止する。
    if not col_mapping:
        raise ValueError(
            f"OMR読取結果ファイルの設問列を認識できませんでした: {mark2_result_path}\n"
            "ヘッダー行（1行目・2行目）が想定外の形式になっている可能性があります。"
        )

    # データを抽出
    results = []
    # データ行はRow 2から
    for row_idx in range(2, len(df)):
        row = df.iloc[row_idx]
        
        # File列の値を確認
        file_val = row[file_col_idx]
        if pd.isna(file_val) or not str(file_val).strip():
            continue
            
        # 画像ファイル名かチェック（簡易的）
        image_name = str(file_val).strip()
        if not (image_name.lower().endswith('.jpg') or image_name.lower().endswith('.png')):
            continue
            
        answers = {}
        for col_idx, q_no in col_mapping.items():
            if col_idx < len(row):
                value = row[col_idx]
                answers[q_no] = normalize_value(value)
        
        results.append({
            'image': image_name,
            'answers': answers
        })
    
    return results


def score_answers(student_answers, template_dict, mark_format=MARK_FORMAT_STANDARD):
    """
    学生の解答を採点

    複数桁モード(mark_format=MARK_FORMAT_MULTI_DIGIT)では、テンプレートの
    'span' に従い q_no から連続する複数マーク行の解答を連結して正答文字列と
    完全一致比較する（完答のみ得点）。グループ内に無マーク・ダブルマークが
    1行でもあれば不正解。0⇔10の等価判定は行わない
    （例: 2行グループの '1'+'0' → "10" が "0" に正規化される事故を防ぐ）。

    Args:
        student_answers: {問題番号: 解答} の辞書
        template_dict: テンプレート辞書
        mark_format: MARK_FORMAT_STANDARD / MARK_FORMAT_MULTI_DIGIT

    Returns:
        採点結果の辞書
    """
    total_score = 0
    max_score = 0
    aspect_scores = {}
    aspect_max_scores = {}
    results = {}
    multi_digit = (mark_format == MARK_FORMAT_MULTI_DIGIT)

    for q_no, template_data in template_dict.items():
        correct_answer = template_data['正答']
        points = template_data['配点']
        aspect = template_data['観点']
        special = template_data.get('特例', '')
        span = template_data.get('span', 1)

        max_score += points

        # 観点別満点・得点を初期化
        if aspect not in aspect_max_scores:
            aspect_max_scores[aspect] = 0
        if aspect not in aspect_scores:
            aspect_scores[aspect] = 0
        aspect_max_scores[aspect] += points

        # 採点
        is_correct = False
        earned_points = 0

        if multi_digit:
            # グループ各行の解答を取得して連結
            parts = [str(student_answers.get(q_no + i, '')).strip().lower() for i in range(span)]
            # 無マーク(空)・ダブルマーク(;入り)が1行でもあればグループ全体を不正解とする
            invalid = any(p == '' or ';' in p for p in parts)
            if invalid:
                # 無効時は行の生値をカンマ区切りで残す(集計・目視確認用)
                student_answer = ','.join(parts)
            else:
                student_answer = ''.join(parts)
            if special == SPECIAL_ALL_CORRECT:
                # 特例(全員正解): グループ単位で満点を与える
                is_correct = True
            elif not invalid:
                # 完答判定: 連結文字列の完全一致(0⇔10正規化は行わない)
                is_correct = (student_answer == correct_answer)
        else:
            # 学生の解答を取得
            student_answer = student_answers.get(q_no, '')

            # 後方互換性: かつてload_mark2_resultsが0→10変換していた時代の
            # 既存データとの互換のため、0⇔10の等価判定を維持する。
            # 現在のパイプラインでは raw_choice(座標Excelの値ヘッダ)ベースで
            # 出力しており、10列テンプレートの最終列は raw_choice=0 として記録される。
            if special == SPECIAL_ALL_CORRECT:
                # 特例(全員正解): 無回答・誤答を問わず全員に満点を与える
                is_correct = True
            elif ';' in correct_answer or '|' in correct_answer:
                # 複数正答: 0⇔10正規化済みの集合同士で比較
                is_correct = normalize_answer_set(correct_answer) == normalize_answer_set(student_answer)
            else:
                # 単一正答: 両辺を0⇔10正規化して比較
                is_correct = normalize_zero_ten(student_answer) == normalize_zero_ten(correct_answer)

        if is_correct:
            earned_points = points
            total_score += points

            # 観点別得点を加算
            if aspect not in aspect_scores:
                aspect_scores[aspect] = 0
            aspect_scores[aspect] += points

        results[q_no] = {
            'correct': is_correct,
            'points': earned_points,
            'max_points': points,
            'student_answer': student_answer,
            'correct_answer': correct_answer,
            'aspect': aspect,
            'special': special
        }
        if multi_digit:
            results[q_no]['span'] = span
            results[q_no]['group_label'] = template_data.get('group_label', str(q_no))

    return {
        'total_score': total_score,
        'max_score': max_score,
        'aspect_scores': aspect_scores,
        'aspect_max_scores': aspect_max_scores,
        'results': results
    }
