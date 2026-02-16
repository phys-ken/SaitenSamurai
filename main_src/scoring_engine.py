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

import logging

import pandas as pd

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


def load_template(template_path):
    """
    採点用テンプレートを読み込み
    
    正答が未登録（空欄・NaN）の問題は採点対象外としてスキップする。
    
    Args:
        template_path: 正答データExcelファイルのパス
    
    Returns:
        問題番号をキーとした辞書（正答登録済みの問題のみ）
    """
    df = pd.read_excel(template_path)
    
    # 必要な列をチェック
    required_columns = ['問題番号', '正答', '配点', '観点']
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"テンプレートに列'{col}'が見つかりません")
    
    # 辞書形式に変換（正答未登録の行はスキップ）
    template_dict = {}
    skipped_questions = []
    for _, row in df.iterrows():
        # 問題番号が空欄・NaNの場合はスキップ（Excelの末尾空行対策）
        q_no_raw = row['問題番号']
        if pd.isna(q_no_raw) or str(q_no_raw).strip() == '':  # type: ignore[arg-type]
            continue
        q_no = int(float(q_no_raw))
        # 正答が空欄・NaNの場合は採点対象外
        answer_val = row['正答']
        if pd.isna(answer_val) or str(answer_val).strip() == '':  # type: ignore[arg-type]
            skipped_questions.append(q_no)
            continue
        # 配点が空欄・NaNの場合も採点対象外
        points_val = row['配点']
        if pd.isna(points_val) or str(points_val).strip() == '':  # type: ignore[arg-type]
            skipped_questions.append(q_no)
            continue
        template_dict[q_no] = {
            '正答': normalize_value(answer_val),
            '配点': int(float(points_val)),
            '観点': int(float(row['観点'])) if not pd.isna(row['観点']) else 1  # type: ignore[arg-type]
        }
    
    if skipped_questions:
        total = len(df)
        registered = len(template_dict)
        logger.info("テンプレート: %d問中%d問が採点対象、%d問スキップ（正答未登録）", total, registered, len(skipped_questions))
        logger.info("スキップされた問題: %s", skipped_questions)
    
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


def score_answers(student_answers, template_dict):
    """
    学生の解答を採点
    
    Args:
        student_answers: {問題番号: 解答} の辞書
        template_dict: テンプレート辞書
    
    Returns:
        採点結果の辞書
    """
    total_score = 0
    max_score = 0
    aspect_scores = {}
    aspect_max_scores = {}
    results = {}
    
    for q_no, template_data in template_dict.items():
        correct_answer = template_data['正答']
        points = template_data['配点']
        aspect = template_data['観点']
        
        max_score += points
        
        # 観点別満点・得点を初期化
        if aspect not in aspect_max_scores:
            aspect_max_scores[aspect] = 0
        if aspect not in aspect_scores:
            aspect_scores[aspect] = 0
        aspect_max_scores[aspect] += points
        
        # 学生の解答を取得
        student_answer = student_answers.get(q_no, '')
        
        # 採点
        is_correct = False
        earned_points = 0
        
        if ';' in correct_answer or '|' in correct_answer:
            # 複数正答の場合
            correct_set = set(correct_answer.replace('|', ';').split(';'))
            student_set = set(student_answer.replace('|', ';').split(';')) if student_answer else set()
            is_correct = correct_set == student_set
        else:
            # 単一正答
            # 後方互換性: かつてload_mark2_resultsが0→10変換していた時代の
            # 既存データとの互換のため、0⇔10の等価判定を維持する。
            # 現在のパイプラインでは raw_choice ベースで出力しており、
            # 10列テンプレートの最終列は raw_choice=0 として記録される。
            if correct_answer == '0' and student_answer == '10':
                is_correct = True
            elif correct_answer == '10' and student_answer == '0':
                is_correct = True
            else:
                is_correct = (student_answer == correct_answer)
        
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
            'aspect': aspect
        }
    
    return {
        'total_score': total_score,
        'max_score': max_score,
        'aspect_scores': aspect_scores,
        'aspect_max_scores': aspect_max_scores,
        'results': results
    }
