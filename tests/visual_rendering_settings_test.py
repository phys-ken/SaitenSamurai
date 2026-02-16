#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
visual_rendering_settings_test.py — ④ 採点結果描画 詳細設定のビジュアルテスト

sample_bigfiles の実データを使って、rendering_settings の各パラメータを
変更した採点済み答案画像を生成し、HTML でプレビューできるレポートを作成する。

実行方法:
  python tests/visual_rendering_settings_test.py

出力:
  tests/visual_rendering_report/
    ├── index.html  (メインレポート)
    ├── case_*.jpg   (各テストケースの採点済み画像)
    └── crop_*.jpg   (拡大クロップ画像)
"""

import sys
import os
import json
import base64
import hashlib
from pathlib import Path
from datetime import datetime

# プロジェクトセットアップ
PROJECT_ROOT = Path(__file__).resolve().parent.parent
MAIN_SRC = PROJECT_ROOT / "main_src"
sys.path.insert(0, str(MAIN_SRC))

import cv2
import numpy as np

from constants import (
    DEFAULT_RENDERING_SETTINGS,
    get_rendering_settings,
    RESULTS_FOLDER, BOXED_FOLDER, RESULTS_DATA_FOLDER,
)
from image_renderer import draw_scoring_results, draw_total_score
from scoring_engine import load_template, load_mark2_results, score_answers
from omr_engine import parse_excel_coordinates, detect_corner_markers, compute_output_scale, apply_perspective_transform


# ========================================
# テストデータのパス
# ========================================
SAMPLE_DIR = PROJECT_ROOT / "sample_bigfiles"
COORD_EXCEL = SAMPLE_DIR / "〇datas" / "M2-03-002_座標ファイル.xlsx"
TEMPLATE_PATH = SAMPLE_DIR / "〇datas" / "answer_key.xlsx"
# 代替: sample_bigfiles 直下のResult
MARK2_RESULT_CANDIDATES = [
    SAMPLE_DIR / "_saiten_grading_results" / "reading_results" / "Mark2-Result-A040-C010-20260211_034340.xlsx",
    SAMPLE_DIR / "_saiten_grading_results" / "reading_results" / "Mark2-Result-A042-C020-20260210_185056.xlsx",
]
BOXED_DIR = SAMPLE_DIR / "_saiten_grading_results" / "boxed_figs"
SKIP_QUESTIONS = 4

# 出力先
OUTPUT_DIR = PROJECT_ROOT / "tests" / "visual_rendering_report"


# ========================================
# テストケース定義
# ========================================
TEST_CASES = [
    {
        "id": "default",
        "label": "デフォルト設定",
        "desc": "全項目ON、オフセット0、透過率50%",
        "settings": {},
    },
    {
        "id": "hide_ox",
        "label": "○×マーク非表示",
        "desc": "show_ox_mark=False: ○×なし、得点・観点が左に詰まる",
        "settings": {"show_ox_mark": False},
    },
    {
        "id": "hide_score",
        "label": "得点非表示",
        "desc": "show_score=False: ○×と観点のみ表示",
        "settings": {"show_score": False},
    },
    {
        "id": "hide_aspect",
        "label": "観点非表示",
        "desc": "show_aspect=False: ○×と得点のみ表示",
        "settings": {"show_aspect": False},
    },
    {
        "id": "hide_correct",
        "label": "正答番号非表示",
        "desc": "show_correct_answer=False: ×のとき赤数字なし",
        "settings": {"show_correct_answer": False},
    },
    {
        "id": "only_ox",
        "label": "○×のみ表示",
        "desc": "得点・観点・正答すべて非表示",
        "settings": {"show_score": False, "show_aspect": False, "show_correct_answer": False},
    },
    {
        "id": "only_score",
        "label": "得点のみ表示",
        "desc": "○×・観点・正答すべて非表示、得点だけ表示",
        "settings": {"show_ox_mark": False, "show_aspect": False, "show_correct_answer": False},
    },
    {
        "id": "all_hidden",
        "label": "全項目非表示",
        "desc": "全項目OFF: 何も描画されない（元画像＋合計点のみ）",
        "settings": {
            "show_ox_mark": False, "show_score": False,
            "show_aspect": False, "show_correct_answer": False,
        },
    },
    {
        "id": "offset_neg2",
        "label": "オフセット -2",
        "desc": "mark_result_offset=-2: デフォルト位置からさらに左に2セル幅分",
        "settings": {"mark_result_offset": -2},
    },
    {
        "id": "offset_neg1",
        "label": "オフセット -1",
        "desc": "mark_result_offset=-1: デフォルトより1セル幅分左",
        "settings": {"mark_result_offset": -1},
    },
    {
        "id": "offset_neg0_5",
        "label": "オフセット -0.5",
        "desc": "mark_result_offset=-0.5: 半セル幅左（小数オフセット）",
        "settings": {"mark_result_offset": -0.5},
    },
    {
        "id": "offset_pos0_3",
        "label": "オフセット +0.3",
        "desc": "mark_result_offset=+0.3: 0.3セル幅右（微調整）",
        "settings": {"mark_result_offset": 0.3},
    },
    {
        "id": "offset_pos0_5",
        "label": "オフセット +0.5",
        "desc": "mark_result_offset=+0.5: 半セル幅右（小数オフセット）",
        "settings": {"mark_result_offset": 0.5},
    },
    {
        "id": "offset_pos1",
        "label": "オフセット +1",
        "desc": "mark_result_offset=+1: デフォルトより1セル幅分右",
        "settings": {"mark_result_offset": 1},
    },
    {
        "id": "offset_pos1_5",
        "label": "オフセット +1.5",
        "desc": "mark_result_offset=+1.5: 1.5セル幅右（枠外はみだし領域）",
        "settings": {"mark_result_offset": 1.5},
    },
    {
        "id": "offset_pos2",
        "label": "オフセット +2",
        "desc": "mark_result_offset=+2: 2セル幅右（枠外はみだし）",
        "settings": {"mark_result_offset": 2},
    },
    {
        "id": "offset_pos3",
        "label": "オフセット +3",
        "desc": "mark_result_offset=+3: 3セル幅右（大きな枠外はみだし確認）",
        "settings": {"mark_result_offset": 3},
    },
    {
        "id": "offset_extreme_left",
        "label": "オフセット -10 (極端左)",
        "desc": "極端な左シフト → 枠外遠方（クランプなし）",
        "settings": {"mark_result_offset": -10},
    },
    {
        "id": "offset_extreme_right",
        "label": "オフセット +10 (極端右)",
        "desc": "極端な右シフト → 枠外遠方（クランプなし）",
        "settings": {"mark_result_offset": 10},
    },
    {
        "id": "combined_hide_ox_neg1",
        "label": "○×非表示 + オフセット-1",
        "desc": "複合: ○×なし＋左シフト",
        "settings": {"show_ox_mark": False, "mark_result_offset": -1},
    },
    {
        "id": "combined_score_only_pos1",
        "label": "得点のみ + オフセット+1",
        "desc": "複合: 得点のみ＋右シフト",
        "settings": {
            "show_ox_mark": False, "show_aspect": False,
            "show_correct_answer": False, "mark_result_offset": 1,
        },
    },
]


def load_test_data():
    """テストに必要なデータを読み込む"""
    print("データ読み込み中...")

    # テンプレート
    template_dict = load_template(str(TEMPLATE_PATH))
    print(f"  テンプレート: {len(template_dict)}問")

    # 座標
    coordinates, _ = parse_excel_coordinates(str(COORD_EXCEL), SKIP_QUESTIONS)
    print(f"  座標: {len(coordinates)}個")

    # Mark2結果
    mark2_result_path = None
    for p in MARK2_RESULT_CANDIDATES:
        if p.exists():
            mark2_result_path = p
            break
    if mark2_result_path is None:
        raise FileNotFoundError("Mark2結果ファイルが見つかりません")

    mark2_results = load_mark2_results(str(mark2_result_path), SKIP_QUESTIONS)
    print(f"  Mark2結果: {len(mark2_results)}件")

    return template_dict, coordinates, mark2_results


def load_and_correct_image(image_name):
    """画像を読み込み、射影変換で補正"""
    # まず boxed_figs から補正済みを試行
    boxed_path = BOXED_DIR / image_name
    if boxed_path.exists():
        img = _imread_safe(boxed_path)
        if img is not None:
            return img, compute_output_scale(img)

    # 直接読み込み + 補正
    image_path = SAMPLE_DIR / image_name
    if not image_path.exists():
        return None, 1.0

    img = _imread_safe(image_path)
    if img is None:
        return None, 1.0

    markers = detect_corner_markers(img, debug=False)
    output_scale = compute_output_scale(img)
    corrected, _ = apply_perspective_transform(img, markers, output_scale=output_scale)
    return corrected, output_scale


def crop_question_area(image, coordinates, skip_questions, q_range=(1, 5), margin=60):
    """指定問題番号範囲の領域をクロップ（○×/得点が見える範囲、枠外も包含）"""
    target_coords = []
    for q_no in range(q_range[0], q_range[1] + 1):
        target_q = q_no + skip_questions
        target_coords.extend([c for c in coordinates if c['question_no'] == target_q])

    if not target_coords:
        h, w = image.shape[:2]
        return image[0:min(400, h), 0:min(800, w)]

    xs = [c['x'] for c in target_coords]
    ys = [c['y'] for c in target_coords]
    ws = [c['width'] for c in target_coords]
    hs = [c['height'] for c in target_coords]

    x1 = max(0, min(xs) - margin)
    y1 = max(0, min(ys) - margin)
    x2 = min(image.shape[1], max(x + w for x, w in zip(xs, ws)) + margin)
    y2 = min(image.shape[0], max(y + h for y, h in zip(ys, hs)) + margin)

    return image[y1:y2, x1:x2]


def images_differ(img1, img2):
    """2つの画像が異なるか判定"""
    if img1.shape != img2.shape:
        return True
    diff = cv2.absdiff(img1, img2)
    return np.sum(diff) > 0


def count_pixel_diff(img1, img2):
    """異なるピクセル数を返す"""
    if img1.shape != img2.shape:
        return -1
    diff = cv2.absdiff(img1, img2)
    gray_diff = cv2.cvtColor(diff, cv2.COLOR_BGR2GRAY)
    return int(np.count_nonzero(gray_diff))


def _imread_safe(path):
    """Unicode パスでも安全に読み込む cv2.imread 代替"""
    try:
        img = cv2.imread(str(path))
        if img is not None:
            return img
    except Exception:
        pass
    # fallback: numpy fromfile
    try:
        buf = np.fromfile(str(path), dtype=np.uint8)
        return cv2.imdecode(buf, cv2.IMREAD_COLOR)
    except Exception:
        return None


def _imwrite_safe(path, img, params=None):
    """Unicode パスでも安全に書き込む cv2.imwrite 代替"""
    try:
        ok = cv2.imwrite(str(path), img, params)
        if ok:
            return True
    except Exception:
        pass
    # fallback: cv2.imencode + tofile
    try:
        ext = Path(path).suffix  # e.g. '.jpg'
        success, buf = cv2.imencode(ext, img, params or [])
        if success:
            buf.tofile(str(path))
            return True
    except Exception:
        pass
    return False


def generate_test_images(template_dict, coordinates, mark2_results):
    """全テストケースの画像を生成"""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # 最初の生徒データを使用
    student_data = mark2_results[0]
    image_name = student_data['image']
    student_answers = student_data['answers']
    print(f"\n対象画像: {image_name}")

    # 画像読み込み
    corrected_image, output_scale = load_and_correct_image(image_name)
    if corrected_image is None:
        raise RuntimeError(f"画像の読み込みに失敗: {image_name}")
    print(f"  画像サイズ: {corrected_image.shape[1]}x{corrected_image.shape[0]}")
    print(f"  output_scale: {output_scale:.3f}")

    # 採点
    scoring_result = score_answers(student_answers, template_dict)
    print(f"  採点結果: {scoring_result['total_score']}/{scoring_result['max_score']}点")
    print(f"  正答数: {sum(1 for r in scoring_result['results'].values() if r['correct'])}")
    print(f"  誤答数: {sum(1 for r in scoring_result['results'].values() if not r['correct'])}")

    results = []
    default_full = None
    default_crop = None

    for tc in TEST_CASES:
        print(f"  生成中: {tc['id']} ({tc['label']})...")
        rs = get_rendering_settings(tc['settings'])

        # 採点結果描画
        scored = draw_scoring_results(
            corrected_image, coordinates, scoring_result,
            skip_questions=SKIP_QUESTIONS,
            output_scale=output_scale,
            rendering_settings=rs,
        )
        # 合計点描画
        scored = draw_total_score(
            scored, coordinates, scoring_result,
            output_scale=output_scale,
        )

        # フルサイズ保存（JPEG品質80で十分）
        full_path = OUTPUT_DIR / f"case_{tc['id']}_full.jpg"
        _imwrite_safe(full_path, scored, [cv2.IMWRITE_JPEG_QUALITY, 80])

        # 問題1-5のクロップ保存
        crop = crop_question_area(scored, coordinates, SKIP_QUESTIONS, q_range=(1, 5), margin=15)
        crop_path = OUTPUT_DIR / f"case_{tc['id']}_crop.jpg"
        _imwrite_safe(crop_path, crop, [cv2.IMWRITE_JPEG_QUALITY, 90])

        # 問題1-15のやや広いクロップ
        wide_crop = crop_question_area(scored, coordinates, SKIP_QUESTIONS, q_range=(1, 15), margin=15)
        wide_path = OUTPUT_DIR / f"case_{tc['id']}_wide.jpg"
        _imwrite_safe(wide_path, wide_crop, [cv2.IMWRITE_JPEG_QUALITY, 85])

        # デフォルト画像との差分
        diff_from_default = 0
        if tc['id'] == 'default':
            default_full = scored.copy()
            default_crop = crop.copy()
        elif default_full is not None:
            diff_from_default = count_pixel_diff(scored, default_full)

        results.append({
            "id": tc['id'],
            "label": tc['label'],
            "desc": tc['desc'],
            "settings": tc['settings'],
            "full_path": full_path.name,
            "crop_path": crop_path.name,
            "wide_path": wide_path.name,
            "diff_pixels": diff_from_default,
            "differs_from_default": tc['id'] != 'default' and diff_from_default > 0,
            "may_match_default": tc.get('may_match_default', False),
        })

    return results, scoring_result


def validate_results(results):
    """AI精読による自動検証ロジック"""
    checks = []

    # (1) デフォルトと実際に異なるケースの検証
    for r in results:
        if r['id'] == 'default':
            continue
        may_match = r.get('may_match_default', False)
        if r['settings'] and not r['differs_from_default']:
            if may_match:
                checks.append({
                    "status": "PASS",
                    "case": r['id'],
                    "message": f"デフォルトと同一画像（選択肢数により同位置に回帰 — 正常動作）"
                })
            else:
                checks.append({
                    "status": "FAIL",
                    "case": r['id'],
                    "message": f"設定変更 {r['settings']} があるのにデフォルトと画像が同一です。設定が反映されていません。"
                })
        elif r['settings'] and r['differs_from_default']:
            checks.append({
                "status": "PASS",
                "case": r['id'],
                "message": f"設定変更が画像に反映されています。差分ピクセル: {r['diff_pixels']:,}"
            })

    # (2) all_hidden が他と異なることを確認
    all_hidden = next((r for r in results if r['id'] == 'all_hidden'), None)
    only_ox = next((r for r in results if r['id'] == 'only_ox'), None)
    if all_hidden and only_ox:
        img_ah = _imread_safe(OUTPUT_DIR / all_hidden['full_path'])
        img_oo = _imread_safe(OUTPUT_DIR / only_ox['full_path'])
        if img_ah is not None and img_oo is not None:
            if images_differ(img_ah, img_oo):
                checks.append({
                    "status": "PASS",
                    "case": "all_hidden vs only_ox",
                    "message": "全項目非表示と○×のみ表示は異なる画像（正しい）"
                })
            else:
                checks.append({
                    "status": "FAIL",
                    "case": "all_hidden vs only_ox",
                    "message": "全項目非表示と○×のみ表示が同一画像（○×描画が機能していない）"
                })

    # (3) オフセット変更が相互に異なることを確認
    # クランプなし: 異なるオフセット値は常に異なる描画位置になるはず
    # ただし combined_ ケースは除外（複合設定は別検証）
    offset_cases = [r for r in results
                    if 'offset' in r['id']
                    and 'extreme' not in r['id']
                    and 'combined' not in r['id']]
    for i, r1 in enumerate(offset_cases):
        for r2 in offset_cases[i+1:]:
            img1 = _imread_safe(OUTPUT_DIR / r1['full_path'])
            img2 = _imread_safe(OUTPUT_DIR / r2['full_path'])
            if img1 is not None and img2 is not None:
                off1 = r1['settings'].get('mark_result_offset', 0)
                off2 = r2['settings'].get('mark_result_offset', 0)
                if not images_differ(img1, img2):
                    checks.append({
                        "status": "FAIL",
                        "case": f"{r1['id']} vs {r2['id']}",
                        "message": f"異なるオフセット値({off1}, {off2})なのに画像が同一です"
                    })
                else:
                    checks.append({
                        "status": "PASS",
                        "case": f"{r1['id']} vs {r2['id']}",
                        "message": f"オフセット({off1}, {off2})で描画位置が異なることを確認"
                    })

    # (4) 極端オフセットがクラッシュなしで画像生成されていることを確認
    for extreme_id in ['offset_extreme_left', 'offset_extreme_right']:
        r = next((r for r in results if r['id'] == extreme_id), None)
        if r:
            img = _imread_safe(OUTPUT_DIR / r['full_path'])
            if img is not None and img.shape[0] > 0:
                checks.append({
                    "status": "PASS",
                    "case": extreme_id,
                    "message": "極端オフセットでもクラッシュせず画像が生成されました"
                })
            else:
                checks.append({
                    "status": "FAIL",
                    "case": extreme_id,
                    "message": "極端オフセットで画像生成に失敗"
                })

    # (5) hide_ox と hide_score が異なること
    hide_ox = next((r for r in results if r['id'] == 'hide_ox'), None)
    hide_score = next((r for r in results if r['id'] == 'hide_score'), None)
    if hide_ox and hide_score:
        img1 = _imread_safe(OUTPUT_DIR / hide_ox['full_path'])
        img2 = _imread_safe(OUTPUT_DIR / hide_score['full_path'])
        if img1 is not None and img2 is not None:
            if images_differ(img1, img2):
                checks.append({
                    "status": "PASS",
                    "case": "hide_ox vs hide_score",
                    "message": "○×非表示と得点非表示は異なる画像（独立制御が正しい）"
                })
            else:
                checks.append({
                    "status": "FAIL",
                    "case": "hide_ox vs hide_score",
                    "message": "○×非表示と得点非表示が同一（制御が独立していない）"
                })

    return checks


def generate_html(results, checks, scoring_result):
    """HTMLレポートを生成"""
    pass_count = sum(1 for c in checks if c['status'] == 'PASS')
    fail_count = sum(1 for c in checks if c['status'] == 'FAIL')

    html_parts = []
    html_parts.append(f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>④ 採点結果描画 詳細設定 ビジュアルテストレポート</title>
<style>
body {{ font-family: 'Yu Gothic UI', sans-serif; background: #f5f7fa; color: #333; padding: 20px; max-width: 1800px; margin: 0 auto; }}
h1 {{ color: #1976D2; border-bottom: 3px solid #1976D2; padding-bottom: 10px; }}
h2 {{ color: #546E7A; margin-top: 40px; }}
h3 {{ color: #37474F; }}
.summary {{ background: #fff; border-radius: 8px; padding: 20px; margin: 20px 0; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
.pass {{ color: #2E7D32; font-weight: bold; }}
.fail {{ color: #C62828; font-weight: bold; }}
.case {{ background: #fff; border-radius: 8px; padding: 15px; margin: 15px 0; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border-left: 4px solid #1976D2; }}
.case.offset {{ border-left-color: #F57C00; }}
.case.hide {{ border-left-color: #7B1FA2; }}
.case.combined {{ border-left-color: #00796B; }}
.case-header {{ display: flex; justify-content: space-between; align-items: center; }}
.settings {{ background: #ECEFF1; padding: 5px 10px; border-radius: 4px; font-family: 'Consolas', monospace; font-size: 12px; color: #455A64; }}
.diff-info {{ font-size: 12px; color: #999; }}
.diff-info.changed {{ color: #2E7D32; }}
.images {{ display: flex; gap: 15px; margin-top: 10px; flex-wrap: wrap; }}
.images img {{ border: 1px solid #ddd; border-radius: 4px; cursor: pointer; transition: transform 0.2s; }}
.images img:hover {{ transform: scale(1.02); box-shadow: 0 4px 8px rgba(0,0,0,0.2); }}
.crop-img {{ max-width: 910px; }}
.wide-img {{ max-width: 1170px; }}
.check-list {{ list-style: none; padding: 0; }}
.check-list li {{ padding: 6px 12px; margin: 4px 0; border-radius: 4px; }}
.check-list li.pass-item {{ background: #E8F5E9; }}
.check-list li.fail-item {{ background: #FFEBEE; }}
table {{ border-collapse: collapse; width: 100%; margin: 10px 0; }}
th, td {{ border: 1px solid #ddd; padding: 8px 12px; text-align: left; }}
th {{ background: #ECEFF1; }}
.scoring-info {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px; }}
.scoring-card {{ background: #E3F2FD; padding: 12px; border-radius: 6px; text-align: center; }}
.scoring-number {{ font-size: 28px; font-weight: bold; color: #1565C0; }}
details {{ margin: 5px 0; }}
summary {{ cursor: pointer; color: #1976D2; }}
</style>
</head>
<body>

<h1>④ 採点結果描画 詳細設定 — ビジュアルテストレポート</h1>
<p>生成日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>

<div class="summary">
<h2>テスト概要</h2>
<div class="scoring-info">
<div class="scoring-card">
    <div>テストケース</div>
    <div class="scoring-number">{len(results)}</div>
</div>
<div class="scoring-card">
    <div>自動検証</div>
    <div class="scoring-number">{len(checks)}</div>
</div>
<div class="scoring-card">
    <div class="pass">PASS</div>
    <div class="scoring-number" style="color:#2E7D32">{pass_count}</div>
</div>
<div class="scoring-card">
    <div class="fail">FAIL</div>
    <div class="scoring-number" style="color:#C62828">{fail_count}</div>
</div>
</div>
<p>採点結果: {scoring_result['total_score']}/{scoring_result['max_score']}点 / 
正答{sum(1 for r in scoring_result['results'].values() if r['correct'])}問 / 
誤答{sum(1 for r in scoring_result['results'].values() if not r['correct'])}問</p>
</div>

<h2>🔍 AI自動検証結果</h2>
<ul class="check-list">
""")

    for ck in checks:
        cls = "pass-item" if ck['status'] == 'PASS' else "fail-item"
        icon = "✅" if ck['status'] == 'PASS' else "❌"
        html_parts.append(f'<li class="{cls}">{icon} [{ck["case"]}] {ck["message"]}</li>')

    html_parts.append("</ul>")

    # 各テストケース
    html_parts.append("<h2>📋 テストケース一覧</h2>")

    for r in results:
        # CSSクラス
        css_cls = "case"
        if "offset" in r['id']:
            css_cls += " offset"
        elif "hide" in r['id'] or "only" in r['id'] or "hidden" in r['id']:
            css_cls += " hide"
        elif "combined" in r['id']:
            css_cls += " combined"

        diff_cls = "diff-info changed" if r['differs_from_default'] else "diff-info"
        diff_text = f"差分: {r['diff_pixels']:,}px" if r['id'] != 'default' else "基準画像"

        settings_str = json.dumps(r['settings'], ensure_ascii=False) if r['settings'] else "(デフォルト)"

        html_parts.append(f"""
<div class="{css_cls}">
<div class="case-header">
    <h3>{r['label']} <small style="color:#999">({r['id']})</small></h3>
    <span class="{diff_cls}">{diff_text}</span>
</div>
<p>{r['desc']}</p>
<div class="settings">{settings_str}</div>
<div class="images">
    <div>
        <p><strong>問題1-5 拡大:</strong></p>
        <img src="{r['crop_path']}" class="crop-img" alt="{r['label']} crop">
    </div>
</div>
<details>
    <summary>▶ 問題1-15 広域表示</summary>
    <img src="{r['wide_path']}" class="wide-img" alt="{r['label']} wide">
</details>
<details>
    <summary>▶ フルサイズ画像</summary>
    <img src="{r['full_path']}" style="max-width:100%" alt="{r['label']} full">
</details>
</div>
""")

    # 比較セクション
    html_parts.append("""
<h2>🔄 並列比較</h2>
<p>各行の左がデフォルト、右が変更後です。問題1-5のクロップを比較します。</p>
<table>
<tr><th>設定</th><th>デフォルト</th><th>変更後</th><th>差分ピクセル</th></tr>
""")

    default_crop = next((r for r in results if r['id'] == 'default'), None)
    for r in results:
        if r['id'] == 'default':
            continue
        html_parts.append(f"""
<tr>
<td><strong>{r['label']}</strong><br><small>{r['desc']}</small></td>
<td><img src="{default_crop['crop_path']}" style="max-width:455px"></td>
<td><img src="{r['crop_path']}" style="max-width:455px"></td>
<td>{r['diff_pixels']:,}px</td>
</tr>
""")

    html_parts.append("</table>")

    # 設定一覧テーブル
    html_parts.append("""
<h2>📊 設定パラメータ一覧</h2>
<table>
<tr>
<th>ID</th><th>offset</th>
<th>show_ox</th><th>show_score</th><th>show_aspect</th><th>show_correct</th>
<th>画像差分</th>
</tr>
""")
    for r in results:
        s = get_rendering_settings(r['settings'])
        status = "✅ 差分あり" if r['differs_from_default'] else ("基準" if r['id'] == 'default' else "⚠️ 差分なし")
        html_parts.append(f"""<tr>
<td>{r['id']}</td>
<td>{s['mark_result_offset']}</td>
<td>{'✓' if s['show_ox_mark'] else '✗'}</td>
<td>{'✓' if s['show_score'] else '✗'}</td>
<td>{'✓' if s['show_aspect'] else '✗'}</td>
<td>{'✓' if s['show_correct_answer'] else '✗'}</td>
<td>{status}</td>
</tr>""")

    html_parts.append("</table>")

    html_parts.append("""
<hr>
<p style="color:#999; font-size:12px;">
④ 採点結果描画 詳細設定 ビジュアルテスト — feature/detailed-settings-window ブランチ
</p>
</body>
</html>
""")

    html_path = OUTPUT_DIR / "index.html"
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write("".join(html_parts))

    return html_path


def main():
    print("=" * 60)
    print("④ 採点結果描画 詳細設定 — ビジュアルテスト")
    print("=" * 60)

    # データ読み込み
    template_dict, coordinates, mark2_results = load_test_data()

    # テスト画像生成
    results, scoring_result = generate_test_images(template_dict, coordinates, mark2_results)

    # AI精読検証
    print("\n自動検証中...")
    checks = validate_results(results)

    pass_count = sum(1 for c in checks if c['status'] == 'PASS')
    fail_count = sum(1 for c in checks if c['status'] == 'FAIL')
    print(f"  検証結果: PASS={pass_count}, FAIL={fail_count}")

    for c in checks:
        icon = "✅" if c['status'] == 'PASS' else "❌"
        print(f"  {icon} [{c['case']}] {c['message']}")

    # HTML生成
    html_path = generate_html(results, checks, scoring_result)
    print(f"\n✓ HTMLレポート生成: {html_path}")
    print(f"  テストケース: {len(results)}")
    print(f"  自動検証: {len(checks)} (PASS={pass_count}, FAIL={fail_count})")

    if fail_count > 0:
        print("\n⚠️ 検証に失敗したケースがあります!")
        sys.exit(1)
    else:
        print("\n✅ 全検証PASS — 詳細設定の変更が正しく画像に反映されています")

    return html_path


if __name__ == "__main__":
    main()
