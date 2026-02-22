# -*- coding: utf-8 -*-
"""
student_answer_pdf.py — 生徒の設問別解答一覧 PDF 生成モジュール

記述採点の前後に、設問ごとの生徒の解答画像を一覧 PDF として出力する。

  - 採点前 (pre-scoring): ファイル名順でソート、キャプションはファイル名+設問名
  - 採点後 (post-scoring): 得点の高い順でソート、キャプションに得点/配点を追加

出力先:
  _saiten_grading_results/04_Answer_Gallery/010_pre_scoring/
  _saiten_grading_results/04_Answer_Gallery/020_post_scoring/
"""

import logging
import math
import os
import re
import tempfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import cv2
import numpy as np
from PIL import Image

logger = logging.getLogger(__name__)

# ── 定数 ────────────────────────────────────────────────
ANSWER_GALLERY_FOLDER = "04_Answer_Gallery"
PRE_SCORING_SUBFOLDER = "010_pre_scoring"
POST_SCORING_SUBFOLDER = "020_post_scoring"

# レイアウト定数 (単位: ポイント, 1pt = 1/72 inch)
_A4_W = 595.28   # A4 横幅 (pt)
_A4_H = 841.89   # A4 縦幅 (pt)
_MARGIN = 18      # 上下左右マージン (pt)
_GAP = 6          # 画像間の隙間 (pt)
_HEADER_H = 20    # ヘッダー高さ (pt)
_FOOTER_H = 14    # フッター高さ (pt)
_CAPTION_H = 12   # キャプション高さ (pt)
_CAPTION_FONT_SIZE = 7
_HEADER_FONT_SIZE = 10
_FOOTER_FONT_SIZE = 7

# スキャン原稿の元サイズ (px) — A4 前提
_SCAN_A4_W = 595   # 基準幅 (72dpi 換算, 元画像はこれの倍率)
_SCAN_A4_H = 842


# ── ヘルパー ────────────────────────────────────────────

def _natural_sort_key(path_str: str):
    """Windows Explorer 互換の自然順ソートキー"""
    return [
        int(c) if c.isdigit() else c.lower()
        for c in re.split(r'(\d+)', str(path_str))
    ]


def _register_japanese_font():
    """reportlab 用に日本語フォントを登録して名称を返す"""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    for name, path in [
        ('Gothic', r'C:\Windows\Fonts\msgothic.ttc'),
        ('Gothic', r'C:\Windows\Fonts\msmincho.ttc'),
    ]:
        try:
            pdfmetrics.registerFont(TTFont(name, path))
            return name
        except Exception:
            continue
    return 'Helvetica'


# ── 画像切り出し ────────────────────────────────────────

def _crop_region_from_image(
    image_path: str,
    region: List[int],
    processing_folder: str,
    original_folder: Optional[str] = None,
) -> Optional[Image.Image]:
    """
    1枚の画像から指定領域を切り出して PIL.Image で返す。

    region は 00_Processing 画像 (595×842) に対するピクセル座標 [x1, y1, x2, y2]。
    original_folder が指定されていれば高解像度元画像から切り出す。
    """
    filename = Path(image_path).name

    try:
        if original_folder:
            orig_path = Path(original_folder) / filename
            if orig_path.exists():
                from omr_engine import (
                    detect_corner_markers,
                    apply_perspective_transform,
                    compute_output_scale,
                )
                with open(str(orig_path), 'rb') as f:
                    img_bytes = f.read()
                orig_img = cv2.imdecode(
                    np.frombuffer(img_bytes, dtype=np.uint8), cv2.IMREAD_COLOR
                )
                if orig_img is not None:
                    corners = detect_corner_markers(orig_img)
                    if corners:
                        corrected, _ = apply_perspective_transform(orig_img, corners)
                        scale = compute_output_scale(corrected)
                        x1 = int(region[0] * scale)
                        y1 = int(region[1] * scale)
                        x2 = int(region[2] * scale)
                        y2 = int(region[3] * scale)
                        h, w = corrected.shape[:2]
                        x1, y1 = max(0, x1), max(0, y1)
                        x2, y2 = min(w, x2), min(h, y2)
                        cropped = corrected[y1:y2, x1:x2]
                        return Image.fromarray(cv2.cvtColor(cropped, cv2.COLOR_BGR2RGB))

        # フォールバック: Processing 画像を使用
        proc_path = Path(processing_folder) / filename
        if not proc_path.exists():
            return None
        img = cv2.imread(str(proc_path))
        if img is None:
            return None
        x1, y1, x2, y2 = region
        h, w = img.shape[:2]
        x1, y1 = max(0, x1), max(0, y1)
        x2, y2 = min(w, x2), min(h, y2)
        cropped = img[y1:y2, x1:x2]
        return Image.fromarray(cv2.cvtColor(cropped, cv2.COLOR_BGR2RGB))
    except Exception as e:
        logger.warning("画像切り出し失敗 %s: %s", filename, e)
        return None


# ── グリッドレイアウト計算 ──────────────────────────────

def _compute_grid(
    region: List[int],
    scan_width: int = _SCAN_A4_W,
    scan_height: int = _SCAN_A4_H,
) -> Tuple[int, int, float, float]:
    """
    領域サイズから A4 ページに収まるグリッド (cols, rows) と
    各セルの画像表示サイズ (cell_w, cell_h) を計算する。

    スキャン原稿が A4 である前提で、元画像に対する相対サイズを維持する。
    """
    rw = abs(region[2] - region[0])
    rh = abs(region[3] - region[1])
    if rw <= 0 or rh <= 0:
        return 1, 1, 100.0, 100.0

    # 領域の A4 に対する比率
    ratio_w = rw / scan_width
    ratio_h = rh / scan_height

    # PDF 上のサイズ (pt) — A4 に対する相対比率を維持
    img_w_pt = ratio_w * _A4_W
    img_h_pt = ratio_h * _A4_H

    # 使用可能領域
    usable_w = _A4_W - 2 * _MARGIN
    usable_h = _A4_H - 2 * _MARGIN - _HEADER_H - _FOOTER_H

    # 幅方向の最大列数
    cols = max(1, int((usable_w + _GAP) / (img_w_pt + _GAP)))

    # 高さ方向の最大行数 (キャプション分を含む)
    cell_total_h = img_h_pt + _CAPTION_H + _GAP
    rows = max(1, int((usable_h + _GAP) / cell_total_h))

    return cols, rows, img_w_pt, img_h_pt


# ── PDF 生成 ────────────────────────────────────────────

def _generate_question_pdf(
    question: dict,
    image_entries: List[dict],
    output_path: str,
    mode: str = "pre",
) -> Optional[str]:
    """
    1設問分のタイル一覧 PDF を生成する。

    Args:
        question: {"id": "D1", "name": "記述1", "max_score": 5, "region": [...]}
        image_entries: [{"filename": str, "image": PIL.Image, "score": int|None}, ...]
        output_path: 出力PDFパス
        mode: "pre" (採点前) or "post" (採点後)
    Returns:
        生成されたPDFパスまたはNone
    """
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas as rl_canvas
        from reportlab.lib.units import mm
    except ImportError:
        logger.error("reportlab が見つかりません。pip install reportlab でインストールしてください。")
        return None

    if not image_entries:
        return None

    region = question["region"]
    q_name = question.get("name", question["id"])
    max_score = question.get("max_score", 0)
    cols, rows, img_w, img_h = _compute_grid(region)
    items_per_page = cols * rows

    font_name = _register_japanese_font()
    total_pages = math.ceil(len(image_entries) / items_per_page)

    c = rl_canvas.Canvas(str(output_path), pagesize=A4)

    for page_idx in range(total_pages):
        start = page_idx * items_per_page
        page_items = image_entries[start:start + items_per_page]

        # ── ヘッダー ──
        c.setFont(font_name, _HEADER_FONT_SIZE)
        header_y = _A4_H - _MARGIN - _HEADER_H + 4
        c.drawString(_MARGIN, header_y, f"設問: {q_name}")

        # ── フッター ──
        c.setFont(font_name, _FOOTER_FONT_SIZE)
        footer_text = f"{page_idx + 1} / {total_pages}"
        c.drawRightString(_A4_W - _MARGIN, _MARGIN - 2, footer_text)

        # ── 画像タイル ──
        # グリッド全体をページ中央に寄せる
        grid_w = cols * img_w + (cols - 1) * _GAP
        offset_x = _MARGIN + ((_A4_W - 2 * _MARGIN) - grid_w) / 2
        offset_x = max(_MARGIN, offset_x)

        content_top = _A4_H - _MARGIN - _HEADER_H - 4

        for idx, entry in enumerate(page_items):
            col = idx % cols
            row = idx // cols
            x = offset_x + col * (img_w + _GAP)
            cell_h = img_h + _CAPTION_H
            y_top = content_top - row * (cell_h + _GAP)

            pil_img = entry["image"]
            if pil_img is None:
                continue

            # PIL → 一時ファイル → reportlab へ渡す
            tmp = tempfile.NamedTemporaryFile(suffix=".jpg", delete=False)
            try:
                pil_img.convert("RGB").save(tmp, format="JPEG", quality=90)
                tmp.close()
                c.drawImage(
                    tmp.name,
                    x, y_top - img_h,
                    width=img_w, height=img_h,
                    preserveAspectRatio=True,
                    anchor='nw',
                )
            finally:
                try:
                    os.unlink(tmp.name)
                except OSError:
                    pass

            # ── キャプション ──
            c.setFont(font_name, _CAPTION_FONT_SIZE)
            cap_y = y_top - img_h - _CAPTION_H + 2
            fname = Path(entry["filename"]).stem
            if mode == "post" and entry.get("score") is not None:
                caption = f"{fname}  [{entry['score']}/{max_score}]"
            else:
                caption = f"{fname}"
            # キャプションを画像幅に収める
            c.drawString(x, cap_y, caption[:60])

        c.showPage()

    c.save()
    logger.info("PDF生成完了: %s (%d画像, %dページ)", output_path, len(image_entries), total_pages)
    return str(output_path)


# ── 公開API ─────────────────────────────────────────────

def generate_pre_scoring_pdfs(
    processing_folder: str,
    config: dict,
    output_base_folder: str,
    original_folder: Optional[str] = None,
    progress_callback=None,
) -> List[str]:
    """
    採点前の設問別解答一覧 PDF を生成する。

    Args:
        processing_folder: 00_Processing フォルダ
        config: descriptive_config (questions 含む)
        output_base_folder: _saiten_grading_results のパス
        original_folder: 元画像フォルダ (高解像度切り出し用, 任意)
        progress_callback: callback(current, total) 進捗通知

    Returns:
        生成された PDF パスのリスト
    """
    questions = config.get("questions", [])
    if not questions:
        logger.warning("generate_pre_scoring_pdfs: 問題が未設定です")
        return []

    from name_trimmer import get_image_files
    image_files = get_image_files(processing_folder)
    if not image_files:
        logger.warning("generate_pre_scoring_pdfs: 画像ファイルがありません")
        return []

    # 自然順ソート
    image_files = sorted(image_files, key=lambda p: _natural_sort_key(Path(p).name))

    # 出力フォルダ
    out_dir = Path(output_base_folder) / ANSWER_GALLERY_FOLDER / PRE_SCORING_SUBFOLDER
    out_dir.mkdir(parents=True, exist_ok=True)

    total_work = len(questions) * len(image_files)
    done = 0
    generated = []

    for q_idx, question in enumerate(questions):
        q_id = question["id"]
        q_name = question.get("name", q_id)
        region = question["region"]

        entries = []
        for img_path in image_files:
            pil_img = _crop_region_from_image(
                img_path, region, processing_folder, original_folder
            )
            entries.append({
                "filename": Path(img_path).name,
                "image": pil_img,
                "score": None,
            })
            done += 1
            if progress_callback:
                progress_callback(done, total_work)

        # ファイル名: 001_設問名_生徒一覧.pdf
        pdf_name = f"{q_idx + 1:03d}_{q_name}_生徒一覧.pdf"
        pdf_path = out_dir / pdf_name
        result = _generate_question_pdf(question, entries, str(pdf_path), mode="pre")
        if result:
            generated.append(result)

    logger.info("採点前PDF生成完了: %d ファイル → %s", len(generated), out_dir)
    return generated


def generate_post_scoring_pdfs(
    processing_folder: str,
    config: dict,
    scores_data: dict,
    output_base_folder: str,
    original_folder: Optional[str] = None,
    progress_callback=None,
) -> List[str]:
    """
    採点後の設問別解答一覧 PDF を生成する (得点の高い順にソート)。

    Args:
        processing_folder: 00_Processing フォルダ
        config: descriptive_config (questions 含む)
        scores_data: {"scores": {filename: {question_id: score, ...}, ...}}
        output_base_folder: _saiten_grading_results のパス
        original_folder: 元画像フォルダ (高解像度切り出し用, 任意)
        progress_callback: callback(current, total) 進捗通知

    Returns:
        生成された PDF パスのリスト
    """
    questions = config.get("questions", [])
    if not questions:
        logger.warning("generate_post_scoring_pdfs: 問題が未設定です")
        return []

    from name_trimmer import get_image_files
    image_files = get_image_files(processing_folder)
    if not image_files:
        logger.warning("generate_post_scoring_pdfs: 画像ファイルがありません")
        return []

    scores_dict = scores_data.get("scores", {})

    # 出力フォルダ
    out_dir = Path(output_base_folder) / ANSWER_GALLERY_FOLDER / POST_SCORING_SUBFOLDER
    out_dir.mkdir(parents=True, exist_ok=True)

    total_work = len(questions) * len(image_files)
    done = 0
    generated = []

    for q_idx, question in enumerate(questions):
        q_id = question["id"]
        q_name = question.get("name", q_id)
        max_score = question.get("max_score", 0)
        region = question["region"]

        entries = []
        for img_path in image_files:
            filename = Path(img_path).name
            pil_img = _crop_region_from_image(
                img_path, region, processing_folder, original_folder
            )
            score = scores_dict.get(filename, {}).get(q_id)
            entries.append({
                "filename": filename,
                "image": pil_img,
                "score": score,
            })
            done += 1
            if progress_callback:
                progress_callback(done, total_work)

        # 得点の高い順にソート (None は末尾)
        entries.sort(
            key=lambda e: (
                -(e["score"] if e["score"] is not None else -1),
                _natural_sort_key(e["filename"]),
            )
        )

        # ファイル名: 001_設問名_生徒一覧.pdf
        pdf_name = f"{q_idx + 1:03d}_{q_name}_生徒一覧.pdf"
        pdf_path = out_dir / pdf_name
        result = _generate_question_pdf(question, entries, str(pdf_path), mode="post")
        if result:
            generated.append(result)

    logger.info("採点後PDF生成完了: %d ファイル → %s", len(generated), out_dir)
    return generated
