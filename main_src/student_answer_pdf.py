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
_GAP = 8          # 画像間の隙間 (pt) — 見やすさのため少し広め
_HEADER_H = 20    # ヘッダー高さ (pt)
_FOOTER_H = 14    # フッター高さ (pt)
_CAPTION_H = 12   # キャプション高さ (pt)
_CAPTION_FONT_SIZE = 7
_HEADER_FONT_SIZE = 10
_FOOTER_FONT_SIZE = 7

# グリッドレイアウトの制約
_SCALE_UPPER = 1.3  # 表示スケールの軟上限（これを超えたら列数を増やす）
_MAX_COLS = 6       # 列数の上限

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
    log_callback=None,
) -> Optional[Image.Image]:
    """
    1枚の画像から指定領域を切り出して PIL.Image で返す。

    region は 00_Processing 画像 (595×842) に対するピクセル座標 [x1, y1, x2, y2]。
    original_folder が指定されていれば高解像度元画像から切り出す（マーカーがある場合のみ補正）。
    日本語パスを含む環境でも動作するよう PIL.Image.open をベースにしている。
    """
    filename = Path(image_path).name

    def _pil_crop(pil_img: Image.Image, scale: float = 1.0) -> Optional[Image.Image]:
        """PIL画像から region をクロップ（スケール適用済み座標でクリッピング）"""
        img_w, img_h = pil_img.size
        x1 = max(0, min(int(region[0] * scale), img_w))
        y1 = max(0, min(int(region[1] * scale), img_h))
        x2 = max(0, min(int(region[2] * scale), img_w))
        y2 = max(0, min(int(region[3] * scale), img_h))
        if x1 >= x2 or y1 >= y2:
            return None
        return pil_img.crop((x1, y1, x2, y2)).copy()

    # ── 高解像度元画像からの切り出し（オプション） ────────────────
    if original_folder:
        orig_path = Path(original_folder) / filename
        if orig_path.exists():
            try:
                from omr_engine import (
                    detect_corner_markers,
                    apply_perspective_transform,
                    compute_output_scale,
                )
                img_bytes = orig_path.read_bytes()  # Unicode パス対応
                orig_img = cv2.imdecode(
                    np.frombuffer(img_bytes, dtype=np.uint8), cv2.IMREAD_COLOR
                )
                if orig_img is not None:
                    scale = compute_output_scale(orig_img)  # 変換前画像で計算
                    corners = detect_corner_markers(orig_img)  # マーカーなし→ValueError
                    corrected, _ = apply_perspective_transform(
                        orig_img, corners, output_scale=scale
                    )
                    corrected_pil = Image.fromarray(
                        cv2.cvtColor(corrected, cv2.COLOR_BGR2RGB)
                    )
                    result = _pil_crop(corrected_pil, scale)
                    if result is not None:
                        return result
            except Exception:
                # マーカーなし or 変換失敗 → 処理済み画像にフォールバック
                pass

    # ── フォールバック: 処理済み (00_boxed) 画像を PIL で直接読む ──
    # PIL.Image.open は日本語パスでも動作する
    try:
        proc_path = Path(processing_folder) / filename
        if not proc_path.exists():
            # image_path が processing_folder 直下にない場合は image_path 自体を試す
            proc_path = Path(image_path)
        if not proc_path.exists():
            return None
        with Image.open(str(proc_path)) as pil_img:
            pil_img.load()  # ファイルを閉じる前に読み込む
            return _pil_crop(pil_img, 1.0)
    except Exception as e:
        msg = f"画像切り出し失敗 {filename}: {e}"
        logger.warning(msg)
        if log_callback:
            log_callback(msg)
        return None


# ── グリッドレイアウト計算 ──────────────────────────────

def _compute_grid(
    region: List[int],
    scan_width: int = _SCAN_A4_W,
    scan_height: int = _SCAN_A4_H,
) -> Tuple[int, int, float, float]:
    """
    領域サイズから最適な列数 (cols) と表示サイズ (img_w, img_h) を返す。

    <決定方針>
    スキャン原稿 (A4) における記述欄の「物理サイズ」を基準とし、
    次の優先順位で列数を決定する:

      ① scale ∈ [1.0, _SCALE_UPPER] を満たす最大列数を採用する
         → 1枚当たりの画像を最大化しつつ, 自然サイズ以上を確保

      ② ①を満たす整数が存在しない（物理サイズが大きく1列でも上限超え）場合:
         _SCALE_UPPER を軟上限として守り, scale < 1.0 をやむなく許容する
         （2倍超えを避けることを優先）

    <サイジング>
      img_w = (usable_w - (cols-1) * GAP) / cols  ← 横いっぱいに均等配置
      img_h = img_h_natural * (img_w / img_w_natural)  ← アスペクト比維持
    """
    rw = abs(region[2] - region[0])
    rh = abs(region[3] - region[1])
    if rw <= 0 or rh <= 0:
        return 1, 1, 100.0, 100.0

    # 物理 1:1 サイズ (pt) — スキャン画像の A4 比率を PDF ポイントに写像
    img_w_natural = (rw / scan_width)  * _A4_W
    img_h_natural = (rh / scan_height) * _A4_H

    usable_w = _A4_W - 2 * _MARGIN
    usable_h = _A4_H - 2 * _MARGIN - _HEADER_H - _FOOTER_H

    # ① scale ≥ 1.0 を保てる最大列数
    #    cell_w(c) = (usable_w - (c-1)*GAP) / c  ≥  img_w_natural
    #    → c ≤ (usable_w + GAP) / (img_w_natural + GAP)
    max_cols_ge1 = max(1, int((usable_w + _GAP) / (img_w_natural + _GAP)))

    # ② scale ≤ _SCALE_UPPER を保てる最小列数
    #    cell_w(c) ≤ img_w_natural * _SCALE_UPPER
    #    → c ≥ (usable_w + GAP) / (img_w_natural * _SCALE_UPPER + GAP)
    min_cols_le_upper = math.ceil(
        (usable_w + _GAP) / (img_w_natural * _SCALE_UPPER + _GAP)
    )

    # ③ 選択
    #    有効範囲 [min_cols_le_upper, max_cols_ge1] が存在する
    #      → max_cols_ge1 を採用（最大列数, scale は [1.0, _SCALE_UPPER] 内）
    #    存在しない（max_cols_ge1 < min_cols_le_upper のとき scale<1.0不可避）
    #      → min_cols_le_upper を採用（_SCALE_UPPER を守り, scale<1.0を許容）
    if max_cols_ge1 >= min_cols_le_upper:
        cols = max_cols_ge1
    else:
        cols = min_cols_le_upper

    cols = max(1, min(cols, _MAX_COLS))

    # ④ 幅を列数で均等分割（列間に _GAP を確保）
    img_w = (usable_w - (cols - 1) * _GAP) / cols

    # ⑤ アスペクト比を維持して高さを決定
    scale_factor = img_w / img_w_natural
    img_h = img_h_natural * scale_factor

    # ⑥ 高さ方向の最大行数（キャプション分を含む）
    cell_total_h = img_h + _CAPTION_H + _GAP
    rows = max(1, int((usable_h + _GAP) / cell_total_h))

    return cols, rows, img_w, img_h


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
    log_callback=None,
) -> List[str]:
    """
    採点前の設問別解答一覧 PDF を生成する。

    Args:
        processing_folder: 00_Processing (boxed) フォルダ
        config: descriptive_config (questions 含む)
        output_base_folder: _saiten_grading_results のパス
        original_folder: 元画像フォルダ (高解像度切り出し用, 任意)
        progress_callback: callback(current, total) 進捗通知
        log_callback: callback(message: str) GUI ログ出力用

    Returns:
        生成された PDF パスのリスト
    """
    def _log(msg: str):
        logger.info(msg)
        if log_callback:
            log_callback(msg)

    questions = config.get("questions", [])
    if not questions:
        _log("解答一覧PDF: 問題が未設定です")
        return []

    from name_trimmer import get_image_files
    image_files = get_image_files(processing_folder)
    if not image_files:
        _log("解答一覧PDF: 画像ファイルがありません")
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
                img_path, region, processing_folder, original_folder,
                log_callback=log_callback,
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
            _log(f"  採点前PDF: {Path(result).name} ({len([e for e in entries if e['image']])}画像)")

    _log(f"採点前PDF生成完了: {len(generated)}ファイル")
    return generated


def generate_post_scoring_pdfs(
    processing_folder: str,
    config: dict,
    scores_data: dict,
    output_base_folder: str,
    original_folder: Optional[str] = None,
    progress_callback=None,
    log_callback=None,
) -> List[str]:
    """
    採点後の設問別解答一覧 PDF を生成する (得点の高い順にソート)。

    Args:
        processing_folder: 00_Processing (boxed) フォルダ
        config: descriptive_config (questions 含む)
        scores_data: {"scores": {filename: {question_id: score, ...}, ...}}
        output_base_folder: _saiten_grading_results のパス
        original_folder: 元画像フォルダ (高解像度切り出し用, 任意)
        progress_callback: callback(current, total) 進捗通知
        log_callback: callback(message: str) GUI ログ出力用

    Returns:
        生成された PDF パスのリスト
    """
    def _log(msg: str):
        logger.info(msg)
        if log_callback:
            log_callback(msg)

    questions = config.get("questions", [])
    if not questions:
        _log("解答一覧PDF（採点後）: 問題が未設定です")
        return []

    from name_trimmer import get_image_files
    image_files = get_image_files(processing_folder)
    if not image_files:
        _log("解答一覧PDF（採点後）: 画像ファイルがありません")
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
        region = question["region"]

        entries = []
        for img_path in image_files:
            filename = Path(img_path).name
            pil_img = _crop_region_from_image(
                img_path, region, processing_folder, original_folder,
                log_callback=log_callback,
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
            _log(f"  採点後PDF: {Path(result).name} ({len([e for e in entries if e['image']])}画像)")

    _log(f"採点後PDF生成完了: {len(generated)}ファイル")
    return generated
