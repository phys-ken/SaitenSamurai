"""make_sample_exam.py — LLM協働採点を試すための合成答案PDFを作る。

実在の生徒は一人も含まない。リポジトリ同梱の Mark2 用紙
（sample_basefile/sample_marksheet.jpg）からマークを消して白紙に戻し、
乱数で決めたマーク・架空の氏名「合成 生徒NN」・記述欄の解答を書き込む。

スキャンらしさを出すため、各ページを少し傾け・ずらし、
いくつかのページは上下逆（180°）や横向き（90°）にして PDF にまとめる。

使い方:
    python -m saiten_llm.tools.make_sample_exam <出力PDF> [--students 12] [--seed 1]

出力 PDF と同じ場所に、答え合わせ用の <名前>_truth.json も書く
（各ページの向き・氏名・マーク・記述の正解。テストが使う）。
"""
import argparse
import json
import random
import sys
from pathlib import Path

import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont

from saiten_llm import REPO_ROOT
from constants import find_japanese_font
from omr_engine import (apply_perspective_transform, find_sheet_markers,
                        imread_unicode, parse_excel_coordinates)

SAMPLE = REPO_ROOT / "sample_basefile"
SHEET = SAMPLE / "sample_marksheet.jpg"
COORD = SAMPLE / "M2-03-002_座標ファイル.xlsx"
SCALE = 2.0            # 595x842 → 1190x1684
SKIP_QUESTIONS = 4     # 座標ファイルの先頭4行は学年・クラス・出席番号

# 記述欄（595x842 の座標系）。用紙右下の空いている所に枠を描く
DESC_BOX = (345, 668, 545, 768)
NAME_POS = (215, 186)

DESC_ANSWERS = [
    ("光合成", "正"),
    ("光合成", "正"),
    ("こうごうせい", "正"),
    ("光合せい", "部分"),
    ("呼吸", "誤"),
    ("", "無答"),
]


def _blank_sheet(coords):
    """サンプル用紙を補正し、マークを消して白紙の用紙に戻す"""
    img = imread_unicode(str(SHEET))
    markers = find_sheet_markers(img)
    sheet, _ = apply_perspective_transform(img, markers, output_scale=SCALE)
    gray = cv2.cvtColor(sheet, cv2.COLOR_BGR2GRAY)
    for c in coords:
        x0, y0 = int((c["x"] - 3) * SCALE), int((c["y"] - 1) * SCALE)
        x1 = int((c["x"] + c["width"] + 3) * SCALE)
        y1 = int((c["y"] + c["height"] + 1) * SCALE)
        cell = sheet[y0:y1, x0:x1]
        cell[gray[y0:y1, x0:x1] < 200] = 255
        cx, cy = (x0 + x1) // 2, (y0 + y1) // 2
        cv2.ellipse(sheet, (cx, cy), ((x1 - x0) // 2 - 2, (y1 - y0) // 2 - 1),
                    0, 0, 360, (170, 170, 170), 1)
    return sheet


def _fill(sheet, c):
    cx = int((c["x"] + c["width"] / 2) * SCALE)
    cy = int((c["y"] + c["height"] / 2) * SCALE)
    cv2.ellipse(sheet, (cx, cy), (int(5 * SCALE), int(7 * SCALE)), 0, 0, 360,
                (55, 55, 55), -1)


def _draw_text(sheet, xy, text, size):
    font_path = find_japanese_font()
    font = (ImageFont.truetype(font_path, int(size * SCALE)) if font_path
            else ImageFont.load_default())
    pil = Image.fromarray(cv2.cvtColor(sheet, cv2.COLOR_BGR2RGB))
    ImageDraw.Draw(pil).text((xy[0] * SCALE, xy[1] * SCALE), text,
                             fill=(30, 30, 40), font=font)
    return cv2.cvtColor(np.array(pil), cv2.COLOR_RGB2BGR)


def _scanify(sheet, rng, rotation, noise=True):
    """少し傾けてずらし、元のサンプルと同じ大きさのページに置く。

    noise=True でスキャンのざらつきを足す（PNG が1枚約3MBになる。
    テストでは容量を食うので切る）
    """
    h, w = sheet.shape[:2]
    angle = rng.uniform(-1.2, 1.2)
    m = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0)
    m[0, 2] += rng.uniform(-15, 15)
    m[1, 2] += rng.uniform(-15, 15)
    page = cv2.warpAffine(sheet, m, (w, h), borderValue=(250, 250, 250))
    if noise:
        page = cv2.add(page, rng_noise(page.shape, rng))
    if rotation == 180:
        page = cv2.rotate(page, cv2.ROTATE_180)
    elif rotation == 90:
        page = cv2.rotate(page, cv2.ROTATE_90_CLOCKWISE)
    return page, round(angle, 2)


def rng_noise(shape, rng):
    seed = rng.randrange(1 << 30)
    return np.random.default_rng(seed).integers(0, 6, shape, dtype=np.uint8)


def make_exam(out_pdf, students=12, seed=1, noise=True):
    import fitz
    rng = random.Random(seed)
    coords, _ = parse_excel_coordinates(str(COORD))
    by_q = {}
    for c in coords:
        by_q.setdefault(c["question_no"], []).append(c)
    for q in by_q:
        by_q[q].sort(key=lambda c: c["x"])
    blank = _blank_sheet(coords)
    x0, y0, x1, y1 = (int(v * SCALE) for v in DESC_BOX)
    cv2.rectangle(blank, (x0, y0), (x1, y1), (40, 40, 40), 2)
    blank = _draw_text(blank, (DESC_BOX[0] + 4, DESC_BOX[1] + 3),
                       "問46 植物が光を使って養分をつくるはたらきを何というか", 6)

    answer_rows = [q for q in sorted(by_q) if q > SKIP_QUESTIONS]
    truth = {"students": [], "skip_questions": SKIP_QUESTIONS,
             "desc_box_595": DESC_BOX, "name_pos_595": NAME_POS}
    doc = fitz.open()
    for i in range(1, students + 1):
        sheet = blank.copy()
        # 学年1・クラスA・出席番号 i（十の位・一の位）
        id_marks = {1: 1, 2: 1, 3: i // 10 or 10, 4: i % 10 or 10}
        for q, k in id_marks.items():
            _fill(sheet, by_q[q][k - 1])
        accuracy = rng.uniform(0.5, 0.95)
        marks = {}
        for idx, q in enumerate(answer_rows):
            correct = idx % 10 + 1
            choice = correct if rng.random() < accuracy else rng.randint(1, 10)
            marks[idx + 1] = choice
            _fill(sheet, by_q[q][choice - 1])
        name = f"合成 生徒{i:02d}"
        sheet = _draw_text(sheet, NAME_POS, name, 12)
        desc_text, desc_truth = DESC_ANSWERS[(i - 1) % len(DESC_ANSWERS)]
        if desc_text:
            sheet = _draw_text(sheet, (DESC_BOX[0] + 30, DESC_BOX[1] + 40), desc_text, 16)
        rotation = 180 if i in (3, 8) else 90 if i == 6 else 0
        page, angle = _scanify(sheet, rng, rotation, noise=noise)
        # スキャナーと同じく JPEG で入れる（PNG だと PDF が1ページ数MBになる）
        ok, jpg = cv2.imencode(".jpg", page, [cv2.IMWRITE_JPEG_QUALITY, 90])
        ph, pw = page.shape[:2]
        p = doc.new_page(width=pw * 72 / 200, height=ph * 72 / 200)
        p.insert_image(p.rect, stream=jpg.tobytes())
        truth["students"].append({
            "page": i, "name": name, "rotation": rotation, "skew": angle,
            "marks": marks, "descriptive": {"text": desc_text, "truth": desc_truth},
        })
    out_pdf = Path(out_pdf)
    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_pdf))
    doc.close()
    truth_path = out_pdf.with_name(out_pdf.stem + "_truth.json")
    truth_path.write_text(json.dumps(truth, ensure_ascii=False, indent=1),
                          encoding="utf-8")
    return out_pdf, truth_path


# --- マーカーの無い普通のプリント（記述のみ） ---------------------------

PLAIN_SIZE = (1190, 1684)                 # 幅・高さ（200dpi の A4 相当）
PLAIN_NAME_BOX = (700, 120, 1100, 200)    # 氏名欄（画素）
PLAIN_QUESTIONS = [
    {"name": "問1", "box": (90, 300, 1100, 560), "text": "植物が光を使って養分をつくるはたらきを何というか",
     "answers": [("光合成", "正"), ("こうごうせい", "正"), ("光合せい", "部分"), ("呼吸", "誤")]},
    {"name": "問2", "box": (90, 640, 1100, 900), "text": "水の化学式を書け",
     "answers": [("H2O", "正"), ("H2O", "正"), ("HO2", "誤"), ("", "無答")]},
    {"name": "問3", "box": (90, 980, 1100, 1240), "text": "地球から最も近い恒星は何か",
     "answers": [("太陽", "正"), ("月", "誤"), ("太陽", "正"), ("たいよう", "正")]},
]


def _text_px(img, xy, text, size):
    """画素の座標で文字を書く（_draw_text は 595x842 の座標系）"""
    font_path = find_japanese_font()
    font = ImageFont.truetype(font_path, size) if font_path else ImageFont.load_default()
    pil = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
    ImageDraw.Draw(pil).text(xy, text, fill=(30, 30, 40), font=font)
    return cv2.cvtColor(np.array(pil), cv2.COLOR_RGB2BGR)


def make_plain_exam(out_pdf, students=8, seed=1, rotate_pages=(3,)):
    """四隅のマーカーが無い、普通のプリントの記述テストを作る。

    一斉にスキャンした想定で、傾き・ずれは小さく（±0.5°・±6画素）、
    rotate_pages のページだけ上下逆にする。
    """
    import fitz
    rng = random.Random(seed)
    w, h = PLAIN_SIZE
    blank = np.full((h, w, 3), 255, np.uint8)
    blank = _text_px(blank, (90, 60), "理科 小テスト（合成データ）", 40)
    x0, y0, x1, y1 = PLAIN_NAME_BOX
    cv2.rectangle(blank, (x0, y0), (x1, y1), (40, 40, 40), 2)
    blank = _text_px(blank, (x0 + 8, y0 + 6), "氏名", 22)
    for q in PLAIN_QUESTIONS:
        bx0, by0, bx1, by1 = q["box"]
        blank = _text_px(blank, (bx0, by0 - 44), f"{q['name']} {q['text']}", 28)
        cv2.rectangle(blank, (bx0, by0), (bx1, by1), (40, 40, 40), 2)
    truth = {"kind": "plain", "size": PLAIN_SIZE, "name_box": PLAIN_NAME_BOX,
             "questions": [{"name": q["name"], "box": q["box"]} for q in PLAIN_QUESTIONS],
             "students": []}
    doc = fitz.open()
    for i in range(1, students + 1):
        sheet = _text_px(blank.copy(), (x0 + 90, y0 + 22), f"合成 生徒{i:02d}", 36)
        answers = {}
        for q in PLAIN_QUESTIONS:
            text, verdict = q["answers"][(i - 1) % len(q["answers"])]
            if text:
                bx0, by0 = q["box"][:2]
                sheet = _text_px(sheet, (bx0 + 60, by0 + 80), text, 64)
            answers[q["name"]] = {"text": text, "truth": verdict}
        angle = rng.uniform(-0.5, 0.5)
        m = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0)
        m[0, 2] += rng.uniform(-6, 6)
        m[1, 2] += rng.uniform(-6, 6)
        page = cv2.warpAffine(sheet, m, (w, h), borderValue=(250, 250, 250))
        rotation = 180 if i in rotate_pages else 0
        if rotation:
            page = cv2.rotate(page, cv2.ROTATE_180)
        ok, jpg = cv2.imencode(".jpg", page, [cv2.IMWRITE_JPEG_QUALITY, 90])
        p = doc.new_page(width=w * 72 / 200, height=h * 72 / 200)
        p.insert_image(p.rect, stream=jpg.tobytes())
        truth["students"].append({"page": i, "name": f"合成 生徒{i:02d}", "rotation": rotation,
                                  "skew": round(angle, 2), "answers": answers})
    out_pdf = Path(out_pdf)
    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_pdf))
    doc.close()
    truth_path = out_pdf.with_name(out_pdf.stem + "_truth.json")
    truth_path.write_text(json.dumps(truth, ensure_ascii=False, indent=1), encoding="utf-8")
    return out_pdf, truth_path


def main(argv=None):
    ap = argparse.ArgumentParser(description="合成答案PDFを作る（実在の生徒は含まない）")
    ap.add_argument("out_pdf")
    ap.add_argument("--students", type=int, default=12)
    ap.add_argument("--seed", type=int, default=1)
    args = ap.parse_args(argv)
    pdf, truth = make_exam(args.out_pdf, args.students, args.seed)
    print(f"作成しました: {pdf}\n答え合わせ用: {truth}")


if __name__ == "__main__":
    sys.exit(main())
