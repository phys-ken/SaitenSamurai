"""
regenerate_docs_screenshots.py — docs/images/ 用のリアルなスクリーンショットを再生成する

ダミー答案画像（手書き風テキスト付き）を生成し、
記述採点 GUI のモックを構築してキャプチャする。
対象:
  - 19_integrated_descriptive_setup.png
  - 06_single_question_scorer.png
  - desc_03_scored_maru.png
  - desc_04_scored_batsu.png
  - desc_05_scored_middle.png
  - desc_06_filter_active.png
  - desc_07_all_scored.png
  - desc_08_grid_mode.png

使用法: python tests/regenerate_docs_screenshots.py
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes as wintypes
import os
import sys
import tempfile
import time
import tkinter as tk
from pathlib import Path
from tkinter import ttk
from typing import Dict, List, Optional, Union

import numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageGrab, ImageTk

# tk.Tk / tk.Toplevel は Wm + Misc 両方を継承しているため、
# winfo_* や update 等のメソッドを Pylance が正しく認識できる型
_TkWindow = Union[tk.Tk, tk.Toplevel]

# ============================================================
# パス設定
# ============================================================

ROOT = Path(__file__).resolve().parent.parent
DOCS_IMAGES = ROOT / "docs" / "images"
assert DOCS_IMAGES.exists(), f"docs/images/ が見つかりません: {DOCS_IMAGES}"


# ============================================================
# Win32 PrintWindow キャプチャ
# ============================================================

class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", wintypes.DWORD),
        ("biWidth", ctypes.c_long),
        ("biHeight", ctypes.c_long),
        ("biPlanes", wintypes.WORD),
        ("biBitCount", wintypes.WORD),
        ("biCompression", wintypes.DWORD),
        ("biSizeImage", wintypes.DWORD),
        ("biXPelsPerMeter", ctypes.c_long),
        ("biYPelsPerMeter", ctypes.c_long),
        ("biClrUsed", wintypes.DWORD),
        ("biClrImportant", wintypes.DWORD),
    ]


def _get_hwnd(widget: _TkWindow) -> Optional[int]:
    try:
        frame_id = widget.wm_frame()
        if frame_id and frame_id != "0x0":
            return int(frame_id, 16)
    except Exception:
        pass
    try:
        return widget.winfo_id()
    except Exception:
        return None


def _capture_with_printwindow(widget: _TkWindow, path: Path) -> bool:
    try:
        hwnd = _get_hwnd(widget)
        if not hwnd:
            return False
        rect = wintypes.RECT()
        ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
        width = rect.right - rect.left
        height = rect.bottom - rect.top
        if width < 10 or height < 10:
            return False
        wDC = ctypes.windll.user32.GetWindowDC(hwnd)
        if not wDC:
            return False
        try:
            dcObj = ctypes.windll.gdi32.CreateCompatibleDC(wDC)
            bmp = ctypes.windll.gdi32.CreateCompatibleBitmap(wDC, width, height)
            old = ctypes.windll.gdi32.SelectObject(dcObj, bmp)
            ok = ctypes.windll.user32.PrintWindow(hwnd, dcObj, 2)
            if ok:
                bmi = BITMAPINFOHEADER()
                bmi.biSize = ctypes.sizeof(BITMAPINFOHEADER)
                bmi.biWidth = width
                bmi.biHeight = -height
                bmi.biPlanes = 1
                bmi.biBitCount = 32
                bmi.biCompression = 0
                buf = ctypes.create_string_buffer(width * height * 4)
                ctypes.windll.gdi32.GetDIBits(
                    dcObj, bmp, 0, height, buf, ctypes.byref(bmi), 0,
                )
                img = Image.frombuffer(
                    "RGBA", (width, height), bytes(buf), "raw", "BGRA", 0, 1,
                )
                img = img.convert("RGB")
                img.save(str(path))
            ctypes.windll.gdi32.SelectObject(dcObj, old)
            ctypes.windll.gdi32.DeleteObject(bmp)
            ctypes.windll.gdi32.DeleteDC(dcObj)
        finally:
            ctypes.windll.user32.ReleaseDC(hwnd, wDC)
        return bool(ok)
    except Exception as e:
        print(f"  [PrintWindow failed] {e}")
        return False


def _capture_with_imagegrab(widget: _TkWindow, path: Path) -> bool:
    try:
        outer_x = widget.winfo_x()
        outer_y = widget.winfo_y()
        client_x = widget.winfo_rootx()
        client_y = widget.winfo_rooty()
        client_w = widget.winfo_width()
        client_h = widget.winfo_height()
        border = client_x - outer_x
        x1, y1 = outer_x, outer_y
        x2 = client_x + client_w + border
        y2 = client_y + client_h + border
        if (x2 - x1) < 10 or (y2 - y1) < 10:
            return False
        img = ImageGrab.grab(bbox=(x1, y1, x2, y2))
        img.save(str(path))
        return True
    except Exception as e:
        print(f"  [ImageGrab fallback failed] {e}")
        return False


def _capture_window(widget: _TkWindow, filename: str,
                    delay_ms: int = 700) -> Optional[Path]:
    """ウィンドウキャプチャ → docs/images/ に保存"""
    widget.update_idletasks()
    widget.update()
    widget.lift()
    widget.update_idletasks()
    widget.update()
    time.sleep(delay_ms / 1000.0)
    widget.update()

    path = DOCS_IMAGES / f"{filename}.png"

    if sys.platform == "win32":
        if _capture_with_printwindow(widget, path):
            print(f"  ✔ {filename}.png (PrintWindow)")
            return path

    if _capture_with_imagegrab(widget, path):
        print(f"  ✔ {filename}.png (ImageGrab)")
        return path

    print(f"  ✗ {filename}.png キャプチャ失敗")
    return None


# ============================================================
# ダミー答案画像の生成 (リアル版)
# ============================================================

# 各問題の手書き風テキスト（生徒ごとに異なる内容）
_ANSWER_TEXTS = {
    "q1": [
        "光合成とは、植物が太陽光のエネルギーを利用して\n二酸化炭素と水からグルコースと酸素を生成する反応である。\n葉緑体のチラコイド膜で光化学反応が起こり、\nストロマでカルビン回路が進行する。",
        "光合成は植物の葉緑体で行われる。\n光エネルギーでCO2とH2Oからブドウ糖を合成する。\n6CO2 + 6H2O → C6H12O6 + 6O2",
        "太陽の光を使って、植物が栄養を作る仕組み。\n葉っぱの中の葉緑体で起こる。",
        "",  # 空欄
        "植物は光合成によって有機物を合成する。\nこの過程では光エネルギーが化学エネルギーに変換される。\nまた、副産物として酸素が放出される。\nこの酸素は動物の呼吸に不可欠である。",
    ],
    "q2": [
        "グラフより、温度が上昇するにつれて\n反応速度も増加するが、40℃を超えると\n急激に低下していることがわかる。\nこれは酵素の失活によるものと考えられる。",
        "温度と反応速度は比例関係ではない。\n最適温度は約37℃付近である。",
        "温度が高いほど反応が速くなるが\nある温度を超えると遅くなる。\n酵素がこわれるから。",
        "グラフの横軸は温度、縦軸は反応速度を表す。\n温度依存性が確認できる。",
        "",  # 空欄
    ],
    "q3": [
        "今回の実験では、酵素カタラーゼの活性を測定した。\n過酸化水素の分解速度を各温度で記録し、\nアレニウスプロットを作成した。結果として、\n活性化エネルギーは約52 kJ/molと算出された。\nまた、最適pHは7.4付近であることが確認された。\n失活温度は55℃であり、生体内の条件に\n適応していることが示唆される。",
        "実験の結果、温度が高いほど反応速度が\n上昇するが、至適温度を超えると急激に\n低下することが確認された。これは酵素\nタンパク質の熱変性が原因である。",
        "カタラーゼの実験をした。\n泡がたくさん出た温度が一番よく働く温度。\n55℃だとほとんど泡が出なかった。",
        "実験結果から、酵素反応には最適な条件が\n存在することがわかった。温度だけでなく\npHも酵素活性に影響を与える重要な因子である。",
        "省略",
    ],
}


def _create_dummy_answer_sheet(idx: int, img_w: int = 600, img_h: int = 850) -> Image.Image:
    """リアルなダミー答案画像を生成する。

    テキストを手書き風に描画し、問題番号・配点・枠線付き。
    """
    names = ["山田太郎", "鈴木花子", "田中一郎", "佐藤美咲", "高橋健太",
             "渡辺愛", "伊藤翔太", "小林由美"]
    name = names[idx % len(names)]

    img = Image.new("RGB", (img_w, img_h), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    # フォント
    try:
        font_title = ImageFont.truetype("msmincho.ttc", 16)
        font_header = ImageFont.truetype("msmincho.ttc", 13)
        font_body = ImageFont.truetype("msmincho.ttc", 11)
        font_small = ImageFont.truetype("msmincho.ttc", 9)
        font_handwrite = ImageFont.truetype("msmincho.ttc", 10)
    except (IOError, OSError):
        font_title = ImageFont.load_default()
        font_header = font_body = font_small = font_handwrite = font_title

    # ── ヘッダー ──
    draw.rectangle([15, 10, img_w - 15, 55], outline=(0, 0, 0), width=2)
    draw.text((25, 14), "理科 期末テスト", fill=(0, 0, 0), font=font_title)
    draw.line([(200, 10), (200, 55)], fill=(0, 0, 0), width=1)
    draw.text((210, 16), f"氏名: {name}", fill=(30, 30, 120), font=font_header)
    draw.line([(420, 10), (420, 55)], fill=(0, 0, 0), width=1)
    draw.text((430, 16), f"番号: {idx + 1:02d}", fill=(30, 30, 120), font=font_header)
    draw.text((210, 36), "_______________", fill=(150, 150, 150), font=font_small)

    # ── 問1 (y: 70-280) ──
    draw.rectangle([15, 65, img_w - 15, 95], fill=(240, 245, 255), outline=(0, 0, 0), width=1)
    draw.text((20, 70), "問1  次の現象について説明しなさい。", fill=(0, 0, 0), font=font_header)
    draw.text((img_w - 65, 72), "(5点)", fill=(100, 100, 100), font=font_small)
    draw.rectangle([15, 95, img_w - 15, 280], outline=(180, 180, 180), width=1)
    # 罫線
    for y in range(115, 275, 20):
        draw.line([(25, y), (img_w - 25, y)], fill=(220, 220, 230), width=1)
    # 手書きテキスト
    answer = _ANSWER_TEXTS["q1"][idx % len(_ANSWER_TEXTS["q1"])]
    if answer:
        np.random.seed(idx * 1000 + 1)
        for li, line in enumerate(answer.split("\n")):
            y = 100 + li * 20
            if y > 265:
                break
            x = 25 + np.random.randint(0, 5)
            # 手書き風: わずかに揺れた色
            r, g, b = 20 + np.random.randint(0, 30), 20 + np.random.randint(0, 15), 60 + np.random.randint(0, 30)
            draw.text((x, y), line, fill=(r, g, b), font=font_handwrite)

    # ── 問2 (y: 295-510) ──
    draw.rectangle([15, 290, img_w - 15, 320], fill=(240, 245, 255), outline=(0, 0, 0), width=1)
    draw.text((20, 295), "問2  以下のグラフから読み取れることを述べよ。", fill=(0, 0, 0), font=font_header)
    draw.text((img_w - 65, 297), "(5点)", fill=(100, 100, 100), font=font_small)
    draw.rectangle([15, 320, img_w - 15, 510], outline=(180, 180, 180), width=1)
    for y in range(340, 505, 20):
        draw.line([(25, y), (img_w - 25, y)], fill=(220, 220, 230), width=1)
    answer = _ANSWER_TEXTS["q2"][idx % len(_ANSWER_TEXTS["q2"])]
    if answer:
        np.random.seed(idx * 1000 + 2)
        for li, line in enumerate(answer.split("\n")):
            y = 325 + li * 20
            if y > 495:
                break
            x = 25 + np.random.randint(0, 5)
            r, g, b = 20 + np.random.randint(0, 30), 20 + np.random.randint(0, 15), 60 + np.random.randint(0, 30)
            draw.text((x, y), line, fill=(r, g, b), font=font_handwrite)

    # ── 問3 (y: 525-835) ──
    draw.rectangle([15, 520, img_w - 15, 550], fill=(240, 245, 255), outline=(0, 0, 0), width=1)
    draw.text((20, 525), "問3  実験結果のまとめを書きなさい。", fill=(0, 0, 0), font=font_header)
    draw.text((img_w - 65, 527), "(10点)", fill=(100, 100, 100), font=font_small)
    draw.rectangle([15, 550, img_w - 15, 835], outline=(180, 180, 180), width=1)
    for y in range(570, 830, 20):
        draw.line([(25, y), (img_w - 25, y)], fill=(220, 220, 230), width=1)
    answer = _ANSWER_TEXTS["q3"][idx % len(_ANSWER_TEXTS["q3"])]
    if answer:
        np.random.seed(idx * 1000 + 3)
        for li, line in enumerate(answer.split("\n")):
            y = 555 + li * 20
            if y > 820:
                break
            x = 25 + np.random.randint(0, 5)
            r, g, b = 20 + np.random.randint(0, 30), 20 + np.random.randint(0, 15), 60 + np.random.randint(0, 30)
            draw.text((x, y), line, fill=(r, g, b), font=font_handwrite)

    return img


def _create_dummy_sheets_in_folder(folder: Path, count: int = 8) -> Dict[str, str]:
    """フォルダにダミー答案を保存し、{filename: path} を返す"""
    paths: Dict[str, str] = {}
    for i in range(count):
        img = _create_dummy_answer_sheet(i)
        fn = f"student_{i + 1:03d}.jpg"
        p = folder / fn
        img.save(str(p), "JPEG", quality=92)
        paths[fn] = str(p)
    return paths


def _crop_region(img_path: str, region: list) -> Image.Image:
    """答案画像から指定領域を切り出す"""
    img = Image.open(img_path)
    x1, y1, x2, y2 = region
    return img.crop((x1, y1, x2, y2))


# ============================================================
# 19. 統合セットアップ (答案画像をキャンバスに描画)
# ============================================================

def capture_19_integrated_setup(root: tk.Tk, sheet_path: str):
    """19_integrated_descriptive_setup.png — 答案画像付きの統合セットアップ"""
    win = tk.Toplevel(root)
    win.title("📝 記述問題の設定")
    win.geometry("1040x720")
    win.configure(bg="#F5F7FA")
    BG = "#F5F7FA"

    main = tk.Frame(win, bg=BG, padx=8, pady=8)
    main.pack(fill=tk.BOTH, expand=True)

    # ── 左: Canvas に答案画像を表示 ──
    left = tk.Frame(main, bg=BG)
    left.pack(side=tk.LEFT, fill=tk.BOTH)
    tk.Label(left, text="答案画像（ドラッグで領域を選択）",
             font=("Yu Gothic UI", 10, "bold"), bg=BG, fg="#333").pack(anchor=tk.W, pady=(0, 3))
    canvas = tk.Canvas(left, width=500, height=600, bg="white",
                       highlightthickness=1, highlightbackground="#999", cursor="crosshair")
    canvas.pack()

    # 答案画像を Canvas に表示
    pil_img = Image.open(sheet_path)
    # Canvas サイズに合わせてリサイズ
    cw, ch = 500, 600
    ratio = min(cw / pil_img.width, ch / pil_img.height)
    new_w = int(pil_img.width * ratio)
    new_h = int(pil_img.height * ratio)
    pil_img_resized = pil_img.resize((new_w, new_h), Image.Resampling.LANCZOS)

    # 半透明のカラー領域をオーバーレイ (RGBA合成)
    overlay = pil_img_resized.convert("RGBA")
    draw_overlay = ImageDraw.Draw(overlay)

    # D1 領域 (問1の解答欄: 赤枠)
    d1_y1 = int(95 * ratio)
    d1_y2 = int(280 * ratio)
    d1_x1 = int(15 * ratio)
    d1_x2 = int(585 * ratio)
    draw_overlay.rectangle([d1_x1, d1_y1, d1_x2, d1_y2],
                           fill=(255, 100, 100, 40), outline=(255, 100, 100, 200), width=3)
    try:
        font_label = ImageFont.truetype("msmincho.ttc", 12)
    except (IOError, OSError):
        font_label = ImageFont.load_default()
    draw_overlay.text((d1_x1 + 4, d1_y1 + 2), "D1: 問1 説明問題",
                      fill=(220, 50, 50, 220), font=font_label)

    # D2 領域 (問2の解答欄: 青枠)
    d2_y1 = int(320 * ratio)
    d2_y2 = int(510 * ratio)
    draw_overlay.rectangle([d1_x1, d2_y1, d1_x2, d2_y2],
                           fill=(100, 100, 255, 40), outline=(100, 100, 255, 200), width=3)
    draw_overlay.text((d1_x1 + 4, d2_y1 + 2), "D2: 問2 グラフ読み取り",
                      fill=(50, 50, 220, 220), font=font_label)

    # D3 領域 (問3の解答欄: 緑枠)
    d3_y1 = int(550 * ratio)
    d3_y2 = int(835 * ratio)
    draw_overlay.rectangle([d1_x1, d3_y1, d1_x2, d3_y2],
                           fill=(100, 200, 100, 40), outline=(100, 200, 100, 200), width=3)
    draw_overlay.text((d1_x1 + 4, d3_y1 + 2), "D3: 問3 実験まとめ",
                      fill=(30, 150, 30, 220), font=font_label)

    # RGB に変換して Canvas に表示
    composite = Image.alpha_composite(
        Image.new("RGBA", overlay.size, (255, 255, 255, 255)), overlay
    ).convert("RGB")
    tk_img = ImageTk.PhotoImage(composite)
    canvas.create_image(cw // 2, ch // 2, image=tk_img, anchor=tk.CENTER)
    setattr(canvas, "_tk_img_ref", tk_img)

    tk.Label(left, text="💡 ドラッグで新しい記述領域を追加できます",
             font=("Yu Gothic UI", 8), bg=BG, fg="#777").pack(anchor=tk.W, pady=(3, 0))

    # ── 右: テーブル ──
    right = tk.Frame(main, bg=BG, padx=10)
    right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    tk.Label(right, text="設問一覧",
             font=("Yu Gothic UI", 11, "bold"), bg=BG, fg="#333").pack(anchor=tk.W, pady=(0, 5))

    cols = ("id", "name", "score", "aspect")
    tree = ttk.Treeview(right, columns=cols, show="headings", height=10,
                        selectmode="browse")
    tree.heading("id", text="#")
    tree.heading("name", text="問題名")
    tree.heading("score", text="配点")
    tree.heading("aspect", text="観点")
    tree.column("id", width=40, anchor="center")
    tree.column("name", width=140)
    tree.column("score", width=50, anchor="center")
    tree.column("aspect", width=50, anchor="center")
    tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    tree.insert("", tk.END, values=("D1", "問1 説明問題", 5, 1))
    tree.insert("", tk.END, values=("D2", "問2 グラフ読み取り", 5, 1))
    tree.insert("", tk.END, values=("D3", "問3 実験まとめ", 10, 2))

    btn_frame = tk.Frame(right, bg=BG)
    btn_frame.pack(fill=tk.X, pady=(8, 0))

    tk.Button(btn_frame, text="🗑 選択行を削除",
              font=("Yu Gothic UI", 9), bg="#FFCDD2", relief=tk.FLAT,
              cursor="hand2").pack(side=tk.LEFT, padx=(0, 5))
    tk.Label(btn_frame, text="", bg=BG).pack(side=tk.LEFT, expand=True)
    tk.Button(btn_frame, text="キャンセル",
              font=("Yu Gothic UI", 9), bg="#E0E0E0", relief=tk.FLAT,
              cursor="hand2").pack(side=tk.RIGHT, padx=(5, 0))
    tk.Button(btn_frame, text="✔ 設定を保存",
              font=("Yu Gothic UI", 10, "bold"), bg="#81C784", fg="white",
              relief=tk.FLAT, cursor="hand2").pack(side=tk.RIGHT)

    tk.Label(right, text="登録済み: 3 問  |  合計配点: 20 点",
             bg=BG, font=("Yu Gothic UI", 9), fg="#555").pack(anchor=tk.W, pady=(5, 0))

    win.update_idletasks()
    win.update()
    _capture_window(win, "19_integrated_descriptive_setup")
    win.destroy()


# ============================================================
# 06. SingleQuestionScorer (答案の切り出し画像付き)
# ============================================================

def capture_06_single_question_scorer(root: tk.Tk, image_paths: Dict[str, str]):
    """06_single_question_scorer.png — 答案の切り出し画像を表示した採点画面"""
    win = tk.Toplevel(root)
    win.title("採点: 問1 説明問題 (配点:5点)")
    win.geometry("1000x700+80+80")
    win.resizable(True, True)

    container = tk.Frame(win)
    container.pack(fill=tk.BOTH, expand=True)

    single_frame = tk.Frame(container)
    single_frame.pack(fill=tk.BOTH, expand=True)

    # ── top_bar (ダークヘッダー) ──
    top_bar = tk.Frame(single_frame, bg="#37474F", padx=6, pady=4)
    top_bar.pack(fill=tk.X)

    tk.Label(top_bar, text="1 / 8",
             font=("Yu Gothic UI", 14, "bold"),
             bg="#37474F", fg="#FFD54F").pack(side=tk.LEFT, padx=(0, 12))

    tk.Label(top_bar, text="student_001.jpg",
             font=("Yu Gothic UI", 8),
             bg="#37474F", fg="#B0BEC5").pack(side=tk.LEFT, padx=(0, 10))

    tk.Label(top_bar, text="得点:",
             font=("Yu Gothic UI", 9),
             bg="#37474F", fg="#FFD54F").pack(side=tk.LEFT)

    tk.Label(top_bar, text="—",
             font=("Yu Gothic UI", 16, "bold"),
             bg="#37474F", fg="#FFD54F").pack(side=tk.LEFT, padx=(2, 10))

    # ○ / × ボタン
    tk.Button(top_bar, text="〇 正解(5点)",
              bg="#E3F2FD", fg="#1565C0",
              font=("Yu Gothic UI", 9, "bold"),
              relief=tk.RAISED, cursor="hand2",
              activebackground="#BBDEFB",
              padx=6, pady=1).pack(side=tk.LEFT, padx=2)
    tk.Button(top_bar, text="× 不正解(0点)",
              bg="#FFEBEE", fg="#C62828",
              font=("Yu Gothic UI", 9, "bold"),
              relief=tk.RAISED, cursor="hand2",
              activebackground="#FFCDD2",
              padx=6, pady=1).pack(side=tk.LEFT, padx=2)

    # 右寄せ
    tk.Button(top_bar, text="キャンセル",
              font=("Yu Gothic UI", 8),
              bg="#37474F", fg="#B0BEC5",
              relief=tk.FLAT, cursor="hand2").pack(side=tk.RIGHT, padx=2)
    tk.Button(top_bar, text="✔ 採点完了",
              bg="#4CAF50", fg="white",
              font=("Yu Gothic UI", 9, "bold"),
              relief=tk.FLAT, cursor="hand2",
              padx=8).pack(side=tk.RIGHT, padx=2)

    # ── mid_bar ──
    mid_bar = tk.Frame(single_frame, bg="#ECEFF1", padx=6, pady=2)
    mid_bar.pack(fill=tk.X)

    tk.Label(mid_bar, text="🔍",
             font=("Yu Gothic UI", 9), bg="#ECEFF1").pack(side=tk.LEFT)
    zoom_slider = tk.Scale(mid_bar, from_=25, to=300,
                           orient=tk.HORIZONTAL, length=120,
                           sliderlength=14, bg="#ECEFF1",
                           highlightthickness=0, showvalue=False)
    zoom_slider.set(100)
    zoom_slider.pack(side=tk.LEFT, padx=(0, 2))
    tk.Label(mid_bar, text="100%",
             font=("Yu Gothic UI", 8), fg="#555",
             bg="#ECEFF1").pack(side=tk.LEFT, padx=(0, 8))

    tk.Checkbutton(mid_bar, text="未採点のみ",
                   font=("Yu Gothic UI", 8),
                   bg="#ECEFF1").pack(side=tk.LEFT, padx=(0, 5))

    tk.Label(mid_bar, text="未採点: 8件",
             font=("Yu Gothic UI", 8), fg="#E65100",
             bg="#ECEFF1").pack(side=tk.LEFT, padx=(0, 10))

    tk.Label(mid_bar, text="|", fg="#ccc", bg="#ECEFF1",
             font=("Yu Gothic UI", 9)).pack(side=tk.LEFT, padx=(0, 3))
    tk.Label(mid_bar, text="入力可能:",
             font=("Yu Gothic UI", 8), bg="#ECEFF1").pack(side=tk.LEFT)
    for i in range(6):
        var = tk.BooleanVar(value=True)
        tk.Checkbutton(mid_bar, text=str(i), variable=var,
                       font=("Yu Gothic UI", 8),
                       bg="#ECEFF1").pack(side=tk.LEFT, padx=1)

    tk.Label(mid_bar, text="📷元画像を開く",
             font=("Yu Gothic UI", 8, "underline"),
             fg="#1976D2", bg="#ECEFF1",
             cursor="hand2").pack(side=tk.RIGHT, padx=5)
    tk.Label(mid_bar, text="❓操作方法",
             font=("Yu Gothic UI", 8, "underline"),
             fg="#1976D2", bg="#ECEFF1",
             cursor="hand2").pack(side=tk.RIGHT, padx=5)

    # ── Canvas エリア (切り出し画像を表示) ──
    canvas_frame = tk.Frame(single_frame)
    canvas_frame.pack(fill=tk.BOTH, expand=True, padx=4, pady=2)
    tk.Scrollbar(canvas_frame, orient=tk.VERTICAL).pack(side=tk.RIGHT, fill=tk.Y)
    tk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL).pack(side=tk.BOTTOM, fill=tk.X)
    canvas = tk.Canvas(canvas_frame, bg="white",
                       highlightthickness=1,
                       highlightbackground="#ccc")
    canvas.pack(fill=tk.BOTH, expand=True)

    win.update_idletasks()
    win.update()

    # 切り出し画像を表示（問1の領域: y=95..280）
    first_fn = sorted(image_paths.keys())[0]
    cropped = _crop_region(image_paths[first_fn], [15, 95, 585, 280])

    # Canvas サイズに合わせて拡大
    canvas.update_idletasks()
    cw = max(canvas.winfo_width(), 600)
    ch = max(canvas.winfo_height(), 400)
    ratio = min(cw / cropped.width, ch / cropped.height) * 0.9
    new_w = int(cropped.width * ratio)
    new_h = int(cropped.height * ratio)
    cropped_resized = cropped.resize((new_w, new_h), Image.Resampling.LANCZOS)
    tk_img = ImageTk.PhotoImage(cropped_resized)
    canvas.create_image(cw // 2, ch // 2, image=tk_img, anchor=tk.CENTER)
    setattr(canvas, "_tk_img_ref", tk_img)

    win.update_idletasks()
    win.update()
    _capture_window(win, "06_single_question_scorer")
    win.destroy()


# ============================================================
# desc_03〜desc_07: 採点状態の各キャプチャ
# ============================================================

def _build_descriptive_scorer_window(root, image_paths, student_idx=0,
                                     score_text="—", bg_color="white",
                                     max_score=5, progress="1 / 8",
                                     unscored_text="未採点: 8人",
                                     filter_on=False, use_entry=False):
    """採点画面のモック（実UIに近い構造）"""
    win = tk.Toplevel(root)
    win.title(f"採点: 問1 説明問題 (配点:{max_score}点)")
    win.geometry("1000x700+80+80")
    win.resizable(True, True)

    container = tk.Frame(win)
    container.pack(fill=tk.BOTH, expand=True)

    main_frame = tk.Frame(container)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # ── 左: Canvas ──
    canvas_frame = tk.Frame(main_frame)
    canvas_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    canvas = tk.Canvas(canvas_frame, bg=bg_color,
                       highlightthickness=1, highlightbackground="#ccc")
    canvas.pack(fill=tk.BOTH, expand=True)

    # ── 右: 操作パネル ──
    panel = tk.Frame(main_frame, width=220, padx=10)
    panel.pack(side=tk.RIGHT, fill=tk.Y)
    panel.pack_propagate(False)

    progress_var = tk.StringVar(value=progress)
    tk.Label(panel, textvariable=progress_var,
             font=("Yu Gothic UI", 11, "bold")).pack(pady=(0, 5))

    fns = sorted(image_paths.keys())
    fn_var = tk.StringVar(value=fns[student_idx] if student_idx < len(fns) else "")
    tk.Label(panel, textvariable=fn_var,
             font=("Yu Gothic UI", 8), fg="gray", wraplength=200).pack(pady=(0, 10))

    tk.Label(panel, text="得点:", font=("Yu Gothic UI", 10)).pack()
    score_var = tk.StringVar(value=score_text)
    tk.Label(panel, textvariable=score_var,
             font=("Yu Gothic UI", 28, "bold"), fg="#1976D2").pack(pady=5)

    # かんたん採点
    tk.Label(panel, text="─" * 20, fg="#ccc").pack(pady=5)
    tk.Label(panel, text="かんたん採点", font=("Yu Gothic UI", 9, "bold")).pack()

    mb_frame = tk.Frame(panel)
    mb_frame.pack(pady=5)
    tk.Button(mb_frame, text=f"〇 正解\n({max_score}点)",
              bg="#E3F2FD", fg="#1565C0",
              font=("Yu Gothic UI", 11, "bold"),
              width=8, height=2, relief=tk.RAISED, cursor="hand2",
              activebackground="#BBDEFB").pack(side=tk.LEFT, padx=3)
    tk.Button(mb_frame, text="× 不正解\n(0点)",
              bg="#FFEBEE", fg="#C62828",
              font=("Yu Gothic UI", 11, "bold"),
              width=8, height=2, relief=tk.RAISED, cursor="hand2",
              activebackground="#FFCDD2").pack(side=tk.LEFT, padx=3)

    tk.Label(panel, text="m: 〇正解  b: ×不正解",
             font=("Yu Gothic UI", 8), fg="#7B1FA2").pack(pady=(0, 3))

    # キーボード操作
    tk.Label(panel, text="─" * 20, fg="#ccc").pack(pady=5)
    tk.Label(panel, text="キーボード操作", font=("Yu Gothic UI", 9, "bold")).pack()
    help_text = f"0〜{min(9, max_score)}: 得点入力(自動で次へ)\nm: 〇正解  b: ×不正解\n←→: 前後に移動\nTab: 次の未採点へ\nDel: 得点クリア"
    tk.Label(panel, text=help_text,
             font=("Yu Gothic UI", 8), fg="gray", justify=tk.LEFT).pack(pady=5)

    # 得点チェックボックス
    tk.Label(panel, text="─" * 20, fg="#ccc").pack(pady=5)
    tk.Label(panel, text="入力可能な得点", font=("Yu Gothic UI", 9, "bold")).pack()
    checks_frame = tk.Frame(panel)
    checks_frame.pack()
    for i in range(min(10, max_score + 1)):
        var = tk.BooleanVar(value=True)
        tk.Checkbutton(checks_frame, text=str(i), variable=var,
                       font=("Yu Gothic UI", 8)).grid(
            row=i // 3, column=i % 3, sticky=tk.W, padx=3)

    # フィルタ
    tk.Label(panel, text="─" * 20, fg="#ccc").pack(pady=5)
    filter_var = tk.BooleanVar(value=filter_on)
    tk.Checkbutton(panel, text="未採点のみ表示", variable=filter_var,
                   font=("Yu Gothic UI", 9)).pack(pady=(0, 3))

    tk.Label(panel, text="─" * 20, fg="#ccc").pack(pady=5)
    tk.Label(panel, text=unscored_text,
             font=("Yu Gothic UI", 9), fg="#E65100").pack(pady=(0, 5))

    tk.Button(panel, text="✔ この問題の採点完了",
              bg="#4CAF50", fg="white",
              font=("Yu Gothic UI", 9, "bold"),
              width=18, relief=tk.FLAT, cursor="hand2").pack(pady=3)
    tk.Button(panel, text="キャンセル",
              font=("Yu Gothic UI", 9), width=18).pack(pady=3)

    return win, canvas


def _show_cropped_on_canvas(canvas, img_path: str, region: list):
    """Canvas に切り出し画像を表示"""
    cropped = _crop_region(img_path, region)
    canvas.update_idletasks()
    cw = max(canvas.winfo_width(), 400)
    ch = max(canvas.winfo_height(), 500)
    ratio = min(cw / cropped.width, ch / cropped.height) * 0.85
    new_w = int(cropped.width * ratio)
    new_h = int(cropped.height * ratio)
    cropped_resized = cropped.resize((new_w, new_h), Image.Resampling.LANCZOS)
    tk_img = ImageTk.PhotoImage(cropped_resized)
    canvas.delete("all")
    canvas.create_image(cw // 2, ch // 2, image=tk_img, anchor=tk.CENTER)
    setattr(canvas, "_tk_img_ref", tk_img)


def capture_desc_03_scored_maru(root, image_paths):
    """desc_03 — ○正解 (背景薄青)"""
    fns = sorted(image_paths.keys())
    win, canvas = _build_descriptive_scorer_window(
        root, image_paths, student_idx=0,
        score_text="5", bg_color="#E3F2FD",
        progress="1 / 8", unscored_text="未採点: 7人")
    win.update_idletasks(); win.update()
    _show_cropped_on_canvas(canvas, image_paths[fns[0]], [15, 95, 585, 280])
    canvas.config(bg="#E3F2FD")
    win.update_idletasks(); win.update()
    _capture_window(win, "desc_03_scored_maru")
    win.destroy()


def capture_desc_04_scored_batsu(root, image_paths):
    """desc_04 — ×不正解 (背景薄赤)"""
    fns = sorted(image_paths.keys())
    win, canvas = _build_descriptive_scorer_window(
        root, image_paths, student_idx=3,
        score_text="0", bg_color="#FFEBEE",
        progress="4 / 8", unscored_text="未採点: 5人")
    win.update_idletasks(); win.update()
    _show_cropped_on_canvas(canvas, image_paths[fns[3]], [15, 95, 585, 280])
    canvas.config(bg="#FFEBEE")
    win.update_idletasks(); win.update()
    _capture_window(win, "desc_04_scored_batsu")
    win.destroy()


def capture_desc_05_scored_middle(root, image_paths):
    """desc_05 — △中間点 (背景薄橙)"""
    fns = sorted(image_paths.keys())
    win, canvas = _build_descriptive_scorer_window(
        root, image_paths, student_idx=2,
        score_text="3", bg_color="#FFF3E0",
        progress="3 / 8", unscored_text="未採点: 5人")
    win.update_idletasks(); win.update()
    _show_cropped_on_canvas(canvas, image_paths[fns[2]], [15, 95, 585, 280])
    canvas.config(bg="#FFF3E0")
    win.update_idletasks(); win.update()
    _capture_window(win, "desc_05_scored_middle")
    win.destroy()


def capture_desc_06_filter_active(root, image_paths):
    """desc_06 — 未採点フィルタ ON"""
    fns = sorted(image_paths.keys())
    win, canvas = _build_descriptive_scorer_window(
        root, image_paths, student_idx=4,
        score_text="—", bg_color="white",
        progress="5 / 8", unscored_text="未採点: 4人（フィルタ中）",
        filter_on=True)
    win.update_idletasks(); win.update()
    _show_cropped_on_canvas(canvas, image_paths[fns[4]], [15, 95, 585, 280])
    win.update_idletasks(); win.update()
    _capture_window(win, "desc_06_filter_active")
    win.destroy()


def capture_desc_07_all_scored(root, image_paths):
    """desc_07 — 全員採点済み"""
    fns = sorted(image_paths.keys())
    win, canvas = _build_descriptive_scorer_window(
        root, image_paths, student_idx=7,
        score_text="4", bg_color="white",
        progress="8 / 8", unscored_text="未採点: 0人 ✔ 全員採点済み")
    win.update_idletasks(); win.update()
    _show_cropped_on_canvas(canvas, image_paths[fns[7 % len(fns)]], [15, 95, 585, 280])
    win.update_idletasks(); win.update()
    _capture_window(win, "desc_07_all_scored")
    win.destroy()


# ============================================================
# desc_08: グリッドモード
# ============================================================

def capture_desc_08_grid_mode(root, image_paths):
    """desc_08 — 一覧グリッドモード"""
    win = tk.Toplevel(root)
    win.title("採点: 問1 説明問題 (配点:5点)")
    win.geometry("1000x700+80+80")
    win.configure(bg="#F5F7FA")
    BG = "#F5F7FA"

    # ズームバー
    zoom_bar = tk.Frame(win, bg="#37474F", padx=10, pady=4)
    zoom_bar.pack(fill=tk.X)
    zoom_frame = tk.Frame(zoom_bar, bg="#37474F")
    zoom_frame.pack(side=tk.RIGHT, fill=tk.X, expand=True)
    tk.Label(zoom_frame, text="🔍 サイズ:",
             font=("Yu Gothic UI", 9), bg="#37474F", fg="white").pack(side=tk.LEFT)
    zoom_slider = tk.Scale(zoom_frame, from_=80, to=800, orient=tk.HORIZONTAL,
                           length=200, bg="#37474F", fg="#FFD54F",
                           troughcolor="#546E7A", highlightthickness=0,
                           sliderlength=20)
    zoom_slider.set(160)
    zoom_slider.pack(side=tk.LEFT, padx=5)
    tk.Label(zoom_frame, text="160px",
             font=("Yu Gothic UI", 9, "bold"), bg="#37474F", fg="#FFD54F",
             width=8, anchor=tk.CENTER).pack(side=tk.LEFT, padx=2)

    # 上部バー
    top_bar = tk.Frame(win, bg=BG, padx=8, pady=5)
    top_bar.pack(fill=tk.X)
    tk.Label(top_bar, text="並び順:", font=("Yu Gothic UI", 9), bg=BG).pack(side=tk.LEFT)
    sort_combo = ttk.Combobox(top_bar, values=["ファイル名順", "得点 昇順", "得点 降順"],
                              state="readonly", width=18)
    sort_combo.set("ファイル名順")
    sort_combo.pack(side=tk.LEFT, padx=(3, 10))
    tk.Label(top_bar, text="採点済み: 4/8  (未採点: 4)",
             font=("Yu Gothic UI", 9, "bold"), bg=BG, fg="#555").pack(side=tk.LEFT)
    tk.Label(top_bar, text="", bg=BG).pack(side=tk.LEFT, expand=True)
    tk.Button(top_bar, text="✔ この問題の採点完了", bg="#4CAF50", fg="white",
              font=("Yu Gothic UI", 9, "bold"), relief=tk.FLAT).pack(side=tk.RIGHT)

    # グリッドエリア
    grid_area = tk.Frame(win, bg=BG, padx=8)
    grid_area.pack(fill=tk.BOTH, expand=True)

    fns = sorted(image_paths.keys())
    card_data = [
        (fns[0], 5, "#E3F2FD", "○", "#1565C0"),
        (fns[1], 3, "#FFF3E0", "△", "#E65100"),
        (fns[2], 0, "#FFEBEE", "×", "#C62828"),
        (fns[3], None, "#F5F5F5", "─", "#999"),
        (fns[4], 5, "#E3F2FD", "○", "#1565C0"),
        (fns[5 % len(fns)], None, "#F5F5F5", "─", "#999"),
        (fns[6 % len(fns)], None, "#F5F5F5", "─", "#999"),
        (fns[7 % len(fns)], None, "#F5F5F5", "─", "#999"),
    ]

    for i, (fn, score, bg_col, mark, mark_fg) in enumerate(card_data):
        r, c = divmod(i, 4)
        card = tk.Frame(grid_area, bg=bg_col, bd=1, relief=tk.RAISED,
                        padx=3, pady=3, width=210, height=170)
        card.grid(row=r, column=c, padx=4, pady=4, sticky="nsew")
        card.grid_propagate(False)

        # サムネイル画像を切り出して表示
        try:
            cropped = _crop_region(image_paths[fn], [15, 95, 585, 280])
            thumb = cropped.resize((180, 55), Image.Resampling.LANCZOS)
            tk_thumb = ImageTk.PhotoImage(thumb)
            lbl = tk.Label(card, image=tk_thumb, bg=bg_col)
            setattr(lbl, "image", tk_thumb)
            lbl.pack(pady=(5, 2))
        except Exception:
            tk.Label(card, text="📄", font=("", 20), bg=bg_col).pack(pady=(5, 2))

        sf = tk.Frame(card, bg=bg_col)
        sf.pack(fill=tk.X)
        tk.Label(sf, text=mark, font=("Yu Gothic UI", 12, "bold"),
                 fg=mark_fg, bg=bg_col).pack(side=tk.LEFT)
        score_text = f"{score}点" if score is not None else "未採点"
        tk.Label(sf, text=score_text, font=("Yu Gothic UI", 8),
                 fg="#555", bg=bg_col).pack(side=tk.LEFT, padx=3)

        tk.Label(card, text=fn, font=("Yu Gothic UI", 7),
                 fg="#777", bg=bg_col).pack()

    for c in range(4):
        grid_area.columnconfigure(c, weight=1)

    # 下部得点ボタンバー
    score_bar = tk.Frame(win, bg="#ECEFF1", padx=8, pady=8)
    score_bar.pack(fill=tk.X, side=tk.BOTTOM)
    tk.Label(score_bar, text="得点ボタン（クリック後、サムネイルをクリックで得点付与）:",
             font=("Yu Gothic UI", 9), bg="#ECEFF1", fg="#555").pack(side=tk.LEFT, padx=(0, 8))
    for i in range(6):
        bg_c = "#FF8A65" if i == 3 else "#E3F2FD"
        rel = tk.SUNKEN if i == 3 else tk.RAISED
        tk.Button(score_bar, text=str(i), width=3, font=("Yu Gothic UI", 10, "bold"),
                  bg=bg_c, relief=rel, cursor="hand2").pack(side=tk.LEFT, padx=2)
    tk.Button(score_bar, text="📷 答案を表示",
              font=("Yu Gothic UI", 9, "bold"),
              bg="#E3F2FD", relief=tk.RAISED,
              cursor="hand2").pack(side=tk.RIGHT, padx=(6, 0))
    tk.Button(score_bar, text="選択解除", font=("Yu Gothic UI", 8),
              bg="#E0E0E0", relief=tk.FLAT,
              cursor="hand2").pack(side=tk.RIGHT, padx=(4, 0))
    tk.Label(score_bar, text="アクティブ: 3点",
             font=("Yu Gothic UI", 9, "bold"), bg="#ECEFF1", fg="#E65100"
             ).pack(side=tk.RIGHT, padx=(0, 10))

    win.update_idletasks(); win.update()
    _capture_window(win, "desc_08_grid_mode")
    win.destroy()


# ============================================================
# メイン実行
# ============================================================

def main():
    print("=" * 60)
    print("  docs/images/ スクリーンショット再生成")
    print("=" * 60)

    root = tk.Tk()
    root.withdraw()

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)

        print("\n[1/3] ダミー答案画像を生成中...")
        image_paths = _create_dummy_sheets_in_folder(td, count=8)
        print(f"  → {len(image_paths)} 枚のダミー答案を生成")

        # 最初の答案のパスを取得
        first_fn = sorted(image_paths.keys())[0]
        first_path = image_paths[first_fn]

        print("\n[2/3] スクリーンショットをキャプチャ中...")

        print("  → 19_integrated_descriptive_setup")
        capture_19_integrated_setup(root, first_path)

        print("  → 06_single_question_scorer")
        capture_06_single_question_scorer(root, image_paths)

        print("  → desc_03_scored_maru")
        capture_desc_03_scored_maru(root, image_paths)

        print("  → desc_04_scored_batsu")
        capture_desc_04_scored_batsu(root, image_paths)

        print("  → desc_05_scored_middle")
        capture_desc_05_scored_middle(root, image_paths)

        print("  → desc_06_filter_active")
        capture_desc_06_filter_active(root, image_paths)

        print("  → desc_07_all_scored")
        capture_desc_07_all_scored(root, image_paths)

        print("  → desc_08_grid_mode")
        capture_desc_08_grid_mode(root, image_paths)

        print("  → 02a_main_mark_only (v4.5 認識方式コンボ)")
        capture_02a_main_mark_only(root)

        print("  → 18_mark_checker (v4.5 タブ+グリッド)")
        capture_18_mark_checker(root)

    print("\n[3/3] 清掃中...")
    root.destroy()

    # 結果確認
    print("\n" + "=" * 60)
    print("  再生成した画像:")
    print("=" * 60)
    targets = [
        "19_integrated_descriptive_setup.png",
        "06_single_question_scorer.png",
        "desc_03_scored_maru.png",
        "desc_04_scored_batsu.png",
        "desc_05_scored_middle.png",
        "desc_06_filter_active.png",
        "desc_07_all_scored.png",
        "desc_08_grid_mode.png",
        "02a_main_mark_only.png",
        "18_mark_checker.png",
    ]
    for t in targets:
        p = DOCS_IMAGES / t
        if p.exists():
            size_kb = p.stat().st_size / 1024
            print(f"  ✔ {t} ({size_kb:.1f} KB)")
        else:
            print(f"  ✗ {t} — 失敗")

    print("\n完了！")


    print("\n完了！")


# ============================================================
# 02a. メイン画面（マーク採点モード） — v4.5 認識方式コンボボックス付き
# ============================================================

def capture_02a_main_mark_only(root: tk.Tk):
    """02a_main_mark_only.png — マーク採点メイン画面（v4.5 認識方式選択付き）"""
    BG = "#F0F4F8"
    SECTION_BG = "#FFFFFF"
    BTN_RUN = "#4CAF50"
    BTN_FILE = "#E3F2FD"
    HEADER_TEXT = "#37474F"
    FONT_NORM = ("Yu Gothic UI", 9)
    FONT_BOLD = ("Yu Gothic UI", 9, "bold")
    FONT_HEAD = ("Yu Gothic UI", 11, "bold")

    win = tk.Toplevel(root)
    win.title("採点侍 — SaitenSamurai v4.5")
    win.geometry("780x620")
    win.configure(bg=BG)
    win.resizable(False, False)

    # ── タイトルバー ──
    title_bar = tk.Frame(win, bg="#1A237E", padx=12, pady=8)
    title_bar.pack(fill=tk.X)
    tk.Label(title_bar, text="⚔ 採点侍  SaitenSamurai v4.5",
             font=("Yu Gothic UI", 14, "bold"), bg="#1A237E", fg="white").pack(side=tk.LEFT)
    tk.Label(title_bar, text="マーク採点モード",
             font=("Yu Gothic UI", 9), bg="#1A237E", fg="#90CAF9").pack(side=tk.RIGHT)

    main_frame = tk.Frame(win, bg=BG, padx=16, pady=12)
    main_frame.pack(fill=tk.BOTH, expand=True)

    # ── ファイル選択セクション ──
    file_group = tk.LabelFrame(main_frame, text="ファイル設定", font=FONT_BOLD,
                               bg=SECTION_BG, fg=HEADER_TEXT, relief=tk.FLAT,
                               padx=10, pady=8)
    file_group.pack(fill=tk.X, pady=(0, 8))

    def file_row(parent, label, placeholder, btn_text="参照"):
        row = tk.Frame(parent, bg=SECTION_BG)
        row.pack(fill=tk.X, pady=3)
        tk.Label(row, text=label, font=FONT_NORM, bg=SECTION_BG,
                 width=16, anchor=tk.W).pack(side=tk.LEFT)
        entry = tk.Entry(row, font=("Yu Gothic UI", 8), fg="#999",
                         relief=tk.SOLID, bd=1)
        entry.insert(0, placeholder)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        tk.Button(row, text=btn_text, font=FONT_NORM, bg=BTN_FILE,
                  relief=tk.FLAT, cursor="hand2", padx=6).pack(side=tk.LEFT)
        return entry

    file_row(main_frame, "スキャン画像フォルダ", "例: C:\\Users\\teacher\\scans\\class_A",
             "フォルダ選択")
    file_row(main_frame, "Mark2 座標ファイル", "例: M2-03-002_座標ファイル.xlsx",
             "ファイル選択")
    file_row(main_frame, "正答データ", "（認識実行後に自動設定）",
             "ファイル選択")

    # Skip 数
    skip_row = tk.Frame(main_frame, bg=BG)
    skip_row.pack(fill=tk.X, pady=(0, 6))
    tk.Label(skip_row, text="Skip 数:", font=FONT_NORM, bg=BG,
             width=16, anchor=tk.W).pack(side=tk.LEFT)
    tk.Spinbox(skip_row, from_=0, to=20, width=5, font=FONT_NORM).pack(side=tk.LEFT, padx=4)
    tk.Label(skip_row, text="（学年・クラス・番号・氏名 など解答欄の前にあるマーク欄の数）",
             font=("Yu Gothic UI", 8), fg="#777", bg=BG).pack(side=tk.LEFT)

    # ── オプション / 認識方式セクション ──
    option_group = tk.LabelFrame(main_frame, text="OMR オプション", font=FONT_BOLD,
                                 bg=SECTION_BG, fg=HEADER_TEXT, relief=tk.FLAT,
                                 padx=10, pady=8)
    option_group.pack(fill=tk.X, pady=(0, 8))

    mode_row = tk.Frame(option_group, bg=SECTION_BG)
    mode_row.pack(fill=tk.X, pady=2)
    tk.Label(mode_row, text="認識方式:", font=FONT_NORM, bg=SECTION_BG).pack(side=tk.LEFT)
    combo = ttk.Combobox(mode_row, width=26,
                         values=["（推奨）クラスタリング", "しきい値による識別（従来式）"],
                         state="readonly")
    combo.set("（推奨）クラスタリング")
    combo.pack(side=tk.LEFT, padx=(5, 12))
    tk.Label(mode_row, text="✔ K-means 7次元特徴量で高精度判定",
             font=("Yu Gothic UI", 8), fg="#388E3C", bg=SECTION_BG).pack(side=tk.LEFT)

    # ── 実行ボタン ──
    pipeline = tk.Frame(main_frame, bg=BG)
    pipeline.pack(fill=tk.X, pady=(4, 0))
    pipeline.columnconfigure(0, weight=1)
    pipeline.columnconfigure(1, weight=1)
    pipeline.columnconfigure(2, weight=1)

    def step_btn(col, step, title, desc, color, enabled=True):
        f = tk.LabelFrame(pipeline, text=step, font=FONT_BOLD,
                          bg=SECTION_BG, fg=HEADER_TEXT, relief=tk.FLAT,
                          padx=8, pady=8)
        f.grid(row=0, column=col, sticky="nsew", padx=4)
        tk.Label(f, text=title, font=FONT_HEAD, bg=SECTION_BG, fg=HEADER_TEXT).pack(pady=(0, 4))
        tk.Button(f, text=desc, font=("Yu Gothic UI", 10, "bold"),
                  bg=color if enabled else "#BDBDBD",
                  fg="white", relief=tk.FLAT, cursor="hand2",
                  padx=10, pady=6,
                  state=tk.NORMAL if enabled else tk.DISABLED).pack(fill=tk.X)

    step_btn(0, "Step 1 — OMR 認識", "マーク読み取り", "▶ 認識実行", "#4CAF50")
    step_btn(1, "Step 2 — 正答・採点", "採点実行", "✔ 採点実行", "#1976D2")
    step_btn(2, "Step 3 — 結果出力", "Excel 出力", "📊 出力実行", "#FF8F00")

    # ステータスバー
    status = tk.Frame(win, bg="#37474F", padx=10, pady=4)
    status.pack(fill=tk.X, side=tk.BOTTOM)
    tk.Label(status, text="就緒 — ファイルを選択して認識を実行してください",
             font=("Yu Gothic UI", 8), bg="#37474F", fg="#B0BEC5").pack(side=tk.LEFT)

    win.update_idletasks()
    win.update()
    _capture_window(win, "02a_main_mark_only")
    win.destroy()


# ============================================================
# 18. マークチェック画面 — v4.5 タブ（薄い/濃い解答）付きグリッドビュー
# ============================================================

def capture_18_mark_checker(root: tk.Tk):
    """18_mark_checker.png — マークチェック画面（v4.5 グリッド＋タブ）"""
    BG = "#FAFAFA"
    TOOLBAR_BG = "#37474F"

    win = tk.Toplevel(root)
    win.title("マークチェック")
    win.geometry("940x680")
    win.configure(bg=BG)

    # ── 進捗バー ──
    prog_bar = tk.Frame(win, bg="#1A237E", padx=10, pady=6)
    prog_bar.pack(fill=tk.X)
    tk.Label(prog_bar, text="マークチェック",
             font=("Yu Gothic UI", 12, "bold"), bg="#1A237E", fg="white").pack(side=tk.LEFT)
    tk.Label(prog_bar, text="エラー: 12件  |  チェック済み: 8件  |  進捗: 67%",
             font=("Yu Gothic UI", 9), bg="#1A237E", fg="#90CAF9").pack(side=tk.LEFT, padx=16)
    tk.Canvas(prog_bar, width=200, height=12, bg="#455A64",
              highlightthickness=0).pack(side=tk.LEFT)

    # ── ツールバー（タブ付き）──
    toolbar = tk.Frame(win, bg=TOOLBAR_BG, padx=8, pady=5)
    toolbar.pack(fill=tk.X)

    # ページング
    tk.Button(toolbar, text="◀", width=3, bg="#546E7A", fg="white", relief=tk.FLAT,
              cursor="hand2").pack(side=tk.LEFT)
    tk.Label(toolbar, text="1/2", font=("Yu Gothic UI", 9), bg=TOOLBAR_BG,
             fg="white").pack(side=tk.LEFT, padx=4)
    tk.Button(toolbar, text="▶", width=3, bg="#546E7A", fg="white", relief=tk.FLAT,
              cursor="hand2").pack(side=tk.LEFT)

    # 件数ラベル
    tk.Label(toolbar, text="12件", font=("Yu Gothic UI", 9, "bold"),
             bg=TOOLBAR_BG, fg="#FFD54F").pack(side=tk.LEFT, padx=10)

    # タブ（v4.5 新機能）
    tk.Button(toolbar, text="薄い解答(ノーマーク疑惑)",
              bg="#42A5F5", fg="white",
              font=("Yu Gothic UI", 9, "bold"),
              relief=tk.FLAT, cursor="hand2", padx=8).pack(side=tk.LEFT, padx=(8, 2))
    tk.Button(toolbar, text="濃い解答(複数マーク疑惑)",
              bg="#546E7A", fg="white",
              font=("Yu Gothic UI", 9),
              relief=tk.FLAT, cursor="hand2", padx=8).pack(side=tk.LEFT, padx=(2, 8))

    # サイズスライダー
    tk.Label(toolbar, text="サイズ:", font=("Yu Gothic UI", 9),
             bg=TOOLBAR_BG, fg="white").pack(side=tk.RIGHT)
    tk.Scale(toolbar, from_=80, to=300, orient=tk.HORIZONTAL, length=120,
             showvalue=False, bg=TOOLBAR_BG, fg="white",
             highlightthickness=0, troughcolor="#546E7A").pack(side=tk.RIGHT, padx=(0, 4))

    # ── グリッドエリア ──
    grid_outer = tk.Frame(win, bg=BG)
    grid_outer.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

    # ダミーカードデータ
    cards = [
        ("student_003.jpg", "Q5", "0",  "#FFCDD2", "未マーク"),
        ("student_007.jpg", "Q3", "3,5", "#FFF3E0", "複数マーク"),
        ("student_012.jpg", "Q8", "0",  "#FFCDD2", "未マーク"),
        ("student_015.jpg", "Q2", "0",  "#FFCDD2", "未マーク"),
        ("student_018.jpg", "Q6", "2,4", "#FFF3E0", "複数マーク"),
        ("student_021.jpg", "Q1", "0",  "#FFCDD2", "未マーク"),
        ("student_025.jpg", "Q9", "1,6", "#FFF3E0", "複数マーク"),
        ("student_029.jpg", "Q4", "0",  "#FFCDD2", "未マーク"),
    ]

    for i, (fn, q, val, bg_col, err_type) in enumerate(cards):
        r, c = divmod(i, 4)
        card = tk.Frame(grid_outer, bg=bg_col, bd=1, relief=tk.RAISED,
                        padx=4, pady=4, width=205, height=160)
        card.grid(row=r, column=c, padx=4, pady=4, sticky="nsew")
        card.grid_propagate(False)

        # ダミー画像エリア（灰色の矩形）
        img_canvas = tk.Canvas(card, width=180, height=80, bg="#EEEEEE",
                               highlightthickness=0)
        img_canvas.pack(pady=(2, 4))
        img_canvas.create_rectangle(20, 20, 160, 60, fill="#CFD8DC", outline="#90A4AE", width=2)
        img_canvas.create_text(90, 40, text=f"マーク欄 {q}",
                               font=("Yu Gothic UI", 9), fill="#546E7A")

        info_f = tk.Frame(card, bg=bg_col)
        info_f.pack(fill=tk.X)
        tk.Label(info_f, text=fn, font=("Yu Gothic UI", 7), fg="#777", bg=bg_col).pack(side=tk.LEFT)
        tk.Label(info_f, text=err_type, font=("Yu Gothic UI", 8, "bold"),
                 fg="#C62828" if "未" in err_type else "#E65100", bg=bg_col).pack(side=tk.RIGHT)

        tk.Label(card, text=f"読取値: {val}", font=("Yu Gothic UI", 8),
                 fg="#555", bg=bg_col).pack()

    for c in range(4):
        grid_outer.columnconfigure(c, weight=1)

    # ── 修正パネル（下部）──
    ctrl = tk.Frame(win, bg="#ECEFF1", padx=8, pady=6)
    ctrl.pack(fill=tk.X, side=tk.BOTTOM)
    tk.Label(ctrl, text="修正値を選択:", font=("Yu Gothic UI", 9, "bold"),
             bg="#ECEFF1").pack(side=tk.LEFT)
    for i in range(1, 8):
        tk.Button(ctrl, text=str(i), width=3, font=("Yu Gothic UI", 10, "bold"),
                  bg="#E3F2FD", relief=tk.RAISED, cursor="hand2").pack(side=tk.LEFT, padx=2)
    tk.Button(ctrl, text="-1 無効", width=6, font=("Yu Gothic UI", 9, "bold"),
              bg="#FFCDD2", fg="#C62828", relief=tk.RAISED, cursor="hand2").pack(side=tk.LEFT, padx=(8, 0))
    tk.Button(ctrl, text="✔ 確定して次へ",
              bg="#4CAF50", fg="white",
              font=("Yu Gothic UI", 9, "bold"),
              relief=tk.FLAT, cursor="hand2", padx=10).pack(side=tk.RIGHT)

    win.update_idletasks()
    win.update()
    _capture_window(win, "18_mark_checker")
    win.destroy()


if __name__ == "__main__":
    main()
