
"""
capture_improvements.py — 改善項目のビジュアル検証用キャプチャスクリプト

以下の2つの機能改善を証明するスクリーンショットを撮影し、HTMLレポートを生成します。
1. 記述式採点: "clean"（枠なし）画像が使用されていること
2. マークチェック: 正答（Answer Key）の赤点線枠が表示されていること

Usage:
    python tests/capture_improvements.py
"""

import sys
import os
import time
import shutil
import ctypes
import ctypes.wintypes as wintypes
import tkinter as tk
import cv2
import numpy as np
import pandas as pd
from pathlib import Path
from PIL import Image, ImageGrab, ImageDraw

# Add main_src to path
PROJECT_ROOT = Path(__file__).parent.parent
MAIN_SRC = PROJECT_ROOT / "main_src"
sys.path.insert(0, str(MAIN_SRC))

# Pre-import to avoid circular dependency (descriptive_gui <-> descriptive_scorer)
import descriptive_scorer
from gui_components import MarkCheckerGUI
from descriptive_gui import _SingleQuestionScorer

# ============================================================
# キャプチャ用ユーティリティ
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

def _get_hwnd(widget):
    try:
        frame_id = widget.wm_frame()
        if frame_id and frame_id != "0x0":
            return int(frame_id, 16)
    except:
        pass
    try:
        return widget.winfo_id()
    except:
        return None

def _capture_window_win32(widget, path):
    try:
        hwnd = _get_hwnd(widget)
        if not hwnd: return False
        
        rect = wintypes.RECT()
        ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
        w = rect.right - rect.left
        h = rect.bottom - rect.top
        if w < 10 or h < 10: return False
        
        wDC = ctypes.windll.user32.GetWindowDC(hwnd)
        if not wDC: return False
        
        try:
            dcObj = ctypes.windll.gdi32.CreateCompatibleDC(wDC)
            bmp = ctypes.windll.gdi32.CreateCompatibleBitmap(wDC, w, h)
            old = ctypes.windll.gdi32.SelectObject(dcObj, bmp)
            ok = ctypes.windll.user32.PrintWindow(hwnd, dcObj, 2)
            
            if ok:
                bmi = BITMAPINFOHEADER()
                bmi.biSize = ctypes.sizeof(BITMAPINFOHEADER)
                bmi.biWidth = w
                bmi.biHeight = -h
                bmi.biPlanes = 1
                bmi.biBitCount = 32
                bmi.biCompression = 0
                
                buf = ctypes.create_string_buffer(w * h * 4)
                ctypes.windll.gdi32.GetDIBits(dcObj, bmp, 0, h, buf, ctypes.byref(bmi), 0)
                img = Image.frombuffer("RGBA", (w, h), buf, "raw", "BGRA", 0, 1)
                img.convert("RGB").save(str(path))
            
            ctypes.windll.gdi32.SelectObject(dcObj, old)
            ctypes.windll.gdi32.DeleteObject(bmp)
            ctypes.windll.gdi32.DeleteDC(dcObj)
        finally:
            ctypes.windll.user32.ReleaseDC(hwnd, wDC)
        return bool(ok)
    except Exception as e:
        print(f"Win32 capture failed: {e}")
        return False

def capture_widget(widget, filename, delay=1.0):
    widget.update_idletasks()
    widget.update()
    time.sleep(delay)
    
    path = Path("tests/visual_report_images") / filename
    path.parent.mkdir(exist_ok=True, parents=True)
    
    if sys.platform == "win32":
        if _capture_window_win32(widget, path):
            print(f"Captured: {path}")
            return path
            
    # Fallback
    try:
        x = widget.winfo_rootx()
        y = widget.winfo_rooty()
        w = widget.winfo_width()
        h = widget.winfo_height()
        ImageGrab.grab((x, y, x+w, y+h)).save(str(path))
        print(f"Captured (fallback): {path}")
        return path
    except:
        print("Capture failed")
        return None

# ============================================================
# テストデータ準備
# ============================================================

def setup_test_data(base_dir):
    base = Path(base_dir)
    if base.exists():
        shutil.rmtree(base)
    base.mkdir(parents=True)
    
    # 1. Images
    boxed_dir = base / "boxed"
    clean_dir = base / "clean"
    boxed_dir.mkdir()
    clean_dir.mkdir()
    
    # Create distinguishable images
    def create_img(text, color, folder):
        img = np.full((842, 595, 3), 255, dtype=np.uint8)
        cv2.putText(img, text, (50, 400), cv2.FONT_HERSHEY_SIMPLEX, 2, color, 5)
        # Add some "marks"
        cv2.rectangle(img, (100, 200), (150, 250), (0, 0, 0), 2)
        cv2.imwrite(str(folder / "test_01.jpg"), img)
        return img

    create_img("BOXED IMAGE (Wrong)", (0, 0, 255), boxed_dir)
    create_img("CLEAN IMAGE (Correct)", (0, 255, 0), clean_dir)
    
    # 2. Descriptive Config
    config = {
        "id": "q1", "name": "Q1", "page": 1, 
        "region": [50, 300, 500, 500],
        "max_score": 10, "aspect": 1
    }
    
    # 3. Mark Checker Data
    # Coordinates CSV (expanded format)
    with open(base / "coordinates.csv", "w", encoding='utf-8') as f:
        # We need "choice" column for updated MarkChecker
        f.write("image_path,question_no,choice,x,y,width,height,mark_coords,choices_bbox\n")
        # Choice 1 at 100,200
        # mark_coords is raw string. choices_bbox is total area.
        f.write(f"test_01.jpg,1,1,100,200,50,50,100;200;50;50,100;200;50;50\n")

    # Mark2 Results
    df_res = pd.DataFrame([{
        "No": 1, "File": "test_01.jpg", "1": "1" # User marked choice 1
    }])
    df_res.to_excel(base / "mark_result.xlsx", index=False)
    
    # Answer Key (Template)
    # Q1 correct answer is 1
    df_key = pd.DataFrame([{
        "問題番号": 1, "正答": "1", "配点": 5, "観点": 1
    }])
    df_key.to_excel(base / "answer_key.xlsx", index=False)
    
    return base, config

# ============================================================
# メイン実行
# ============================================================

def main():
    root = tk.Tk()
    root.withdraw() 
    
    test_dir = Path("tests/temp_capture_env")
    base, q_config = setup_test_data(test_dir)
    
    clean_folder = test_dir / "clean"
    boxed_folder = test_dir / "boxed"
    
    # --- Capture 1: Descriptive Preview (Clean Image) ---
    print("--- Capturing Descriptive Preview ---")
    
    image_paths = {"test_01.jpg": str(clean_folder / "test_01.jpg")}
    
    # Instantiate _SingleQuestionScorer
    # Note: image_folder argument determines where to look for images if not found in image_paths?
    # Actually, in main_gui, image_folder is passed as the "clean_folder".
    
    scorer = _SingleQuestionScorer(
        root, 
        question_config=q_config, 
        image_paths=image_paths, 
        existing_scores={},
        initial_mode="1枚ずつ",
        image_folder=str(clean_folder) # This is the key for the improvement
    )
    
    captured_path1 = None
    
    def on_scorer_open():
        nonlocal captured_path1
        try:
            # win is created in run()
            win = scorer._win  
            win.deiconify()
            win.lift()
            win.update()
            print(f"Scorer Win Size: {win.winfo_width()}x{win.winfo_height()}")
            time.sleep(1.0)
            captured_path1 = capture_widget(win, "01_descriptive_clean.png", delay=0.5)
            win.destroy() # Close to finish run()
            root.quit() # Exit mainloop for this part
        except Exception as e:
            print(f"Error in on_scorer_open: {e}")
            if hasattr(scorer, '_win') and scorer._win:
                scorer._win.destroy()
            root.quit()

    root.after(1000, on_scorer_open)
    
    # run() starts the window and likely blocks with wait_window
    try:
        scorer.run()
    except tk.TclError:
        pass # Window destroyed
        
    print(f"Descriptive capture done: {captured_path1}")

    # --- Capture 2: Mark Checker (Red Dashed Box) ---
    print("--- Capturing Mark Checker ---")
    
    # MarkCheckerGUI is non-blocking in init, but we usually run it as a tool.
    # We can just create it on a Toplevel.
    
    mark_win = tk.Toplevel(root)
    mark_win.geometry("1000x800+50+50")
    
    checker_gui = MarkCheckerGUI(
        mark_win, 
        image_folder=str(boxed_folder), # Checker usually views boxed or original
        coords_csv_path=str(test_dir / "coordinates.csv"),
        xlsx_path=str(test_dir / "mark_result.xlsx"),
        skip_questions=0,
        template_path=str(test_dir / "answer_key.xlsx")
    )
    
    checker_gui.load_data()
    checker_gui.current_index = 0
    checker_gui.show_current()
    
    checker_gui.window.deiconify()
    checker_gui.window.lift()
    checker_gui.window.update()
    print(f"Checker Win Size: {checker_gui.window.winfo_width()}x{checker_gui.window.winfo_height()}")
    time.sleep(1.0)
    captured_path2 = capture_widget(checker_gui.window, "02_mark_checker_dashed.png", delay=0.5)
    mark_win.destroy()


    # --- Generate HTML ---
    html = f"""
    <!DOCTYPE html>
    <html lang="ja">
    <head>
        <meta charset="UTF-8">
        <title>SaitenSamurai v4.3.2 Improvements Visual Report</title>
        <style>
            body {{ font-family: "Helvetica Neue", Arial, sans-serif; padding: 20px; background: #f5f5f5; color: #333; }}
            .container {{ max-width: 1000px; margin: 0 auto; background: white; padding: 30px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); border-radius: 8px; }}
            h1 {{ border-bottom: 2px solid #333; padding-bottom: 15px; margin-bottom: 30px; }}
            .section {{ margin-bottom: 50px; border-bottom: 1px solid #eee; padding-bottom: 30px; }}
            .section:last-child {{ border-bottom: none; }}
            h2 {{ display: flex; align-items: center; margin-bottom: 15px; }}
            .screenshot {{ 
                border: 1px solid #ddd; 
                max-width: 100%; 
                margin-top: 15px; 
                box-shadow: 0 4px 12px rgba(0,0,0,0.15); 
                border-radius: 4px;
            }}
            .meta {{ color: #666; font-size: 0.9em; margin-bottom: 20px; }}
            .badge {{ 
                display: inline-block; padding: 4px 10px; border-radius: 20px; 
                color: white; font-weight: bold; font-size: 0.8em; margin-left: 10px; 
            }}
            .badge-green {{ background: #4CAF50; }}
            .badge-blue {{ background: #2196F3; }}
            .desc {{ line-height: 1.6; margin-bottom: 15px; }}
            .check-list {{ background: #e8f5e9; padding: 15px; border-radius: 6px; }}
            .check-list li {{ margin-bottom: 5px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>SaitenSamurai v4.3.2 改善項目ビジュアルレポート</h1>
            <p class="meta">生成日時: {time.strftime('%Y-%m-%d %H:%M:%S')}</p>
            
            <div class="section">
                <h2>1. 記述式採点プレビューの改善 <span class="badge badge-green">Improvement</span></h2>
                <p class="desc">
                    記述式採点のプレビュー画面で、マーク認識枠が描画されていない「クリーンな画像」が使用されていることを確認します。
                    従来は枠（BOXED）が表示されていましたが、文字の視認性を向上させるため、補正後の元画像を表示するように変更しました。
                </p>
                <div class="check-list">
                    <strong>検証ポイント:</strong>
                    <ul>
                        <li>画像内に "CLEAN IMAGE (Correct)" という文字（テスト用ダミー）が表示されていること。</li>
                        <li>"BOXED IMAGE (Wrong)" ではないこと。</li>
                    </ul>
                </div>
                <img src="visual_report_images/{captured_path1.name if captured_path1 else 'error.png'}" class="screenshot">
            </div>
            
            <div class="section">
                <h2>2. マークチェック画面の正答表示 <span class="badge badge-blue">Improvement</span></h2>
                <p class="desc">
                    マークチェック画面において、正答（Answer Key）の箇所に「赤色の点線枠」を表示する機能です。
                    ユーザーは採点結果の修正を行う際に、正解の場所を直感的に把握できます。
                </p>
                <div class="check-list">
                    <strong>検証ポイント:</strong>
                    <ul>
                        <li>正答（Choice 1）の箇所（画面上の矩形）に赤色の点線枠が表示されていること。</li>
                        <li>ダミー画像上の枠線（黒実線）とは別に、GUIによって描画された赤枠であること。</li>
                    </ul>
                </div>
                <img src="visual_report_images/{captured_path2.name if captured_path2 else 'error.png'}" class="screenshot">
            </div>
            
        </div>
    </body>
    </html>
    """
    
    report_path = PROJECT_ROOT / "tests/visual_report.html"
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(html)
        
    print(f"Report generated: {report_path}")

if __name__ == "__main__":
    main()
