"""preview.py — 開発用: Xvfb 上で webui を起動してスクリーンショットを撮る。

使い方: xvfb-run -a -s "-screen 0 1400x900x24" python webui/tools/preview.py out.png
tk 時代のキャプチャ評価サイクル（撮る→目視→直す）を webui でも回すための道具。
"""
import sys
import threading
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import webview  # noqa: E402
from app import create_app  # noqa: E402


def _setup_demo(bridge):
    """合成答案でデータソース選択＋認識実行まで自動で流す（目視評価用）。

    ダイアログ応答を偽装して bridge の正規経路をそのまま通す。
    """
    import tempfile
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent / "tests"))
    import cv2
    from test_multidigit_image_e2e import make_coord_xlsx, make_sheet, sym_positions

    tmp = Path(tempfile.mkdtemp(prefix="saiten_demo_"))
    coord = make_coord_xlsx(tmp / "coord.xlsx", 3)
    scans = tmp / "scans"
    scans.mkdir()
    s1 = {q: p for q, p in zip([1, 2, 3], sym_positions('-24'))}
    s2 = {1: sym_positions('9')[0], 3: sym_positions('4')[0]}
    for name, filled in [("s1.png", s1), ("s2.png", s2)]:
        cv2.imwrite(str(scans / name), make_sheet(filled, with_markers=True))

    real_folder = bridge._win.open_folder_dialog
    real_file = bridge._win.open_file_dialog
    bridge._win.open_folder_dialog = lambda **kw: str(scans)
    bridge._win.open_file_dialog = lambda **kw: str(coord)
    bridge.set_mode("mark_only", "multi_digit")
    bridge.set_skip_questions(0)
    time.sleep(2.5)  # 画面初期化を待ってから UI へ push させる
    bridge._win.eval_js("document.querySelector('[data-action=select_image_folder]').click()")
    time.sleep(0.6)
    bridge._win.eval_js("document.querySelector('[data-action=select_coord_file]').click()")
    time.sleep(0.6)
    bridge._win.eval_js("document.getElementById('btn-run-recognition').click()")
    time.sleep(3)
    bridge._win.open_folder_dialog = real_folder
    bridge._win.open_file_dialog = real_file


def main():
    out = Path(sys.argv[1] if len(sys.argv) > 1 else "webui_preview.png")
    demo = "--demo" in sys.argv
    window, bridge = create_app()

    def shoot():
        if demo:
            _setup_demo(bridge)
            time.sleep(2)
        else:
            time.sleep(4)  # 描画待ち
        try:
            from PIL import ImageGrab
            img = ImageGrab.grab(xdisplay="")
            img.save(out)
            print(f"saved: {out} ({img.size[0]}x{img.size[1]})")
        finally:
            window.destroy()

    threading.Thread(target=shoot, daemon=True).start()
    webview.start()


if __name__ == "__main__":
    main()
