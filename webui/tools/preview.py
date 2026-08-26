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


def main():
    out = Path(sys.argv[1] if len(sys.argv) > 1 else "webui_preview.png")
    window, _bridge = create_app()

    def shoot():
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
