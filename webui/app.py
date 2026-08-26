"""app.py — webui のエントリポイント。

このファイルだけが pywebview を import する（bridge は import しない —
webui/docs/plan.md の設計原則）。起動: python webui/app.py
"""
import argparse
import logging
import sys
from pathlib import Path

import webview

# PyInstaller (onefile) では同梱データは sys._MEIPASS 配下に展開される
_BASE = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
if not getattr(sys, "frozen", False):
    sys.path.insert(0, str(_BASE))
    sys.path.insert(0, str(_BASE.parent / "main_src"))
from api.bridge import Bridge  # noqa: E402

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(name)s %(levelname)s %(message)s")

STATIC_INDEX = _BASE / "static" / "index.html"


class WindowAdapter:
    """bridge へ注入するネイティブ機能。webview 依存はここに閉じ込める"""

    def __init__(self):
        self.window = None  # start 後に代入

    def open_file_dialog(self, file_types=None, directory=""):
        result = self.window.create_file_dialog(
            webview.FileDialog.OPEN, directory=directory,
            file_types=file_types or ())
        return result[0] if result else None

    def open_files_dialog(self, file_types=None, directory=""):
        result = self.window.create_file_dialog(
            webview.FileDialog.OPEN, directory=directory,
            allow_multiple=True, file_types=file_types or ())
        return list(result) if result else None

    def open_folder_dialog(self, directory=""):
        result = self.window.create_file_dialog(
            webview.FileDialog.FOLDER, directory=directory)
        return result[0] if result else None

    def eval_js(self, code):
        self.window.evaluate_js(code)


def create_app():
    adapter = WindowAdapter()
    bridge = Bridge(window_adapter=adapter)
    window = webview.create_window(
        "採点侍",
        str(STATIC_INDEX),
        js_api=bridge,
        width=1200,
        height=760,
        min_size=(900, 600),
    )
    adapter.window = window
    return window, bridge


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--debug", action="store_true", help="開発者ツールを有効化")
    parser.add_argument("--smoke", action="store_true",
                        help="パッケージ検証: 主要モジュールと同梱ファイルを確認して終了")
    args = parser.parse_args()

    if args.smoke:
        # exe ビルドの欠品検知（v4.5.1 の「依存が抜けたまま exe が公開される」
        # 事故の再発防止）。ウィンドウは開かず import と同梱物だけ検証する。
        # windowed exe では stdout が届かないため、結果は exe と同じ場所の
        # smoke_result.txt にも書き、成否は終了コードで返す
        result_path = Path(sys.executable if getattr(sys, "frozen", False)
                           else __file__).resolve().parent / "smoke_result.txt"
        try:
            from constants import APP_VERSION
            import cv2, numpy, pandas, openpyxl, sklearn, PIL  # noqa: F401
            from api.bridge import WEBUI_VERSION
            assert STATIC_INDEX.exists(), f"static が同梱されていません: {STATIC_INDEX}"
            bridge = Bridge(window_adapter=None)
            assert bridge.ping()["ok"]
            msg = f"SMOKE OK SaitenSamurai {APP_VERSION} (webui {WEBUI_VERSION})"
            result_path.write_text(msg + "\n", encoding="utf-8")
            print(msg)
            sys.exit(0)
        except Exception:
            import traceback
            result_path.write_text("SMOKE FAILED\n" + traceback.format_exc(),
                                   encoding="utf-8")
            traceback.print_exc()
            sys.exit(1)

    create_app()
    webview.start(debug=args.debug)


if __name__ == "__main__":
    main()
