"""app.py — webui のエントリポイント。

このファイルだけが pywebview を import する（bridge は import しない —
webui/docs/plan.md の設計原則）。起動: python webui/app.py
"""
import argparse
import logging
import sys
from pathlib import Path

import webview

sys.path.insert(0, str(Path(__file__).resolve().parent))
from api.bridge import Bridge  # noqa: E402

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(name)s %(levelname)s %(message)s")

STATIC_INDEX = Path(__file__).resolve().parent / "static" / "index.html"


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
    args = parser.parse_args()

    create_app()
    webview.start(debug=args.debug)


if __name__ == "__main__":
    main()
