"""L2テスト設定 — システムの Chrome で webui/static を直接開いてUIを検証する。

pywebview は使わない。bridge.js の window.__mockApi 注入点にモックを差し、
「UIが正しく振る舞うか」だけを高速に固定する（webui/docs/plan.md L2層）。
Chrome は WebView2 と同じ Chromium 系なので、エンジン差の面でも本番に近い。

file:// で ES modules を読むため --allow-file-access-from-files を付ける
（テスト専用フラグ。アプリ本体には不要 — pywebview のエンジンは許可している）。
"""
from pathlib import Path

import pytest

STATIC_DIR = Path(__file__).resolve().parent.parent / "static"
INDEX_URL = (STATIC_DIR / "index.html").as_uri()


@pytest.fixture(scope="session")
def browser_type_launch_args(browser_type_launch_args):
    return {
        **browser_type_launch_args,
        "channel": "chrome",
        "args": ["--allow-file-access-from-files"],
    }


@pytest.fixture
def open_app(page):
    """モックAPIを注入してから index.html を開くヘルパー"""
    def _open(mock_api_js: str):
        page.add_init_script(f"window.__mockApi = {mock_api_js};")
        page.goto(INDEX_URL)
        return page
    return _open
