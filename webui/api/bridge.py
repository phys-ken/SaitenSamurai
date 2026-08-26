"""bridge.py — JS から呼ばれる API 層。

設計原則（webui/docs/plan.md）:
- このモジュールは **pywebview を import しない**。ネイティブダイアログ等の
  ウィンドウ依存機能は app.py から注入される adapter 経由で使う。
  → CI や pywebview の無い環境でもそのまま pytest できる
- 各メソッドは JSON 化可能な dict を返す。成功は {"ok": True, ...}、
  失敗は {"ok": False, "error": "利用者向けの日本語メッセージ"}。
  例外を JS 側に漏らさない（漏らすと画面ごとに握り方がバラつく）
- 採点ロジックは main_src のドライバを呼ぶだけ。ここにロジックを書かない
"""
import logging
import sys
from pathlib import Path

# main_src をインポートパスへ（リポジトリ同居構成）
_REPO_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_REPO_ROOT / "main_src") not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT / "main_src"))

from constants import APP_VERSION  # noqa: E402

logger = logging.getLogger(__name__)

WEBUI_VERSION = "0.1.0"


def _ok(**data):
    return {"ok": True, **data}


def _err(message):
    return {"ok": False, "error": str(message)}


class Bridge:
    """pywebview の js_api として渡すクラス。

    window_adapter: ネイティブ機能の注入点。必要なメソッド:
        - open_file_dialog(file_types, directory) -> str | None
        - open_folder_dialog(directory) -> str | None
      app.py が pywebview 実装を、テストが偽物を渡す。
    """

    def __init__(self, window_adapter=None):
        self._win = window_adapter

    # --- 疎通・情報 -------------------------------------------------

    def ping(self):
        """疎通確認。JS 側の起動シーケンスがブリッジ準備完了を検知するのに使う"""
        return _ok(pong=True)

    def get_app_info(self):
        return _ok(
            app_version=APP_VERSION,
            webui_version=WEBUI_VERSION,
            python=sys.version.split()[0],
        )
