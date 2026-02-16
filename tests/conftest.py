"""
pytest 共通設定: main_src をインポートパスに追加 + Tk ルート管理
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "main_src"))


# ================================================================
# Tk ルートウィンドウの一元管理
# ================================================================
# 各テストファイルが独自に tk.Tk() / root.destroy() を行うと、
# Tcl インタプリタが破壊され後続テストで tk.Tk() が失敗する。
# このフィクスチャでセッション全体で 1 つの Tk ルートを共有する。

import tkinter as tk

_shared_root = None


def get_shared_tk_root():
    """セッション共通の Tk ルートを取得（遅延初期化）"""
    global _shared_root
    if _shared_root is not None:
        try:
            _shared_root.winfo_exists()
        except tk.TclError:
            _shared_root = None
    if _shared_root is None:
        _shared_root = tk.Tk()
        _shared_root.withdraw()
    return _shared_root


def pytest_sessionfinish(session, exitstatus):
    """pytest セッション終了時に Tk ルートを安全に破棄"""
    global _shared_root
    if _shared_root is not None:
        try:
            _shared_root.destroy()
        except Exception:
            pass
        _shared_root = None
