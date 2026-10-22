"""engine.py — 採点侍の Bridge を画面なしで動かす。

採点の処理は画面（webui）と同じ Bridge を通す。設定は採点侍の
session_state.json に保存されるので、LLM が準備した状態を
先生は採点侍の画面でそのまま開ける（app.py --folder）。
"""
import json
import time

from saiten_llm.workspace import Project


class ScreenOpenError(RuntimeError):
    """採点侍の画面がこの試験を開いている（先生の番）"""


class CliAdapter:
    """Bridge に渡す偽の画面。ダイアログはあらかじめ決めた答えを返す"""

    def __init__(self):
        self.folder_returns = []
        self.file_returns = []
        self.files_returns = []
        self.events = []

    def open_folder_dialog(self, directory=""):
        return self.folder_returns.pop(0) if self.folder_returns else None

    def open_file_dialog(self, file_types=None, directory=""):
        return self.file_returns.pop(0) if self.file_returns else None

    def open_files_dialog(self, file_types=None, directory=""):
        return self.files_returns.pop(0) if self.files_returns else None

    def eval_js(self, code):
        start = code.index("(") + 1
        self.events.append(json.loads(code[start:-1]))


def ensure_screen_closed(project: Project):
    """画面が開いている間は書き込まない（開いていれば ScreenOpenError）"""
    from api import uilock
    info = uilock.holder(project.scans)
    if info:
        raise ScreenOpenError(
            f"採点侍の画面がこの試験を開いています（{info.get('opened_at', '')} から）。"
            "先生が画面を閉じてから、もう一度実行してください。")


def session_path(project: Project):
    from constants import RESULTS_FOLDER, RESULTS_DATA_FOLDER, SESSION_STATE_FILE
    return project.scans / RESULTS_FOLDER / RESULTS_DATA_FOLDER / SESSION_STATE_FILE


def open_bridge(project: Project, write=True):
    """この試験の Bridge を用意する（前回の設定があれば引き継ぐ）。

    write=True なら、画面が開いていないことを先に確かめる
    """
    from api.bridge import Bridge
    if write:
        ensure_screen_closed(project)
    if not project.scan_files():
        raise FileNotFoundError("答案の画像がありません。先に pdf を実行してください")
    adapter = CliAdapter()
    bridge = Bridge(window_adapter=adapter)
    path = session_path(project)
    if path.exists():
        res = bridge.restore_session(str(path))
        if not res["ok"]:
            raise RuntimeError(res["error"])
        for w in res.get("warnings", []):
            print(f"⚠ {w}")
    else:
        adapter.folder_returns = [str(project.scans)]
        res = bridge.select_image_folder()
        if not res["ok"]:
            raise RuntimeError(res["error"])
        bridge.save_session()
    return bridge, adapter


def call(bridge, name, *args):
    """Bridge のメソッドを呼び、失敗なら RuntimeError にする"""
    res = getattr(bridge, name)(*args)
    if not res.get("ok"):
        raise RuntimeError(res.get("error", f"{name} に失敗しました"))
    return res


def run_job(bridge, adapter, name, *args, timeout=3600):
    """時間のかかる処理を始めて、終わるまで待つ。Returns: job_done の内容"""
    before = len(adapter.events)
    call(bridge, name, *args)
    deadline = time.monotonic() + timeout
    last_report = 0
    while time.monotonic() < deadline:
        done = [e for e in adapter.events[before:] if e.get("type") == "job_done"]
        if done and not bridge.state["job"]["running"]:
            event = done[-1]
            if not event.get("ok"):
                raise RuntimeError(event.get("message", "処理に失敗しました"))
            return event
        job = bridge.state["job"]
        if job["total"] and time.monotonic() - last_report > 15:
            print(f"  …{job['current']} / {job['total']}", flush=True)
            last_report = time.monotonic()
        time.sleep(0.2)
    bridge.cancel_job()
    raise TimeoutError("処理が時間内に終わりませんでした")
