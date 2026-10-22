"""uilock.py — 「採点侍の画面がこの試験を開いている」ことを示す印（ロックファイル）。

LLM と協働で採点するときは交代制にする: 画面を開いている間は先生の番、
閉じたら LLM の番。画面は記述の得点ファイル全体を上書き保存するため、
開いている間に LLM が書き込むと、その内容が黙って消える。

app.py --folder で開いたときにこの印を置き、閉じたら消す。
LLM 用の CLI（saiten_llm）は、印があって持ち主のプロセスが生きている間は
採点のファイルに書き込まない。
"""
import json
import os
import sys
from datetime import datetime
from pathlib import Path

LOCK_NAME = "ui_open.lock"


def lock_path(image_folder):
    from constants import RESULTS_FOLDER
    return Path(image_folder) / RESULTS_FOLDER / LOCK_NAME


def pid_alive(pid):
    """そのプロセスがまだ動いているか（Windows と Linux/macOS の両方）"""
    if not isinstance(pid, int) or pid <= 0:
        return False
    if sys.platform == "win32":
        import ctypes
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        STILL_ACTIVE = 259
        kernel32 = ctypes.windll.kernel32  # type: ignore[attr-defined]
        handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if not handle:
            return False
        try:
            code = ctypes.c_ulong()
            if not kernel32.GetExitCodeProcess(handle, ctypes.byref(code)):
                return False
            return code.value == STILL_ACTIVE
        finally:
            kernel32.CloseHandle(handle)
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


def acquire(image_folder):
    path = lock_path(image_folder)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({
        "pid": os.getpid(),
        "opened_at": datetime.now().astimezone().isoformat(timespec="seconds"),
    }), encoding="utf-8")
    return path


def release(image_folder):
    path = lock_path(image_folder)
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if data.get("pid") == os.getpid():
            path.unlink()
    except (OSError, ValueError):
        pass


def holder(image_folder):
    """画面が開いていれば {pid, opened_at}、開いていなければ None。

    持ち主のプロセスが落ちて印だけ残った場合は、開いていないとみなす
    """
    path = lock_path(image_folder)
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    return data if pid_alive(data.get("pid")) else None
