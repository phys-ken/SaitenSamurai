"""jobs.py — 長時間処理（認識・採点・集計）のスレッド実行と進捗push。

設計（webui/docs/plan.md）:
- pywebview を import しない。JS への push は adapter.eval_js 経由
  （app.py が window.evaluate_js を、テストが記録用の偽物を注入する）
- 同時に走るジョブは 1 つ（tk 版の _set_processing_state と同じ思想）
- 中断は main_src ドライバが既に持つ cancel_event をそのまま使う
"""
import json
import logging
import threading
from pathlib import Path

logger = logging.getLogger(__name__)


def _ok(**data):
    return {"ok": True, **data}


def _err(message):
    return {"ok": False, "error": str(message)}


class JobsMixin:

    def _init_jobs(self):
        self.state.update({
            "omr_mode": "kmeans",        # tk 版の既定（推奨クラスタリング）
            "color_threshold": 0.1,
            "area_threshold": 0.4,
            "omr_result": None,          # 認識結果 xlsx（自動設定 or 手動選択）
            "job": {"running": False, "kind": None, "current": 0, "total": 0},
        })
        self._cancel_event = threading.Event()
        self._job_thread = None

    # --- JS への push ----------------------------------------------

    def _push(self, event_type, **payload):
        """window.saitenEvents({type, ...}) を呼ぶ。失敗しても処理は止めない"""
        data = json.dumps({"type": event_type, **payload}, ensure_ascii=False)
        try:
            self._win.eval_js(f"window.saitenEvents && window.saitenEvents({data})")
        except Exception as e:  # 画面が閉じた直後など
            logger.debug("eval_js push 失敗: %s", e)

    # --- 設定 -------------------------------------------------------

    def set_omr_mode(self, mode):
        from constants import OMR_MODE_THRESHOLD, OMR_MODE_KMEANS
        if mode not in (OMR_MODE_THRESHOLD, OMR_MODE_KMEANS):
            return _err(f"不明な認識方式です: {mode}")
        self.state["omr_mode"] = mode
        return _ok(state=self.state)

    def set_thresholds(self, color_threshold, area_threshold):
        try:
            c, a = float(color_threshold), float(area_threshold)
            if not (0.0 < c < 1.0 and 0.0 < a < 1.0):
                raise ValueError
        except (TypeError, ValueError):
            return _err("しきい値は 0〜1 の間の数値で指定してください")
        self.state["color_threshold"] = c
        self.state["area_threshold"] = a
        return _ok(state=self.state)

    def select_omr_result(self):
        """OMR結果 xlsx の手動選択（過去の結果で採点し直すケース）"""
        from api.bridge import _XLSX_FILE_TYPES
        path = self._win.open_file_dialog(file_types=_XLSX_FILE_TYPES)
        if not path:
            return _ok(state=self.state, cancelled=True)
        self.state["omr_result"] = str(path)
        return _ok(state=self.state)

    # --- 認識実行 ---------------------------------------------------

    def run_recognition(self):
        if self.state["job"]["running"]:
            return _err("別の処理が実行中です。完了または中断を待ってください")
        if not self.state["image_folder"]:
            return _err("画像フォルダを選択してください")
        if not self.state["coord_file"]:
            return _err("座標ファイルを選択してください")

        self._cancel_event.clear()
        self.state["job"] = {"running": True, "kind": "recognition",
                             "current": 0, "total": self.state["image_count"]}
        self._job_thread = threading.Thread(target=self._recognition_worker,
                                            daemon=True)
        self._job_thread.start()
        return _ok(state=self.state)

    def _recognition_worker(self):
        from omr_engine import process_box_drawer

        def on_progress(current, total):
            self.state["job"].update(current=current, total=total)
            self._push("progress", kind="recognition", current=current, total=total)

        try:
            result = process_box_drawer(
                self.state["image_folder"],
                self.state["coord_file"],
                skip_questions=self.state["skip_questions"],
                color_threshold=self.state["color_threshold"],
                area_threshold=self.state["area_threshold"],
                progress_callback=on_progress,
                cancel_event=self._cancel_event,
                omr_mode=self.state["omr_mode"],
                mark_format=self.state["mark_format"],
            )
            cancelled = self._cancel_event.is_set()
            if not cancelled:
                self._autoselect_omr_result()
            done_event = dict(
                kind="recognition", ok=True, cancelled=cancelled,
                success_count=result.get("success_count", 0),
                error_count=result.get("error_count", 0),
                omr_result=self.state["omr_result"],
                message=("中断しました" if cancelled else
                         f"認識完了: 成功 {result.get('success_count', 0)} 件 / "
                         f"エラー {result.get('error_count', 0)} 件"),
            )
        except Exception as e:
            logger.exception("認識実行に失敗")
            done_event = dict(kind="recognition", ok=False,
                              message=f"認識実行に失敗しました: {e}")
        # push は state リセットの「後」に行う。先に push すると JS 側の
        # get_state が running=True の古い状態を読む競合が起き、完了後も
        # 進捗バーと中断ボタンが画面に残る（demo キャプチャで実際に発生）
        self.state["job"] = {"running": False, "kind": None,
                             "current": 0, "total": 0}
        self._push("job_done", **done_event)

    def _autoselect_omr_result(self):
        """最新の Mark2-Result xlsx を自動選択（tk 版 main_gui.py:1079 と同じ）"""
        from constants import (RESULTS_FOLDER, RESULTS_DATA_FOLDER,
                               READING_RESULTS_FOLDER_NAME)
        folder = (Path(self.state["image_folder"]) / RESULTS_FOLDER /
                  RESULTS_DATA_FOLDER / READING_RESULTS_FOLDER_NAME)
        candidates = sorted(folder.glob("Mark2-Result-*.xlsx"))
        if candidates:
            self.state["omr_result"] = str(candidates[-1])

    def cancel_job(self):
        if not self.state["job"]["running"]:
            return _ok(state=self.state)
        self._cancel_event.set()
        return _ok(state=self.state)
