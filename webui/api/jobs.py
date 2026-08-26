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

    def _start_job(self, kind, worker):
        """ジョブ共通の起動処理（同時実行1つ・cancel初期化・スレッド化）"""
        if self.state["job"]["running"]:
            return _err("別の処理が実行中です。完了または中断を待ってください")
        self._cancel_event.clear()
        self.state["job"] = {"running": True, "kind": kind,
                             "current": 0, "total": self.state["image_count"]}

        def wrapped():
            try:
                done_event = worker()
            except Exception as e:
                logger.exception("%s に失敗", kind)
                done_event = dict(kind=kind, ok=False,
                                  message=f"処理に失敗しました: {e}")
            # push は state リセットの「後」。先に push すると JS の get_state が
            # running=True の古い状態を読み、完了後も進捗バーが残る競合になる
            self.state["job"] = {"running": False, "kind": None,
                                 "current": 0, "total": 0}
            self._push("job_done", **done_event)

        self._job_thread = threading.Thread(target=wrapped, daemon=True)
        self._job_thread.start()
        return _ok(state=self.state)

    def _progress_cb(self, kind):
        def on_progress(current, total):
            self.state["job"].update(current=current, total=total)
            self._push("progress", kind=kind, current=current, total=total)
        return on_progress

    # --- 認識 -------------------------------------------------------

    def run_recognition(self):
        if not self.state["image_folder"]:
            return _err("画像フォルダを選択してください")
        if not self.state["coord_file"]:
            return _err("座標ファイルを選択してください")
        return self._start_job("recognition", self._recognition_worker)

    def _recognition_worker(self):
        from omr_engine import process_box_drawer
        result = process_box_drawer(
            self.state["image_folder"],
            self.state["coord_file"],
            skip_questions=self.state["skip_questions"],
            color_threshold=self.state["color_threshold"],
            area_threshold=self.state["area_threshold"],
            progress_callback=self._progress_cb("recognition"),
            cancel_event=self._cancel_event,
            omr_mode=self.state["omr_mode"],
            mark_format=self.state["mark_format"],
        )
        cancelled = self._cancel_event.is_set()
        if not cancelled:
            self._autoselect_omr_result()
        return dict(
            kind="recognition", ok=True, cancelled=cancelled,
            success_count=result.get("success_count", 0),
            error_count=result.get("error_count", 0),
            omr_result=self.state["omr_result"],
            message=("中断しました" if cancelled else
                     f"認識完了: 成功 {result.get('success_count', 0)} 件 / "
                     f"エラー {result.get('error_count', 0)} 件"),
        )

    # --- 採点（採点済み答案の生成） ---------------------------------

    def run_scoring(self):
        mode = self.state["app_mode"]
        if not self.state["image_folder"]:
            return _err("画像フォルダを選択してください")
        if mode != "descriptive_only":
            if not self.state["coord_file"]:
                return _err("座標ファイルを選択してください")
            if not self.state["answer_key"]:
                return _err("正答データを選択してください")
            if not self.state["omr_result"]:
                return _err("OMR結果がありません。先に認識を実行してください")
            ks = self.state.get("key_summary")
            if ks and not ks["ok"]:
                return _err("正答データにエラーがあります。修正してから採点してください:\n"
                            + "\n".join(ks["errors"][:3]))
        if mode in ("mark_and_descriptive", "descriptive_only"):
            desc = self.state.get("descriptive")
            if not desc or not desc["questions"]:
                return _err("記述問題が設定されていません。先に「記述問題設定」を行ってください")
        return self._start_job("scoring", self._scoring_worker)

    def _scoring_worker(self):
        """採点済み答案の生成。tk 版と同じモード分岐:
        mark_only → process_scoring / mark_and_descriptive →
        generate_return_sheets（マーク＋記述の合成描画） /
        descriptive_only → generate_descriptive_only_sheets"""
        from constants import RESULTS_FOLDER, SCORED_FOLDER, BOXED_FOLDER
        mode = self.state["app_mode"]
        out = Path(self.state["image_folder"]) / RESULTS_FOLDER / SCORED_FOLDER

        if mode == "mark_only":
            from image_renderer import process_scoring
            process_scoring(
                self.state["image_folder"],
                self.state["coord_file"],
                self.state["answer_key"],
                self.state["omr_result"],
                skip_questions=self.state["skip_questions"],
                progress_callback=self._progress_cb("scoring"),
                cancel_event=self._cancel_event,
                mark_format=self.state["mark_format"],
            )
        elif mode == "mark_and_descriptive":
            from descriptive_scorer import generate_return_sheets
            generate_return_sheets(
                image_folder=self.state["image_folder"],
                config=self._desc_config,
                descriptive_scores=self._desc_scores["scores"],
                coord_excel_path=self.state["coord_file"],
                template_path=self.state["answer_key"],
                mark2_result_path=self.state["omr_result"],
                skip_questions=self.state["skip_questions"],
                output_folder=str(out),
                mark_format=self.state["mark_format"],
            )
        else:  # descriptive_only
            from descriptive_scorer import generate_descriptive_only_sheets
            boxed = Path(self.state["image_folder"]) / RESULTS_FOLDER / BOXED_FOLDER
            generate_descriptive_only_sheets(
                str(boxed), self._desc_config, self._desc_scores["scores"],
                str(out))
        cancelled = self._cancel_event.is_set()
        return dict(kind="scoring", ok=True, cancelled=cancelled,
                    message=("中断しました" if cancelled
                             else f"採点済み答案の生成が完了しました → {out}"))

    # --- 集計 -------------------------------------------------------

    def run_summary(self):
        mode = self.state["app_mode"]
        if not self.state["image_folder"]:
            return _err("画像フォルダを選択してください")
        if mode == "descriptive_only":
            desc = self.state.get("descriptive")
            if not desc or not desc["questions"]:
                return _err("記述問題が設定されていません。先に「記述問題設定」を行ってください")
            return self._start_job("summary", self._summary_worker)
        if not self.state["coord_file"]:
            return _err("座標ファイルを選択してください")
        if not self.state["answer_key"]:
            return _err("正答データを選択してください")
        if not self.state["omr_result"]:
            return _err("OMR結果がありません。先に認識を実行してください")
        return self._start_job("summary", self._summary_worker)

    def _summary_worker(self):
        from constants import RESULTS_FOLDER, FINAL_REPORT_FOLDER
        mode = self.state["app_mode"]
        out = str(Path(self.state["image_folder"]) / RESULTS_FOLDER / FINAL_REPORT_FOLDER)

        if mode == "descriptive_only":
            from summary_generator import process_descriptive_only_summary
            result = process_descriptive_only_summary(
                self.state["image_folder"],
                self._desc_config,
                self._desc_scores["scores"],
            )
            if not result.get("success", True):
                return dict(kind="summary", ok=False,
                            message=f"集計に失敗しました: {result.get('error')}")
        else:
            from summary_generator import process_summary_generation
            desc = self.state.get("descriptive")
            has_desc = (mode == "mark_and_descriptive" and desc and
                        desc["questions"])
            process_summary_generation(
                self.state["image_folder"],
                self.state["coord_file"],
                self.state["answer_key"],
                self.state["omr_result"],
                skip_questions=self.state["skip_questions"],
                descriptive_config=self._desc_config if has_desc else None,
                descriptive_scores=(self._desc_scores["scores"]
                                    if has_desc else None),
                include_descriptive_in_analysis=bool(has_desc),
                progress_callback=self._progress_cb("summary"),
                cancel_event=self._cancel_event,
                mark_format=self.state["mark_format"],
            )
        cancelled = self._cancel_event.is_set()
        return dict(kind="summary", ok=True, cancelled=cancelled,
                    output_folder=out,
                    message=("中断しました" if cancelled
                             else f"集計が完了しました → {out}"))

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
