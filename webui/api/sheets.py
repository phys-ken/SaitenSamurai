"""sheets.py — 答案全体画像の提供と、画像上の位置/領域設定。

region-picker（JS側の汎用ドラッグ選択）に全体画像を渡し、選ばれた
矩形/位置を各設定（合計点位置・氏名欄・記述領域）として保存する層。
座標系は tk 版と同じ「00_Processing 内画像のピクセル座標」。
"""
import base64
import logging
from pathlib import Path

logger = logging.getLogger(__name__)


def _ok(**data):
    return {"ok": True, **data}


def _err(message):
    return {"ok": False, "error": str(message)}


class SheetsMixin:

    def _boxed_folder(self):
        from constants import RESULTS_FOLDER, BOXED_FOLDER
        if not self.state["image_folder"]:
            return None
        return Path(self.state["image_folder"]) / RESULTS_FOLDER / BOXED_FOLDER

    def _results_data_folder(self):
        from constants import RESULTS_FOLDER, RESULTS_DATA_FOLDER
        return (Path(self.state["image_folder"]) / RESULTS_FOLDER /
                RESULTS_DATA_FOLDER)

    def list_sheet_files(self):
        """補正済み画像（00_Processing）の一覧"""
        folder = self._boxed_folder()
        if not folder or not folder.exists():
            return _err("補正済み画像がありません。先に認識（または画像準備）を実行してください")
        files = sorted(p.name for p in folder.iterdir()
                       if p.suffix.lower() in ('.jpg', '.jpeg', '.png'))
        if not files:
            return _err("補正済み画像がありません。先に認識（または画像準備）を実行してください")
        return _ok(files=files)

    def get_sheet_image(self, filename=None):
        """答案全体画像を data URL で返す（filename 省略時は先頭の1枚）"""
        listing = self.list_sheet_files()
        if not listing["ok"]:
            return listing
        name = filename or listing["files"][0]
        if name not in listing["files"]:
            return _err(f"画像が見つかりません: {name}")
        path = self._boxed_folder() / name
        suffix = path.suffix.lower().lstrip(".")
        mime = "jpeg" if suffix in ("jpg", "jpeg") else suffix
        b64 = base64.b64encode(path.read_bytes()).decode("ascii")
        return _ok(filename=name,
                   data_url=f"data:image/{mime};base64,{b64}")

    # --- 合計点表示位置 ---------------------------------------------

    def _load_total_display_state(self):
        """image_folder 選択/復元時に保存済みの合計点表示位置を state へ載せる"""
        from descriptive_scorer import (load_total_display_config,
                                        TOTAL_DISPLAY_CONFIG_FILE)
        cfg = load_total_display_config(
            str(self._results_data_folder() / TOTAL_DISPLAY_CONFIG_FILE))
        self.state["total_display_region"] = (
            cfg.get("total_display_region") if cfg else None)

    def get_total_display_region(self):
        from descriptive_scorer import (load_total_display_config,
                                        TOTAL_DISPLAY_CONFIG_FILE)
        if not self.state["image_folder"]:
            return _err("画像フォルダを選択してください")
        cfg = load_total_display_config(
            str(self._results_data_folder() / TOTAL_DISPLAY_CONFIG_FILE))
        region = cfg.get("total_display_region") if cfg else None
        return _ok(region=region)

    def set_total_display_region(self, region):
        """合計点表示位置を保存（region=[x1,y1,x2,y2]、None でリセット）"""
        from descriptive_scorer import (save_total_display_config,
                                        TOTAL_DISPLAY_CONFIG_FILE)
        if not self.state["image_folder"]:
            return _err("画像フォルダを選択してください")
        data_folder = self._results_data_folder()
        config_path = data_folder / TOTAL_DISPLAY_CONFIG_FILE
        if region is None:
            if config_path.exists():
                config_path.unlink()
            self.state["total_display_region"] = None
            return _ok(region=None, state=self.state)
        try:
            region = [int(v) for v in region]
            assert len(region) == 4
        except (TypeError, ValueError, AssertionError):
            return _err("領域は [x1, y1, x2, y2] の数値で指定してください")
        data_folder.mkdir(parents=True, exist_ok=True)
        save_total_display_config(str(config_path), region)
        self.state["total_display_region"] = region
        return _ok(region=region, state=self.state)

    # --- 氏名トリミング（集計シートに氏名画像を表示） --------------

    def set_name_trim_enabled(self, enabled):
        self.state["name_trim_enabled"] = bool(enabled)
        return _ok(state=self.state)

    def set_name_trim_region(self, region):
        """氏名欄の領域を設定（None で解除）。00_Processing 座標系"""
        if region is None:
            self.state["name_trim_region"] = None
            return _ok(state=self.state)
        try:
            region = [int(v) for v in region]
            assert len(region) == 4
        except (TypeError, ValueError, AssertionError):
            return _err("領域は [x1,y1,x2,y2] で指定してください")
        x1, y1, x2, y2 = region
        if x2 <= x1 or y2 <= y1:
            return _err("領域の幅と高さは正である必要があります")
        self.state["name_trim_region"] = region
        return _ok(state=self.state)

    # --- 採点結果の表示項目（描画詳細設定） --------------------------

    def set_rendering_settings(self, overrides):
        """描画詳細設定を更新する。既知のキーだけ受け付け、型を検証する。

        state["rendering_settings"] には既定値との差分だけを保持する
        （セッションに tk 互換の形でそのまま保存されるため）。
        """
        from constants import DEFAULT_RENDERING_SETTINGS
        if not isinstance(overrides, dict):
            return _err("設定は {キー: 値} の形式で指定してください")
        current = dict(self.state["rendering_settings"] or {})
        for key, value in overrides.items():
            if key not in DEFAULT_RENDERING_SETTINGS:
                return _err(f"不明な設定項目です: {key}")
            default = DEFAULT_RENDERING_SETTINGS[key]
            if isinstance(default, bool):
                value = bool(value)
            elif isinstance(default, float):
                try:
                    value = float(value)
                except (TypeError, ValueError):
                    return _err(f"{key} は数値で指定してください")
            current[key] = value
        # 既定値と同じものは差分から落とす
        current = {k: v for k, v in current.items()
                   if v != DEFAULT_RENDERING_SETTINGS[k]}
        self.state["rendering_settings"] = current or None
        self._save_session_quietly()
        return _ok(state=self.state)

    def get_rendering_settings(self):
        """既定値に差分を適用した完全な設定辞書と、既定値そのものを返す"""
        from constants import get_rendering_settings, DEFAULT_RENDERING_SETTINGS
        return _ok(settings=get_rendering_settings(self.state["rendering_settings"]),
                   defaults=dict(DEFAULT_RENDERING_SETTINGS))

    def get_render_preview(self):
        """マーク描き込みのプレビュー画像を返す（tk 版「位置プレビュー」の後継）。

        最初の答案（00_Processing）に、現在の表示項目設定で
        「正解例・不正解例・全員正解（★）例」をダミー採点結果として描画し、
        該当行の周辺を切り出して data URL で返す。実際の描画関数を使うので
        オフセット・白塗り・各トグルの効き方がそのまま確認できる。
        """
        if self.state["job"]["running"]:
            return _err("別の処理が実行中はプレビューできません")
        import base64
        import cv2
        import numpy as np
        if not self.state["coord_file"]:
            return _err("座標ファイルを選択するとプレビューできます")
        folder = self._boxed_folder()
        files = (sorted(p.name for p in folder.iterdir()
                        if p.suffix.lower() in ('.jpg', '.jpeg', '.png'))
                 if folder and folder.exists() else [])
        if not files:
            return _err("認識実行の後にプレビューできます（補正済み画像を使うため）")

        from omr_engine import parse_excel_coordinates
        from image_renderer import draw_scoring_results
        skip = self.state["skip_questions"]
        coordinates, _ = parse_excel_coordinates(self.state["coord_file"], skip)
        q_nos = sorted({c["question_no"] for c in coordinates})
        if not q_nos:
            return _err("座標ファイルに設問がありません")

        img = cv2.imdecode(np.fromfile(str(folder / files[0]), dtype=np.uint8),
                           cv2.IMREAD_COLOR)
        if img is None:
            return _err("補正済み画像を読み込めません")

        # ダミー採点結果: 1問目=正解 / 2問目=不正解 / 3問目=全員正解(★)
        samples = [
            dict(correct=True, points=3, aspect=1, correct_answer="2"),
            dict(correct=False, points=3, aspect=1, correct_answer="2"),
            dict(correct=True, points=3, aspect=1, correct_answer="2",
                 special="全員正解"),
        ]
        results = {}
        for q_abs, sample in zip(q_nos, samples):
            results[q_abs - skip] = sample
        scoring_result = {
            "results": results,
            "total_score": 3, "max_score": 9,
            "aspect_scores": {1: 3}, "aspect_max_scores": {1: 9},
        }
        drawn = draw_scoring_results(
            img, coordinates, scoring_result, skip_questions=skip,
            output_scale=1.0,
            rendering_settings=self.state["rendering_settings"],
            mark_format=self.state["mark_format"])

        # 使った設問行の周辺を切り出す（左右にはみ出しぶんの余白を持たせる）
        used = [c for c in coordinates
                if c["question_no"] in list(results.keys()) or
                   c["question_no"] in [q + skip for q in results.keys()]]
        ys = [c["y"] for c in used] + [c["y"] + c["height"] for c in used]
        xs = [c["x"] for c in used] + [c["x"] + c["width"] for c in used]
        h = img.shape[0]; w = img.shape[1]
        row_h = used[0]["height"]
        y1 = max(0, int(min(ys) - row_h * 0.8))
        y2 = min(h, int(max(ys) + row_h * 0.8))
        cell_w = used[0]["width"]
        x1 = max(0, int(min(xs) - cell_w * 3))
        x2 = min(w, int(max(xs) + cell_w * 3))
        crop = drawn[y1:y2, x1:x2]
        ok, buf = cv2.imencode(".png", crop)
        if not ok:
            return _err("プレビュー画像の生成に失敗しました")
        b64 = base64.b64encode(buf.tobytes()).decode("ascii")
        return _ok(data_url=f"data:image/png;base64,{b64}",
                   sample_note="上から: 正解例 / 不正解例 / 全員正解（★）の例")

    def get_progress(self):
        """ステッパー用の進行度（既存 state と出力フォルダからの読み取り専用導出）"""
        from constants import (RESULTS_FOLDER, SCORED_FOLDER,
                               FINAL_REPORT_FOLDER, STUDENT_SUMMARY_FILE)
        mode = self.state["app_mode"]
        if mode == "descriptive_only":
            prepared = bool(self.state["image_folder"])
            read_done = ((self.state.get("descriptive") or {})
                         .get("prepared_count", 0) or 0) > 0
        else:
            prepared = bool(self.state["image_folder"] and
                            self.state["coord_file"])
            read_done = bool(self.state["omr_result"])
        scored = False
        summarized = False
        if self.state["image_folder"]:
            base = Path(self.state["image_folder"]) / RESULTS_FOLDER
            sf = base / SCORED_FOLDER
            scored = sf.exists() and any(
                p.suffix.lower() in ('.jpg', '.jpeg', '.png')
                for p in sf.iterdir())
            summarized = (base / FINAL_REPORT_FOLDER /
                          STUDENT_SUMMARY_FILE).exists()
        return _ok(progress={"prepared": prepared, "read": read_done,
                             "scored": scored, "summarized": summarized})

    def set_include_descriptive_in_analysis(self, enabled):
        self.state["include_descriptive_in_analysis"] = bool(enabled)
        return _ok(state=self.state)

    # --- フォルダを開く（tk 版の 📁 ボタン群相当） -------------------

    def open_folder(self, kind):
        """結果フォルダ等を OS のファイルマネージャで開く。

        kind: "boxed"(補正済み画像) / "scored"(採点済み答案) /
              "report"(集計レポート) / "results_data"(正答データ等)
        """
        import os
        import subprocess
        from constants import (RESULTS_FOLDER, BOXED_FOLDER, SCORED_FOLDER,
                               FINAL_REPORT_FOLDER, RESULTS_DATA_FOLDER)
        if not self.state["image_folder"]:
            return _err("画像フォルダを選択してください")
        sub = {"boxed": BOXED_FOLDER, "scored": SCORED_FOLDER,
               "report": FINAL_REPORT_FOLDER,
               "results_data": RESULTS_DATA_FOLDER}.get(kind)
        if sub is None:
            return _err(f"不明なフォルダ種別です: {kind}")
        target = Path(self.state["image_folder"]) / RESULTS_FOLDER / sub
        if not target.exists():
            return _err(f"フォルダがまだありません: {target}")
        try:
            if os.name == "nt":
                os.startfile(str(target))  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["xdg-open", str(target)])
        except Exception as e:
            return _err(f"フォルダを開けませんでした: {e}")
        return _ok(path=str(target))

    def open_original_image(self, filename):
        """答案の元画像を OS の既定ビューアで開く（tk 📷 相当）"""
        import os
        import subprocess
        if not self.state["image_folder"]:
            return _err("画像フォルダを選択してください")
        target = Path(self.state["image_folder"]) / filename
        if not target.exists():
            return _err(f"元画像が見つかりません: {target}")
        try:
            if os.name == "nt":
                os.startfile(str(target))  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["xdg-open", str(target)])
        except Exception as e:
            return _err(f"画像を開けませんでした: {e}")
        return _ok(path=str(target))
