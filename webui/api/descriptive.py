"""descriptive.py — 記述式採点の API 層。

設定・スコアの永続化は descriptive_scorer（ロジック層）をそのまま使い、
ファイル形式は tk 版と完全互換:
  - descriptive_config.json: {"questions": [{id, name, max_score, aspect, region}]}
  - descriptive_scores.json: {"version": 1, "scores": {filename: {qid: score}}}
切り出しは trim_descriptive_regions（元画像からの高解像度モード）。
"""
import base64
import logging
import shutil
from pathlib import Path

logger = logging.getLogger(__name__)


def _ok(**data):
    return {"ok": True, **data}


def _err(message):
    return {"ok": False, "error": str(message)}


class DescriptiveMixin:

    def _init_descriptive(self):
        self._desc_crops = None   # {qid: {filename: crop_path}}
        self.state["descriptive"] = None

    # --- パスとロード ----------------------------------------------

    def _desc_config_path(self):
        from descriptive_scorer import DESCRIPTIVE_CONFIG_FILE
        return self._results_data_folder() / DESCRIPTIVE_CONFIG_FILE

    def _desc_scores_path(self):
        from descriptive_scorer import DESCRIPTIVE_SCORES_FILE
        return self._results_data_folder() / DESCRIPTIVE_SCORES_FILE

    def _load_descriptive_state(self):
        """image_folder 選択後に呼び、既存の設定・スコアを state に反映"""
        from descriptive_scorer import (load_descriptive_config,
                                        load_descriptive_scores)
        if not self.state["image_folder"]:
            self.state["descriptive"] = None
            return
        config = load_descriptive_config(str(self._desc_config_path())) or \
            {"questions": []}
        scores_data = load_descriptive_scores(str(self._desc_scores_path())) or \
            {"version": 1, "scores": {}}
        self._desc_config = config
        self._desc_scores = scores_data
        self._refresh_desc_summary()

    def _refresh_desc_summary(self):
        listing = self.list_sheet_files()
        prepared = len(listing["files"]) if listing["ok"] else 0
        scores = self._desc_scores["scores"]
        per_q = {}
        for q in self._desc_config["questions"]:
            per_q[q["id"]] = sum(
                1 for f_scores in scores.values()
                if f_scores.get(q["id"]) is not None)
        self.state["descriptive"] = {
            "questions": self._desc_config["questions"],
            "scored_counts": per_q,
            "prepared_count": prepared,
        }

    # --- 画像準備（記述のみモード） ---------------------------------

    def run_prepare_images(self):
        """記述のみモード: 画像を 00_Processing へコピーする
        （tk 版 _run_prepare_images_thread と同じ「補正なしコピー」）"""
        if not self.state["image_folder"]:
            return _err("画像フォルダを選択してください")
        return self._start_job("prepare_images", self._prepare_images_worker)

    def _prepare_images_worker(self):
        from constants import RESULTS_FOLDER, BOXED_FOLDER, RESULTS_DATA_FOLDER
        src = Path(self.state["image_folder"])
        results = src / RESULTS_FOLDER
        boxed = results / BOXED_FOLDER
        boxed.mkdir(parents=True, exist_ok=True)
        (results / RESULTS_DATA_FOLDER).mkdir(parents=True, exist_ok=True)
        self._write_results_readme()
        files = sorted(p for p in src.iterdir()
                       if p.suffix.lower() in ('.jpg', '.jpeg', '.png'))
        for i, p in enumerate(files, 1):
            dst = boxed / p.name
            if not dst.exists() or dst.stat().st_mtime < p.stat().st_mtime:
                shutil.copy2(str(p), str(dst))
            self._progress_cb("prepare_images")(i, len(files))
        self._load_descriptive_state()
        return dict(kind="prepare_images", ok=True,
                    message=f"画像準備が完了しました（{len(files)}枚）。"
                            "次は「記述問題設定」で採点領域を設定してください")

    # --- 記述問題設定 ----------------------------------------------

    def add_descriptive_question(self, name, max_score, aspect, region):
        from descriptive_scorer import save_descriptive_config
        if self.state["descriptive"] is None:
            return _err("画像フォルダを選択してください")
        try:
            max_score = int(max_score)
            aspect = int(aspect)
            region = [int(v) for v in region]
            assert max_score > 0 and len(region) == 4
        except (TypeError, ValueError, AssertionError):
            return _err("配点は正の整数、領域は [x1,y1,x2,y2] で指定してください")
        qs = self._desc_config["questions"]
        # tk 版は連番 D{n}。既存の最大番号+1 で欠番衝突を避ける
        used = [int(q["id"][1:]) for q in qs
                if q["id"].startswith("D") and q["id"][1:].isdigit()]
        qid = f"D{max(used) + 1 if used else 1}"
        qs.append({"id": qid, "name": str(name) or qid,
                   "max_score": max_score, "aspect": aspect,
                   "region": region})
        save_descriptive_config(str(self._desc_config_path()), self._desc_config)
        self._desc_crops = None   # 領域が変わったので切り出しを無効化
        self._refresh_desc_summary()
        return _ok(state=self.state, question_id=qid)

    def update_descriptive_region(self, qid, region):
        from descriptive_scorer import save_descriptive_config
        q = self._find_question(qid)
        if q is None:
            return _err(f"問題が見つかりません: {qid}")
        try:
            q["region"] = [int(v) for v in region]
            assert len(q["region"]) == 4
        except (TypeError, ValueError, AssertionError):
            return _err("領域は [x1,y1,x2,y2] で指定してください")
        save_descriptive_config(str(self._desc_config_path()), self._desc_config)
        self._desc_crops = None
        self._refresh_desc_summary()
        return _ok(state=self.state)

    def update_descriptive_question(self, qid, name, max_score, aspect):
        """問題名・配点・観点を後から変更する（tk の設定変更ダイアログ相当）。

        配点を下げて既存の得点が超過する場合は新配点にキャップし、
        件数を capped で返す（tk は確認ダイアログ、webui は実施後に通知）。
        """
        from descriptive_scorer import (save_descriptive_config,
                                        save_descriptive_scores)
        q = self._find_question(qid)
        if q is None:
            return _err(f"問題が見つかりません: {qid}")
        try:
            max_score = int(max_score)
            aspect = int(aspect)
            assert max_score > 0
            name = str(name).strip()
            assert name
        except (TypeError, ValueError, AssertionError):
            return _err("問題名は必須、配点は正の整数で指定してください")
        q["name"] = name
        q["aspect"] = aspect
        capped = 0
        if max_score < q["max_score"]:
            for scores in self._desc_scores["scores"].values():
                if scores.get(qid) is not None and scores[qid] > max_score:
                    scores[qid] = max_score
                    capped += 1
            if capped:
                save_descriptive_scores(str(self._desc_scores_path()),
                                        self._desc_scores)
        q["max_score"] = max_score
        save_descriptive_config(str(self._desc_config_path()), self._desc_config)
        self._refresh_desc_summary()
        return _ok(state=self.state, capped=capped)

    def reset_descriptive_scores(self, qid):
        """1問ぶんの採点データをすべて消す（tk 採点リセット相当）"""
        from descriptive_scorer import save_descriptive_scores
        q = self._find_question(qid)
        if q is None:
            return _err(f"問題が見つかりません: {qid}")
        removed = 0
        for scores in self._desc_scores["scores"].values():
            if qid in scores:
                del scores[qid]
                removed += 1
        save_descriptive_scores(str(self._desc_scores_path()), self._desc_scores)
        self._refresh_desc_summary()
        return _ok(state=self.state, removed=removed)

    def delete_descriptive_question(self, qid):
        from descriptive_scorer import (save_descriptive_config,
                                        save_descriptive_scores)
        q = self._find_question(qid)
        if q is None:
            return _err(f"問題が見つかりません: {qid}")
        self._desc_config["questions"].remove(q)
        save_descriptive_config(str(self._desc_config_path()), self._desc_config)
        # スコアも掃除（集計に消した問題が混ざらないように）
        removed = 0
        for f_scores in self._desc_scores["scores"].values():
            if qid in f_scores:
                del f_scores[qid]
                removed += 1
        if removed:
            save_descriptive_scores(str(self._desc_scores_path()),
                                    self._desc_scores)
        self._desc_crops = None
        self._refresh_desc_summary()
        return _ok(state=self.state)

    def _find_question(self, qid):
        if self.state["descriptive"] is None:
            return None
        return next((q for q in self._desc_config["questions"]
                     if q["id"] == qid), None)

    # --- 採点 -------------------------------------------------------

    def start_descriptive_scoring(self):
        """領域切り出しを行い、採点ビューを開ける状態にする"""
        from descriptive_scorer import trim_descriptive_regions
        from constants import RESULTS_FOLDER, BOXED_FOLDER
        if self.state["descriptive"] is None:
            return _err("画像フォルダを選択してください")
        if not self._desc_config["questions"]:
            return _err("記述問題が設定されていません。先に「記述問題設定」を行ってください")
        boxed = Path(self.state["image_folder"]) / RESULTS_FOLDER / BOXED_FOLDER
        if not boxed.exists():
            return _err("補正済み画像がありません。先に認識（または画像準備）を実行してください")
        try:
            self._desc_crops = trim_descriptive_regions(
                str(boxed), self._desc_config,
                original_image_folder=self.state["image_folder"])
        except Exception as e:
            logger.exception("領域切り出しに失敗")
            return _err(f"採点用の切り出しに失敗しました: {e}")
        self._refresh_desc_summary()
        return _ok(state=self.state)

    def list_descriptive_targets(self, qid):
        """指定問題の採点対象一覧（ファイル名と現在の得点）"""
        if self._desc_crops is None:
            return _err("先に start_descriptive_scoring を実行してください")
        crops = self._desc_crops.get(qid)
        if crops is None:
            return _err(f"問題が見つかりません: {qid}")
        scores = self._desc_scores["scores"]
        items = [{"filename": f, "score": scores.get(f, {}).get(qid)}
                 for f in sorted(crops.keys())]
        return _ok(items=items)

    def get_descriptive_crop(self, qid, filename):
        if self._desc_crops is None:
            return _err("先に start_descriptive_scoring を実行してください")
        path = (self._desc_crops.get(qid) or {}).get(filename)
        if not path or not Path(path).exists():
            return _err(f"切り出し画像がありません: {qid}/{filename}")
        b64 = base64.b64encode(Path(path).read_bytes()).decode("ascii")
        return _ok(data_url=f"data:image/png;base64,{b64}")

    def set_descriptive_score(self, filename, qid, score):
        """得点を設定（None で未採点に戻す）。設定のたびに保存（tk 同様アトミック）"""
        from descriptive_scorer import save_descriptive_scores
        q = self._find_question(qid)
        if q is None:
            return _err(f"問題が見つかりません: {qid}")
        if score is not None:
            try:
                score = int(score)
            except (TypeError, ValueError):
                return _err("得点は整数で指定してください")
            if not (0 <= score <= q["max_score"]):
                return _err(f"得点は 0〜{q['max_score']} の範囲で指定してください")
        f_scores = self._desc_scores["scores"].setdefault(str(filename), {})
        if score is None:
            f_scores.pop(qid, None)
        else:
            f_scores[qid] = score
        save_descriptive_scores(str(self._desc_scores_path()), self._desc_scores)
        self._refresh_desc_summary()
        return _ok(state=self.state, filename=filename, qid=qid, score=score)
