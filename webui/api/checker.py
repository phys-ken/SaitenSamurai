"""checker.py — マークチェック（読取結果の確認・訂正）の API 層。

tk 版 MarkCheckerGUI の中核ワークフローを移植する:
  全エントリ読込 → カテゴリ絞り込み → 切り出し画像で目視 → 訂正入力
  → 訂正CSVへ自動保存 → 「xlsxに反映」（バックアップ付き）

分類・切り出し・xlsx書き戻しはすべて main_src/mark_checker.py の
既存関数を使う（ロジック再実装なし）。tk 版との対応で重要な点:
- 訂正CSVの場所と名前は tk と同一（xlsx と同じフォルダの
  tmp_checking_dm_nm.csv）→ tk 版で途中まで訂正したデータを
  webui で開いても続きから作業できる（後方互換）
- coordinates.csv 参照時の問題番号は question_no + skip_questions
  （DEVELOPMENT.md「問題番号のオフセット」）
"""
import base64
import io
import logging
from pathlib import Path

logger = logging.getLogger(__name__)


def _ok(**data):
    return {"ok": True, **data}


def _err(message):
    return {"ok": False, "error": str(message)}


# 表示カテゴリの並び（tk 版のタブ相当。値カテゴリは動的に続く）
_ERROR_CATEGORIES = ["ノーマーク", "複数マーク", "不正な値", "無効回答(-1)"]


class CheckerMixin:

    def _init_checker(self):
        self._checker = None  # {entries: [dict], coords_df, image_cache}
        self.state["checker"] = None

    # ----------------------------------------------------------------
    def open_mark_checker(self):
        if self.state["job"]["running"]:
            return _err("別の処理が実行中です。完了または中断を待ってください")
        from constants import RESULTS_FOLDER, RESULTS_DATA_FOLDER
        from mark_checker import (
            detect_all_entries_checker, load_errors_checker,
            load_coordinates_csv_checker, CorrectedImageCache,
        )

        if not self.state["image_folder"]:
            return _err("答案画像フォルダを選択してください")
        if not self.state["omr_result"]:
            return _err("読み取り結果がありません。先に「答案を読み取る」を実行してください")
        coords_csv = (Path(self.state["image_folder"]) / RESULTS_FOLDER /
                      RESULTS_DATA_FOLDER / "coordinates.csv")
        if not coords_csv.exists():
            return _err("coordinates.csv が見つかりません。先に「答案を読み取る」を実行してください")

        registered = None
        if self.state["answer_key"]:
            try:
                from scoring_engine import load_template
                registered = set(load_template(
                    self.state["answer_key"],
                    mark_format=self.state["mark_format"]).keys())
            except Exception as e:
                logger.warning("正答データの読込に失敗（全問チェックに切替）: %s", e)

        try:
            df = detect_all_entries_checker(
                self.state["omr_result"], registered_questions=registered,
                mark_format=self.state["mark_format"])
        except Exception as e:
            logger.exception("エントリ読込失敗")
            return _err(f"読取結果を読み込めませんでした: {e}")

        entries = df.to_dict("records")
        # 白さキャッシュ（Step1 保存）を各エントリへ付与 — tk の「白さ順」と同じ根拠
        from constants import WHITENESS_CACHE_FILE
        whiteness_map = {}
        wpath = coords_csv.parent / WHITENESS_CACHE_FILE
        if wpath.exists():
            try:
                import json
                whiteness_map = json.loads(wpath.read_text(encoding="utf-8"))
            except Exception as e:
                logger.warning("白さキャッシュ読込エラー: %s", e)
        skip = self.state["skip_questions"]
        for i, e in enumerate(entries):
            e["id"] = i
            e["after"] = "" if not isinstance(e.get("after"), str) else e["after"]
            by_file = whiteness_map.get(e["filename"], {}) or {}
            q = int(e["question_no"])
            # 互換: skip 込み/抜きの両キーを参照（tk と同じ）
            e["whiteness"] = float(by_file.get(str(q + skip), by_file.get(str(q), 0.0)))

        # 保存済み訂正のマージ（tk 版 _merge_corrections と同じキー）
        csv_path = Path(self.state["omr_result"]).parent / "tmp_checking_dm_nm.csv"
        if csv_path.exists():
            saved = load_errors_checker(csv_path)
            merged = 0
            by_key = {(e["filename"], int(e["question_no"])): e for e in entries}
            import pandas as pd
            for _, row in saved.iterrows():
                after = row.get("after")
                # CSV 読み戻しで数字の訂正は int64 になる（pandas の型推論）。
                # 文字列へ正規化しないと isinstance(str) 判定で訂正が消える
                if pd.isna(after):
                    continue
                s_val = str(after).strip()
                if s_val == "":
                    continue
                try:
                    f = float(s_val)
                    if f.is_integer():
                        s_val = str(int(f))
                except ValueError:
                    pass
                target = by_key.get((row["filename"], int(row["question_no"])))
                if target is not None:
                    target["after"] = s_val
                    merged += 1
            if merged:
                logger.info("保存済みの訂正 %d 件を読み込みました", merged)

        self._checker = {
            "entries": entries,
            "coords_df": load_coordinates_csv_checker(str(coords_csv)),
            "csv_path": str(csv_path),
            "cache": CorrectedImageCache(max_size=8),
        }
        self.state["checker"] = self._checker_summary()
        return _ok(state=self.state)

    def _checker_summary(self):
        """カテゴリ別件数と訂正数（UIのタブ表示用）"""
        entries = self._checker["entries"]
        counts = {}
        for e in entries:
            counts[e["category"]] = counts.get(e["category"], 0) + 1
        value_cats = sorted([c for c in counts if c not in _ERROR_CATEGORIES])
        categories = ([{"name": c, "count": counts[c], "is_error": True}
                       for c in _ERROR_CATEGORIES if c in counts] +
                      [{"name": c, "count": counts[c], "is_error": False}
                       for c in value_cats])
        return {
            "open": True,
            "total": len(entries),
            "corrected": sum(1 for e in entries if e["after"]),
            "error_count": sum(1 for e in entries
                               if e["category"] in _ERROR_CATEGORIES),
            "categories": categories,
        }

    def batch_correct_no_mark(self):
        """未訂正のノーマーク全件を無効回答(-1)に設定する（tk の一括ボタン相当）。

        1件ずつ set_correction を呼ぶと数百回のCSV保存になるため、
        まとめて書き換えて保存は1回にする。
        """
        if not self._checker:
            return _err("マークチェックが開かれていません")
        targets = [e for e in self._checker["entries"]
                   if e["category"] == "ノーマーク" and not e["after"]]
        for e in targets:
            e["after"] = "-1"
        self._save_corrections_csv()
        self.state["checker"] = self._checker_summary()
        return _ok(state=self.state, applied=len(targets))

    # ----------------------------------------------------------------
    def get_checker_entries(self, category=None, page=0, page_size=24):
        if not self._checker:
            return _err("マークチェックが開かれていません")
        entries = self._checker["entries"]
        if category == "__errors__":
            filtered = [e for e in entries if e["category"] in _ERROR_CATEGORIES]
        elif category:
            filtered = [e for e in entries if e["category"] == category]
        else:
            filtered = entries
        page, page_size = int(page), int(page_size)
        start = page * page_size
        items = [{k: e[k] for k in
                  ("id", "filename", "question_no", "before", "after",
                   "category", "error_type")}
                 for e in filtered[start:start + page_size]]
        return _ok(items=items, total=len(filtered), page=page,
                   page_size=page_size)

    def get_entry_image(self, entry_id):
        """エントリのマーク行を切り出して data URL で返す"""
        from mark_checker import get_display_image_checker
        if not self._checker:
            return _err("マークチェックが開かれていません")
        try:
            e = self._checker["entries"][int(entry_id)]
        except (IndexError, ValueError, TypeError):
            return _err(f"不正なエントリIDです: {entry_id}")

        pil = get_display_image_checker(
            self._checker["coords_df"],
            self.state["image_folder"],
            e["filename"],
            int(e["question_no"]) + self.state["skip_questions"],  # 座標は元番号
            cache=self._checker["cache"],
        )
        if pil is None:
            return _err("画像を切り出せませんでした（座標または画像がありません）")
        buf = io.BytesIO()
        pil.save(buf, format="PNG")
        b64 = base64.b64encode(buf.getvalue()).decode("ascii")
        return _ok(data_url=f"data:image/png;base64,{b64}")

    # ----------------------------------------------------------------
    def set_correction(self, entry_id, value):
        """訂正値を設定（空文字で取消）。妥当性検証は tk 版と同じ規則。

        設定のたびに訂正CSVへ保存する（tk と同じファイルなので相互運用可能）。
        """
        from constants import (MARK_FORMAT_MULTI_DIGIT,
                               MULTI_DIGIT_VALID_SYMBOLS)
        if not self._checker:
            return _err("マークチェックが開かれていません")
        try:
            e = self._checker["entries"][int(entry_id)]
        except (IndexError, ValueError, TypeError):
            return _err(f"不正なエントリIDです: {entry_id}")

        v = str(value).strip()
        if v == "" or v == "-1":
            pass  # 取消 / 無効回答マーカー
        elif self.state["mark_format"] == MARK_FORMAT_MULTI_DIGIT:
            v = v.lower()
            if v not in MULTI_DIGIT_VALID_SYMBOLS:
                return _err("マーク記号1文字（- 0〜9 a〜d)または -1 を入力してください")
        else:
            if not (v.isdigit() and len(v) == 1):
                return _err("選択肢番号1桁（0〜9）または -1 を入力してください")

        e["after"] = v
        self._save_corrections_csv()
        self.state["checker"] = self._checker_summary()
        return _ok(state=self.state, entry={k: e[k] for k in
                   ("id", "before", "after", "category")})

    def _save_corrections_csv(self):
        """訂正の入った行だけを tk 互換のCSVへアトミック保存する。

        1打鍵ごとに全件を書き直すため、直書きだと書き込み途中の異常終了で
        目視訂正が丸ごと失われる（A5）。同一フォルダの一時ファイルへ書いて
        os.replace で差し替える（atomic_json_save と同じ手順のCSV版）
        """
        import os
        import tempfile
        import pandas as pd
        rows = [{"filename": e["filename"], "question_no": e["question_no"],
                 "before": e["before"], "after": e["after"],
                 "error_type": e["error_type"]}
                for e in self._checker["entries"] if e["after"]]
        df = pd.DataFrame(rows, columns=["filename", "question_no", "before",
                                         "after", "error_type"])
        target = Path(self._checker["csv_path"])
        fd, tmp = tempfile.mkstemp(prefix=target.stem + "_",
                                   suffix=".tmp", dir=str(target.parent))
        try:
            with os.fdopen(fd, "w", encoding="utf-8-sig", newline="") as f:
                df.to_csv(f, index=False)
            os.replace(tmp, str(target))
        except Exception:
            try:
                os.unlink(tmp)
            except OSError:
                pass
            raise

    def apply_corrections(self):
        """訂正を xlsx に反映（バックアップ作成込み）→ エントリ再読込"""
        if self.state["job"]["running"]:
            return _err("採点・集計の実行中は反映できません。完了を待ってください")
        from mark_checker import apply_corrections_checker
        if not self._checker:
            return _err("マークチェックが開かれていません")
        corrected = [e for e in self._checker["entries"] if e["after"]]
        if not corrected:
            return _err("反映する訂正がありません")
        try:
            backup, count = apply_corrections_checker(
                self.state["omr_result"], self._checker["csv_path"])
        except Exception as e:
            logger.exception("xlsx反映に失敗")
            return _err(f"xlsxへの反映に失敗しました: {e}")
        # 反映後は最新の xlsx で開き直す（before が更新された状態になる）
        result = self.open_mark_checker()
        if not result["ok"]:
            return result
        return _ok(state=self.state, applied=count, backup=str(backup))

    def close_mark_checker(self):
        self._checker = None
        self.state["checker"] = None
        return _ok(state=self.state)
