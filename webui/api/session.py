"""セッション保存/復元 — tk 版の session_state.json（version 1）と互換。

- 保存: tk 版 main_gui._save_session_state と同じキー・同じ相対パス化。
  tk で保存したセッションを webui で読めるし、その逆も読める。
- 復元: tk 版はモード起動後にファイルを選ばせ、形式不一致ならエラーにするが、
  webui はモード選択画面の「採点再開」から入り、モード自体をセッションから
  復元する（不一致が起こり得ない分だけ安全）。
- 壊れたパスは tk 版の修復ダイアログの代わりに warnings で返し、
  ユーザーは通常の選択ボタンで選び直す。
"""
from __future__ import annotations

import logging
from pathlib import Path

logger = logging.getLogger(__name__)

_SESSION_FILE_TYPES = ("セッションファイル (session_state.json)",
                       "JSONファイル (*.json)")


def _ok(**data):
    return {"ok": True, **data}


def _err(message):
    return {"ok": False, "error": message}


class SessionMixin:
    def _session_path(self):
        from constants import SESSION_STATE_FILE
        if not self.state["image_folder"]:
            return None
        return self._results_data_folder() / SESSION_STATE_FILE

    # --- 保存 -------------------------------------------------------

    def save_session(self):
        """現在の状態を tk 互換形式で保存する"""
        import datetime
        from constants import atomic_json_save
        path = self._session_path()
        if path is None:
            return _err("画像フォルダを選択してください")
        img_folder = Path(self.state["image_folder"])
        if not img_folder.exists():
            return _err(f"画像フォルダが見つかりません: {img_folder}")

        def _to_rel(abs_path_str):
            if not abs_path_str:
                return ""
            try:
                return str(Path(abs_path_str).relative_to(img_folder))
            except ValueError:
                return abs_path_str  # 別ドライブ等は絶対パスのまま（tk と同じ）

        data = {
            "version": 1,
            "app_mode": self.state["app_mode"],
            "mark_format": self.state["mark_format"],
            "image_folder": str(img_folder),
            "coord_excel": _to_rel(self.state["coord_file"]),
            "template": _to_rel(self.state["answer_key"]),
            "omr_result": _to_rel(self.state["omr_result"]),
            # tk は StringVar なので文字列で保存される。互換のため揃える
            "skip_questions": str(self.state["skip_questions"]),
            "color_threshold": self.state["color_threshold"],
            "area_threshold": self.state["area_threshold"],
            "descriptive_enabled": self.state["app_mode"] in (
                "mark_and_descriptive", "descriptive_only"),
            "rendering_settings": self.state["rendering_settings"] or {},
            "saved_at": datetime.datetime.now().isoformat(),
            # webui 拡張（tk は無視する）
            "webui": {
                "omr_mode": self.state["omr_mode"],
                "name_trim_enabled": self.state["name_trim_enabled"],
                "name_trim_region": self.state["name_trim_region"],
                "include_descriptive_in_analysis":
                    self.state["include_descriptive_in_analysis"],
            },
        }
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            atomic_json_save(path, data)
        except Exception as e:
            return _err(f"セッション保存に失敗しました: {e}")
        return _ok(path=str(path))

    def _save_session_quietly(self):
        """要所での自動保存（tk 版と同様、失敗は許容してログのみ）"""
        try:
            if self.state["image_folder"]:
                self.save_session()
        except Exception:
            logger.exception("セッション自動保存に失敗")

    # --- 復元 -------------------------------------------------------

    def restore_session(self, path=None):
        """session_state.json を選んで状態を復元する。

        戻り値: {ok, state, warnings: [壊れていて適用しなかった項目の説明]}
        """
        from constants import load_json_safe
        if path is None:
            path = self._win.open_file_dialog(file_types=_SESSION_FILE_TYPES)
            if not path:
                return _ok(state=self.state, cancelled=True)
        data = load_json_safe(Path(path), required_keys=["version"])
        if not data:
            return _err(f"セッションファイルを読み込めません: {path}")

        base = Path(data.get("image_folder", ""))
        if not data.get("image_folder") or not base.exists():
            return _err(
                f"画像フォルダが見つかりません:\n{base}\n"
                "フォルダを移動・削除していないか確認してください")

        res = self.set_mode(data.get("app_mode", "mark_only"),
                            data.get("mark_format", "standard"))
        if not res["ok"]:
            return res

        # 画像フォルダ（select_image_folder と同じ検証）
        from api.bridge import _IMAGE_SUFFIXES
        count = sum(1 for f in base.iterdir()
                    if f.is_file() and f.suffix.lower() in _IMAGE_SUFFIXES)
        if count == 0:
            return _err(f"選択したフォルダに画像（jpg/png）がありません: {base}")
        self.state["image_folder"] = str(base)
        self.state["image_count"] = count
        self._load_descriptive_state()
        self._load_total_display_state()
        self._load_handwriting_state()

        # 数値の復元（tk は skip を文字列で持つ）
        try:
            self.state["skip_questions"] = int(str(data.get("skip_questions", 4)))
        except ValueError:
            pass
        for key in ("color_threshold", "area_threshold"):
            try:
                self.state[key] = float(data.get(key, self.state[key]))
            except (TypeError, ValueError):
                pass
        webui_ext = data.get("webui") or {}
        omr_mode = webui_ext.get("omr_mode")
        if omr_mode in ("kmeans", "threshold"):
            self.state["omr_mode"] = omr_mode
        if isinstance(webui_ext.get("name_trim_enabled"), bool):
            self.state["name_trim_enabled"] = webui_ext["name_trim_enabled"]
        if isinstance(webui_ext.get("include_descriptive_in_analysis"), bool):
            self.state["include_descriptive_in_analysis"] = \
                webui_ext["include_descriptive_in_analysis"]
        region = webui_ext.get("name_trim_region")
        if (isinstance(region, list) and len(region) == 4 and
                all(isinstance(v, int) for v in region)):
            self.state["name_trim_region"] = region

        # tk セッションの描画詳細設定はそのまま引き継いで採点描画に渡す
        rs = data.get("rendering_settings")
        if isinstance(rs, dict) and rs:
            self.state["rendering_settings"] = rs

        # パスの復元（解決できないものは warnings に落として続行）
        warnings = []

        def _resolve(rel_or_abs):
            if not rel_or_abs:
                return None
            candidate = base / rel_or_abs
            if candidate.exists():
                return candidate
            p = Path(rel_or_abs)
            if p.is_absolute() and p.exists():
                return p
            return False   # 記録はあるが見つからない

        for key, state_key, label in (
                ("coord_excel", "coord_file", "座標ファイル"),
                ("template", "answer_key", "正答データ"),
                ("omr_result", "omr_result", "OMR読取結果")):
            resolved = _resolve(data.get(key, ""))
            if resolved is False:
                warnings.append(
                    f"{label}が見つかりません（{data.get(key)}）。選び直してください")
                continue
            if resolved is not None:
                self.state[state_key] = str(resolved)

        if self.state["coord_file"]:
            self._refresh_coord_summary()
        if self.state["answer_key"]:
            self._refresh_key_summary()
        return _ok(state=self.state, warnings=warnings)
