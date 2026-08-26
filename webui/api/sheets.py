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
