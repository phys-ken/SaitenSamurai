"""bridge.py — JS から呼ばれる API 層。

設計原則（webui/docs/plan.md）:
- このモジュールは **pywebview を import しない**。ネイティブダイアログ等の
  ウィンドウ依存機能は app.py から注入される adapter 経由で使う。
  → CI や pywebview の無い環境でもそのまま pytest できる
- 各メソッドは JSON 化可能な dict を返す。成功は {"ok": True, ...}、
  失敗は {"ok": False, "error": "利用者向けの日本語メッセージ"}。
  例外を JS 側に漏らさない（漏らすと画面ごとに握り方がバラつく）
- 採点ロジックは main_src のドライバを呼ぶだけ。ここにロジックを書かない
"""
import logging
import sys
from pathlib import Path

# main_src をインポートパスへ（リポジトリ同居構成）
_REPO_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_REPO_ROOT / "main_src") not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT / "main_src"))

from constants import APP_VERSION  # noqa: E402

from api.jobs import JobsMixin  # noqa: E402
from api.checker import CheckerMixin  # noqa: E402
from api.sheets import SheetsMixin  # noqa: E402
from api.descriptive import DescriptiveMixin  # noqa: E402
from api.session import SessionMixin  # noqa: E402

logger = logging.getLogger(__name__)

WEBUI_VERSION = "0.1.0"

_IMAGE_SUFFIXES = ('.jpg', '.jpeg', '.png')
_XLSX_FILE_TYPES = ("Excelファイル (*.xlsx;*.xls)",)


def _ok(**data):
    return {"ok": True, **data}


def _err(message):
    return {"ok": False, "error": str(message)}


class DataSourceMixin:
    """データソース選択（画像フォルダ・座標ファイル・正答データ）。

    tk 版と同じ思想を引き継ぐ:
    - 選んだ瞬間に検証する（座標ファイルは行数・マーク数を要約、
      正答データは answer_key_checker を自動実行）。
      「認識実行を押して初めて間違いに気づく」を作らない
    - 検証結果は state に載せ、UI は state の再描画だけを行う
    """

    def _init_state(self):
        self.state = {
            "app_mode": "mark_only",
            "mark_format": "standard",       # standard / multi_digit
            "skip_questions": 4,
            "image_folder": None,
            "image_count": 0,
            "coord_file": None,
            "coord_summary": None,           # {answer_rows, marks_per_row, warning}
            "answer_key": None,
            "key_summary": None,             # {ok, errors, warnings, stats_line, ...}
            "name_trim_enabled": True,       # 氏名画像を集計シートに表示（tk 既定 ON）
            "name_trim_region": None,        # [x1,y1,x2,y2] 00_Processing 座標系
            "rendering_settings": None,      # 描画詳細設定の差分（None=既定のまま）
            "total_display_region": None,    # 合計点表示位置 [x1,y1,x2,y2]
        }

    def get_state(self):
        return _ok(state=self.state)

    def set_mode(self, app_mode, mark_format):
        from constants import (MODE_MARK_ONLY, MODE_MARK_AND_DESCRIPTIVE,
                               MODE_DESCRIPTIVE_ONLY,
                               MARK_FORMAT_STANDARD, MARK_FORMAT_MULTI_DIGIT)
        if app_mode not in (MODE_MARK_ONLY, MODE_MARK_AND_DESCRIPTIVE,
                            MODE_DESCRIPTIVE_ONLY):
            return _err(f"不明なモードです: {app_mode}")
        if mark_format not in (MARK_FORMAT_STANDARD, MARK_FORMAT_MULTI_DIGIT):
            return _err(f"不明なマーク形式です: {mark_format}")
        self.state["app_mode"] = app_mode
        self.state["mark_format"] = mark_format
        # モードが変わると座標整合・正答チェックの判定も変わるため再検証
        if self.state["coord_file"]:
            self._refresh_coord_summary()
        if self.state["answer_key"]:
            self._refresh_key_summary()
        return _ok(state=self.state)

    def set_skip_questions(self, n):
        try:
            n = int(n)
            if n < 0:
                raise ValueError
        except (TypeError, ValueError):
            return _err("Skip（ID欄の数）は 0 以上の整数で指定してください")
        self.state["skip_questions"] = n
        if self.state["coord_file"]:
            self._refresh_coord_summary()
        return _ok(state=self.state)

    # --- 選択 -------------------------------------------------------

    def select_image_folder(self):
        folder = self._win.open_folder_dialog()
        if not folder:
            return _ok(state=self.state, cancelled=True)
        folder_path = Path(folder)
        count = sum(1 for f in folder_path.iterdir()
                    if f.is_file() and f.suffix.lower() in _IMAGE_SUFFIXES)
        if count == 0:
            return _err(
                f"選択したフォルダに画像（jpg/png）がありません: {folder}\n"
                "スキャンした答案画像の入ったフォルダを選んでください")
        self.state["image_folder"] = str(folder_path)
        self.state["image_count"] = count
        self._load_descriptive_state()
        self._load_total_display_state()
        self._save_session_quietly()
        return _ok(state=self.state)

    def select_coord_file(self):
        path = self._win.open_file_dialog(file_types=_XLSX_FILE_TYPES)
        if not path:
            return _ok(state=self.state, cancelled=True)
        self.state["coord_file"] = str(path)
        result = self._refresh_coord_summary()
        if not result["ok"]:
            # 読めないファイルは選択自体を取り消す（tk版と同じ「その場で気づく」）
            self.state["coord_file"] = None
            self.state["coord_summary"] = None
            return result
        self._save_session_quietly()
        return _ok(state=self.state)

    def _refresh_coord_summary(self):
        """座標ファイルの要約と、モード整合の警告を state に反映する。

        tk 版 check_coord_file_gui と同じ判定
        （tk 側は凍結ファイル内のため共有できず、意図的に再実装）。
        """
        from collections import Counter
        from omr_engine import parse_excel_coordinates
        try:
            coords, _ = parse_excel_coordinates(self.state["coord_file"])
        except Exception as e:
            logger.warning("座標ファイルの読み込みに失敗: %s", e)
            return _err(
                "座標ファイルを読み込めませんでした。\n"
                "Mark2 で書き出した座標 Excel を選んでいるか確認してください"
                f"（正答データや読取結果を間違えて選んでいませんか？）\n詳細: {e}")

        skip = self.state["skip_questions"]
        per_q = {}
        for c in coords:
            per_q[c['question_no']] = per_q.get(c['question_no'], 0) + 1
        answer_counts = [n for q, n in per_q.items()
                         if isinstance(q, (int, float)) and q > skip]
        if not answer_counts:
            self.state["coord_summary"] = {
                "answer_rows": 0, "marks_per_row": 0,
                "warning": "採点対象の設問が見つかりません（Skip 数が大きすぎませんか？）",
            }
            return _ok()

        typical = Counter(answer_counts).most_common(1)[0][0]
        warning = None
        is_md = self.state["mark_format"] == "multi_digit"
        if is_md and typical <= 10:
            warning = ("数学マーク採点（複数桁）モードですが、標準テンプレート相当の"
                       "座標ファイルです。座標ファイルかモードが違っていませんか？")
        elif (not is_md) and typical >= 11:
            warning = ("標準マーク採点モードですが、複数桁テンプレート相当の"
                       "座標ファイルです。数学マーク採点モードではありませんか？")
        self.state["coord_summary"] = {
            "answer_rows": len(answer_counts),
            "marks_per_row": typical,
            "warning": warning,
        }
        return _ok()

    def select_answer_key(self):
        path = self._win.open_file_dialog(file_types=_XLSX_FILE_TYPES)
        if not path:
            return _ok(state=self.state, cancelled=True)
        self.state["answer_key"] = str(path)
        self._refresh_key_summary()
        self._save_session_quietly()
        return _ok(state=self.state)

    def _refresh_key_summary(self):
        """正答データの自動チェック（tk 版と同じく選択直後に実行）"""
        from answer_key_checker import run_answer_key_check
        try:
            res, check_md, model_md = run_answer_key_check(
                self.state["answer_key"],
                mark_format=self.state["mark_format"],
                coord_excel_path=self.state["coord_file"],
                skip_questions=self.state["skip_questions"],
            )
        except Exception as e:
            logger.warning("正答チェックに失敗: %s", e)
            self.state["key_summary"] = {
                "ok": False, "errors": [f"正答データを読み込めませんでした: {e}"],
                "warnings": [], "stats_line": None,
                "check_md": None, "model_md": None,
            }
            return

        stats = res.get("stats") or {}
        stats_line = None
        if stats:
            stats_line = (f"{stats.get('問題数', '?')}問 / 満点{stats.get('満点', '?')}点"
                          f" / 使用{stats.get('使用マーク行数', '?')}行")
        self.state["key_summary"] = {
            "ok": res["ok"],
            "errors": res.get("errors", []),
            "warnings": res.get("warnings", []),
            "stats_line": stats_line,
            "check_md": str(check_md) if check_md else None,
            "model_md": str(model_md) if model_md else None,
        }


class Bridge(DataSourceMixin, JobsMixin, CheckerMixin, SheetsMixin,
             DescriptiveMixin, SessionMixin):
    """pywebview の js_api として渡すクラス。

    window_adapter: ネイティブ機能の注入点。必要なメソッド:
        - open_file_dialog(file_types, directory) -> str | None
        - open_folder_dialog(directory) -> str | None
      app.py が pywebview 実装を、テストが偽物を渡す。
    """

    def __init__(self, window_adapter=None):
        self._win = window_adapter
        self._init_state()
        self._init_jobs()
        self._init_checker()
        self._init_descriptive()

    # --- 疎通・情報 -------------------------------------------------

    def ping(self):
        """疎通確認。JS 側の起動シーケンスがブリッジ準備完了を検知するのに使う"""
        return _ok(pong=True)

    def get_app_info(self):
        return _ok(
            app_version=APP_VERSION,
            webui_version=WEBUI_VERSION,
            python=sys.version.split()[0],
        )
