"""workspace.py — 試験フォルダ（1回の採点＝1フォルダ）の場所と記録。

試験フォルダの中身:
    <試験フォルダ>/
        *.pdf                     先生が置いた答案のPDF
        scans/                    PDFを画像にしたもの（元画像。LLMは見ない）
          _saiten_grading_results/  採点侍の結果フォルダ（既存の仕組みのまま）
        _llm_view/                LLMに見せてよい黒塗り済みの写し
        project.json              このCLIの記録（向き・黒塗り範囲・承認の日時）
        進捗.md / 採点基準.md      LLMと先生が読む作業メモ
"""
import json
import os
import tempfile
from datetime import datetime
from pathlib import Path

SCANS_DIR = "scans"
LLM_VIEW_DIR = "_llm_view"
PROJECT_FILE = "project.json"
IMAGE_SUFFIXES = (".jpg", ".jpeg", ".png")


def now_iso():
    return datetime.now().astimezone().isoformat(timespec="seconds")


class Project:
    """1つの試験フォルダ。project.json の読み書きを受け持つ"""

    def __init__(self, root):
        self.root = Path(root).resolve()
        if not self.root.is_dir():
            raise FileNotFoundError(f"試験フォルダがありません: {self.root}")
        self.data = self._load()

    # --- 場所 -------------------------------------------------------

    @property
    def scans(self):
        return self.root / SCANS_DIR

    @property
    def llm_view(self):
        return self.root / LLM_VIEW_DIR

    @property
    def manifest_path(self):
        return self.root / PROJECT_FILE

    def pdfs(self):
        return sorted(p for p in self.root.iterdir()
                      if p.is_file() and p.suffix.lower() == ".pdf")

    def scan_files(self):
        """元画像の一覧（ファイル名順）"""
        if not self.scans.is_dir():
            return []
        return sorted(p for p in self.scans.iterdir()
                      if p.is_file() and p.suffix.lower() in IMAGE_SUFFIXES)

    # --- 記録 -------------------------------------------------------

    def _load(self):
        if self.manifest_path.exists():
            return json.loads(self.manifest_path.read_text(encoding="utf-8"))
        return {"version": 1, "created": now_iso()}

    def save(self):
        """途中で落ちても壊れたファイルを残さないよう、書いてから差し替える"""
        self.data["updated"] = now_iso()
        fd, tmp = tempfile.mkstemp(dir=str(self.root), suffix=".tmp")
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, self.manifest_path)
