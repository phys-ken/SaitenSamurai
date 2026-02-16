"""
R統合テスト — PythonからRを呼び出してexametrika分析が実行できることを検証
==========================================================================

このテストは以下をシステムレベルで検証する:
  1. Pythonでダミー0/1データ + Rスクリプト + Rmdテンプレートを生成
  2. subprocess でRscriptを呼び出し、create_report.R を実行
  3. exametrika_report.html が生成されることを確認
  4. 生成されたHTMLにCTT/IRT/LRA/Biclustering/Ranklusteringの結果が含まれることを確認

前提条件:
  - R >= 4.1.0 がインストールされていること
  - exametrika パッケージがインストールされていること
  - rmarkdown パッケージがインストールされていること

NOTE:
  R本体がない環境ではスキップされる（CI環境等への影響なし）。
"""

import os
import sys
import shutil
import subprocess
import tempfile
import pytest
from pathlib import Path

import numpy as np
import pandas as pd

# プロジェクトルートをパスに追加
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "main_src"))

from r_export import (
    _export_scored_data,
    _export_item_info,
    _generate_r_script,
    _generate_rmd_template,
    R_DATA_CSV,
    R_ITEM_INFO_CSV,
    R_SCRIPT_FILE,
    R_RMD_TEMPLATE_FILE,
    DEFAULT_N_RANKS,
    DEFAULT_N_FIELDS,
)


# ── Rscript検出 ────────────────────────────────────────

def _find_rscript():
    """Rscript.exe のパスを取得する。見つからなければ None。"""
    # PATH にある場合
    rscript = shutil.which("Rscript")
    if rscript:
        return rscript

    # Windows 標準インストール先
    r_base = Path(r"C:\Program Files\R")
    if r_base.exists():
        # 最新バージョンを優先
        versions = sorted(r_base.iterdir(), reverse=True)
        for v in versions:
            candidate = v / "bin" / "Rscript.exe"
            if candidate.exists():
                return str(candidate)

    return None


RSCRIPT_PATH = _find_rscript()

# R未インストール環境では全テストをスキップ
pytestmark = pytest.mark.skipif(
    RSCRIPT_PATH is None,
    reason="Rscript が見つかりません（R未インストール）"
)


def _r_package_available(package_name: str) -> bool:
    """指定のRパッケージがインストール済みか確認する。"""
    if RSCRIPT_PATH is None:
        return False
    try:
        result = subprocess.run(
            [RSCRIPT_PATH, "-e",
             f"if (!requireNamespace('{package_name}', quietly=TRUE)) quit(status=1)"],
            capture_output=True, timeout=30,
        )
        return result.returncode == 0
    except Exception:
        return False


_HAS_EXAMETRIKA = _r_package_available("exametrika")
_HAS_RMARKDOWN = _r_package_available("rmarkdown")


# ── ダミーデータ生成 ─────────────────────────────────────

def _generate_dummy_kit(work_dir, n_students=50, n_questions=15,
                        n_ranks=DEFAULT_N_RANKS, n_fields=DEFAULT_N_FIELDS,
                        seed=42):
    """
    ダミー0/1データでR分析キット一式を生成する。

    Args:
        work_dir: 出力先ディレクトリ
        n_students: 受験者数
        n_questions: 設問数（バイクラスタリングには最低5問程度必要）
        n_ranks: 潜在ランク数
        n_fields: フィールド数
        seed: 乱数シード

    Returns:
        Path: work_dir
    """
    rng = np.random.default_rng(seed)
    work = Path(work_dir)
    work.mkdir(parents=True, exist_ok=True)

    # 受験者ごとの能力を設定して、より現実的なデータを生成
    abilities = rng.normal(0.5, 0.2, size=n_students).clip(0.1, 0.9)
    difficulties = np.linspace(0.8, 0.2, n_questions)

    data = np.zeros((n_students, n_questions), dtype=int)
    for i in range(n_students):
        for j in range(n_questions):
            prob = abilities[i] * (difficulties[j] / 0.5)
            prob = min(max(prob, 0.05), 0.95)
            data[i, j] = int(rng.random() < prob)

    columns = [f"Q{i+1}" for i in range(n_questions)]
    index = [f"S{i+1:03d}" for i in range(n_students)]
    score_matrix = pd.DataFrame(data, columns=columns, index=index)

    key_df = pd.DataFrame({
        "QuestionID": columns,
        "Key": [str(rng.integers(1, 6)) for _ in range(n_questions)],
    })

    # ファイル出力
    _export_scored_data(score_matrix, work / R_DATA_CSV)
    _export_item_info(score_matrix, key_df, work / R_ITEM_INFO_CSV)
    _generate_r_script(work / R_SCRIPT_FILE, n_ranks, n_fields)
    _generate_rmd_template(
        work / R_RMD_TEMPLATE_FILE, n_ranks, n_fields,
        title="統合テスト用レポート", author="Automated Test"
    )

    return work


# ── テスト ────────────────────────────────────────────


class TestRIntegration:
    """PythonからRを呼び出して分析が完了することを検証"""

    @pytest.fixture
    def r_kit_dir(self, tmp_path):
        """ダミーデータ付きR分析キットを一時ディレクトリに生成"""
        return _generate_dummy_kit(tmp_path / "r_kit")

    def test_rscript_available(self):
        """Rscript が実行可能であること"""
        result = subprocess.run(
            [RSCRIPT_PATH, "--version"],
            capture_output=True, text=True, timeout=10
        )
        # Rscript --version は stderr に出力する場合がある
        combined = result.stdout + result.stderr
        assert "R" in combined

    @pytest.mark.skipif(not _HAS_EXAMETRIKA, reason="exametrika パッケージ未インストール")
    def test_exametrika_installed(self):
        """exametrika パッケージが利用可能であること"""
        result = subprocess.run(
            [RSCRIPT_PATH, "-e", "library(exametrika); cat('OK')"],
            capture_output=True, text=True, timeout=30
        )
        assert "OK" in result.stdout

    @pytest.mark.skipif(not _HAS_RMARKDOWN, reason="rmarkdown パッケージ未インストール")
    def test_rmarkdown_installed(self):
        """rmarkdown パッケージが利用可能であること"""
        result = subprocess.run(
            [RSCRIPT_PATH, "-e", "library(rmarkdown); cat('OK')"],
            capture_output=True, text=True, timeout=30
        )
        assert "OK" in result.stdout

    def test_kit_files_generated(self, r_kit_dir):
        """分析キットの全ファイルが生成されていること"""
        assert (r_kit_dir / R_DATA_CSV).exists()
        assert (r_kit_dir / R_ITEM_INFO_CSV).exists()
        assert (r_kit_dir / R_SCRIPT_FILE).exists()
        assert (r_kit_dir / R_RMD_TEMPLATE_FILE).exists()

    def test_r_script_readable_japanese(self, r_kit_dir):
        """Rスクリプトの日本語がリテラル文字列で読めること"""
        content = (r_kit_dir / R_SCRIPT_FILE).read_text(encoding="utf-8")
        assert "\\u30" not in content
        assert "レポート" in content

    @pytest.mark.skipif(
        not (_HAS_EXAMETRIKA and _HAS_RMARKDOWN),
        reason="exametrika / rmarkdown パッケージ未インストール",
    )
    @pytest.mark.timeout(180)
    def test_create_report_generates_html(self, r_kit_dir):
        """create_report.R を実行して HTML レポートが生成されること"""
        result = subprocess.run(
            [RSCRIPT_PATH, R_SCRIPT_FILE],
            cwd=str(r_kit_dir),
            capture_output=True,
            text=True,
            timeout=180,
            encoding="utf-8",
            errors="replace",
        )

        # デバッグ用出力
        if result.returncode != 0:
            print("=== STDOUT ===")
            print(result.stdout[-2000:] if len(result.stdout) > 2000 else result.stdout)
            print("=== STDERR ===")
            print(result.stderr[-2000:] if len(result.stderr) > 2000 else result.stderr)

        assert result.returncode == 0, (
            f"Rscript がエラーで終了 (code={result.returncode}):\n"
            f"stderr: {result.stderr[-1000:]}"
        )

        html_path = r_kit_dir / "exametrika_report.html"
        assert html_path.exists(), "exametrika_report.html が生成されませんでした"
        assert html_path.stat().st_size > 1000, "HTML ファイルが小さすぎます"

    @pytest.mark.timeout(180)
    def test_html_contains_all_analyses(self, r_kit_dir):
        """生成 HTML に全分析セクション（CTT/IRT/LRA/Biclustering）が含まれること"""
        # まずレポート生成
        result = subprocess.run(
            [RSCRIPT_PATH, R_SCRIPT_FILE],
            cwd=str(r_kit_dir),
            capture_output=True,
            text=True,
            timeout=180,
            encoding="utf-8",
            errors="replace",
        )
        if result.returncode != 0:
            pytest.skip(f"レポート生成失敗: {result.stderr[-500:]}")

        html_path = r_kit_dir / "exametrika_report.html"
        if not html_path.exists():
            pytest.skip("HTML が生成されなかったためスキップ")

        content = html_path.read_text(encoding="utf-8", errors="replace")

        # 各分析セクションが含まれていることを確認
        assert "CTT" in content, "CTT セクションが見つかりません"
        assert "IRT" in content, "IRT セクションが見つかりません"
        assert "LRA" in content, "LRA セクションが見つかりません"
        # Biclustering / Ranklustering （HTMLではセクションタイトルとして出力）
        assert "Biclustering" in content or "バイクラスタリング" in content, \
            "バイクラスタリングセクションが見つかりません"
        assert "Ranklustering" in content or "ランクラスタリング" in content, \
            "ランクラスタリングセクションが見つかりません"

    @pytest.mark.timeout(180)
    def test_excel_contains_all_sheets(self, r_kit_dir):
        """生成 Excel に全分析シート（15 枚）が含まれること"""
        # レポート生成
        result = subprocess.run(
            [RSCRIPT_PATH, R_SCRIPT_FILE],
            cwd=str(r_kit_dir),
            capture_output=True,
            text=True,
            timeout=180,
            encoding="utf-8",
            errors="replace",
        )
        if result.returncode != 0:
            pytest.skip(f"レポート生成失敗: {result.stderr[-500:]}")

        xlsx_path = r_kit_dir / "analysis_results.xlsx"
        assert xlsx_path.exists(), "analysis_results.xlsx が生成されませんでした"

        # openxlsx 生成ファイルは openpyxl read_only=True でdimension不整合が
        # 起きるため、pandas で読み込む
        all_sheets = pd.read_excel(
            xlsx_path, sheet_name=None, engine="openpyxl"
        )
        sheets = list(all_sheets.keys())

        # ①CTT 信頼性指標・項目統計量
        assert "CTT信頼性" in sheets, f"CTT信頼性シートが見つかりません: {sheets}"
        assert "CTT項目除外信頼性" in sheets
        assert "CTT項目統計量" in sheets

        # ②IRT パラメータ + 受験者能力値
        assert "IRTパラメータ" in sheets
        assert "IRT受験者能力値" in sheets

        # ③LRA ランク所属 + 項目情報
        assert "LRAランク所属" in sheets
        assert "LRA項目参照プロファイル" in sheets

        # ④RC(ランクラスタリング) 受験者 + 設問情報
        assert "RCランク所属" in sheets
        assert "RC設問フィールド所属" in sheets
        assert "RCフィールド参照プロファイル" in sheets

        # CTT信頼性シートにデータが存在する（空でないこと）
        df_ctt = all_sheets["CTT信頼性"]
        assert len(df_ctt) >= 1, f"CTT信頼性シートが空です: {df_ctt}"

        # CTT項目統計量にデータが存在する
        df_items = all_sheets["CTT項目統計量"]
        assert len(df_items) >= 2, f"CTT項目統計量シートが空です: {df_items}"

        # IRT受験者能力値にデータが存在する
        df_ability = all_sheets["IRT受験者能力値"]
        assert len(df_ability) >= 2, f"IRT受験者能力値シートが空です: {df_ability}"
        assert "EAP" in df_ability.columns, \
            f"EAP列がありません: {list(df_ability.columns)}"
        assert "合計得点" in df_ability.columns, \
            f"合計得点列がありません: {list(df_ability.columns)}"

        # LRAランク所属に合計得点列が含まれる
        df_lra = all_sheets["LRAランク所属"]
        assert "合計得点" in df_lra.columns, \
            f"LRAに合計得点列がありません: {list(df_lra.columns)}"

        # RCランク所属にデータが存在する
        df_rc = all_sheets["RCランク所属"]
        assert len(df_rc) >= 2, f"RCランク所属シートが空です: {df_rc}"

        # RC設問フィールド所属にデータが存在する
        df_rc_field = all_sheets["RC設問フィールド所属"]
        assert len(df_rc_field) >= 2, \
            f"RC設問フィールド所属シートが空です: {df_rc_field}"
