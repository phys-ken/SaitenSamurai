"""LLM協働採点（saiten_llm）のテスト共通設定。

試験データは毎回 make_sample_exam で合成する（実在の生徒は含まない）。
"""
import json
import shutil
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent.parent
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

import saiten_llm  # noqa: E402,F401  main_src / webui をパスに入れる
from saiten_llm.tools.make_sample_exam import make_exam  # noqa: E402


@pytest.fixture(scope="session")
def sample_pdf(tmp_path_factory):
    """合成答案 6 人分（3人目は上下逆・6人目は横向き）"""
    out = tmp_path_factory.mktemp("sample") / "小テスト.pdf"
    pdf, truth = make_exam(out, students=6, seed=3)
    return pdf, json.loads(truth.read_text(encoding="utf-8"))


@pytest.fixture
def project_dir(tmp_path, sample_pdf):
    """PDF だけが置かれた試験フォルダ"""
    root = tmp_path / "期末テスト"
    root.mkdir()
    shutil.copy(sample_pdf[0], root / sample_pdf[0].name)
    return root


@pytest.fixture(scope="session")
def _oriented_template(tmp_path_factory, sample_pdf):
    """PDF→画像→向き揃えを1回だけ行った原本（PNG の書き出しが遅いので使い回す）"""
    from saiten_llm.cli import main
    root = tmp_path_factory.mktemp("oriented") / "期末テスト"
    root.mkdir()
    shutil.copy(sample_pdf[0], root / sample_pdf[0].name)
    assert main(["pdf", str(root)]) == 0
    assert main(["orient", str(root), "--apply"]) == 0
    return root


@pytest.fixture
def oriented_project(tmp_path, _oriented_template):
    """PDF を画像にし、向きを揃えた試験フォルダ（テストごとの写し）"""
    root = tmp_path / "期末テスト"
    shutil.copytree(_oriented_template, root)
    return root
