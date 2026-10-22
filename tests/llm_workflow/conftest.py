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
    # 雑音なし: 画像が小さくなり、テストごとの写しで /tmp を食い潰さない
    pdf, truth = make_exam(out, students=6, seed=3, noise=False)
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


NAME_REGION = [0.24, 0.2, 0.57, 0.255]   # 合成答案の氏名欄（補正後の用紙に対する割合）
COORD = REPO / "sample_basefile" / "M2-03-002_座標ファイル.xlsx"


@pytest.fixture(scope="session")
def _recognized_template(tmp_path_factory, _oriented_template):
    """黒塗り承認・LLM用の写し・モード・座標・読み取りまで済ませた原本"""
    from saiten_llm import mask as maskmod
    from saiten_llm.cli import main
    from saiten_llm.workspace import Project
    root = tmp_path_factory.mktemp("recognized") / "期末テスト"
    shutil.copytree(_oriented_template, root)
    project = Project(root)
    maskmod.save_spec(project, {"regions": [NAME_REGION], "extra": {}}, approved=True)
    maskmod.build_llm_view(project)
    assert main(["mode", str(root), "--set", "mark_and_descriptive", "--skip", "4",
                 "--teacher-approved"]) == 0
    assert main(["coord", str(root), str(COORD)]) == 0
    assert main(["recognize", str(root)]) == 0
    return root


@pytest.fixture
def recognized_project(tmp_path, _recognized_template):
    """読み取りまで済んだ試験フォルダ（テストごとの写し）。

    採点侍の設定（session_state.json）は絶対パスを持つので、写した先に合わせて直す
    """
    from saiten_llm.engine import session_path
    from saiten_llm.workspace import Project
    root = tmp_path / "期末テスト"
    shutil.copytree(_recognized_template, root)
    path = session_path(Project(root))
    data = json.loads(path.read_text(encoding="utf-8"))
    data["image_folder"] = str(root / "scans")
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return root
