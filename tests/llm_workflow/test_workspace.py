"""作業フォルダの用意・会話記録の点検・進捗.md・手順書（S7）"""
import json
import re
from pathlib import Path

from saiten_llm import cli
from saiten_llm.cli import main
from saiten_llm.workspace_cmds import encode_cwd, suspicious

GUIDE = Path(cli.__file__).with_name("GUIDE.md")


def test_setup_workspace_creates_and_merges(tmp_path):
    ws = tmp_path / "採点"
    (ws / ".claude").mkdir(parents=True)
    (ws / ".claude" / "settings.json").write_text(json.dumps(
        {"model": "x", "permissions": {"deny": ["Read(./secret)"], "allow": ["Bash(ls)"]}}),
        encoding="utf-8")
    (ws / "AGENTS.md").write_text("先生が置いた入口", encoding="utf-8")
    assert main(["setup-workspace", str(ws), "--skip-check"]) == 0
    assert (ws / "projects").is_dir() and (ws / "templates").is_dir()
    data = json.loads((ws / ".claude" / "settings.json").read_text(encoding="utf-8"))
    assert data["model"] == "x" and data["permissions"]["allow"] == ["Bash(ls)"]
    deny = data["permissions"]["deny"]
    assert "Read(./secret)" in deny and "Read(/projects/*/scans/**)" in deny
    assert (ws / "AGENTS.md").read_text(encoding="utf-8") == "先生が置いた入口"   # 上書きしない
    assert "@AGENTS.md" in (ws / "CLAUDE.md").read_text(encoding="utf-8")
    # 2回目は何も足さない
    assert main(["setup-workspace", str(ws), "--skip-check"]) == 0
    again = json.loads((ws / ".claude" / "settings.json").read_text(encoding="utf-8"))
    assert again["permissions"]["deny"] == deny


def _doctor(ws):
    from saiten_llm.workspace_cmds import doctor_checks
    return {item: level for level, item, _r, _f in doctor_checks(ws)}


def test_doctor_requires_settings_and_folders(tmp_path):
    ws = tmp_path / "採点"
    ws.mkdir()
    before = _doctor(ws)
    assert before["元の答案を読めなくする設定"] == "ng"
    assert before["projects/ に書き込み"] == "ng"
    assert main(["doctor", str(ws)]) == 1
    (ws / "AGENTS.md").write_text("入口", encoding="utf-8")
    assert main(["setup-workspace", str(ws), "--skip-check"]) == 0
    after = _doctor(ws)
    for item in ("元の答案を読めなくする設定", "入口の AGENTS.md", "projects/ に書き込み",
                 "templates/ に書き込み", "Python", "PDF（PyMuPDF）", "クラウド同期の外にあるか"):
        assert after[item] == "ok", item


def test_doctor_refuses_cloud_synced_folder(tmp_path):
    ws = tmp_path / "OneDrive" / "ドキュメント" / "採点"
    ws.mkdir(parents=True)
    assert main(["setup-workspace", str(ws), "--skip-check"]) == 0
    assert _doctor(ws)["クラウド同期の外にあるか"] == "ng"


def test_setup_workspace_self_check(tmp_path):
    assert main(["setup-workspace", str(tmp_path)]) == 0


def _transcript(base, ws, tool_uses):
    folder = base / encode_cwd(ws)
    folder.mkdir(parents=True)
    lines = [json.dumps({"type": "user", "message": {"content": "採点を始める"}})]
    for name, tool_input in tool_uses:
        lines.append(json.dumps({"type": "assistant", "timestamp": "2026-10-22T00:00:00Z",
                                 "message": {"content": [{"type": "tool_use", "name": name,
                                                          "input": tool_input}]}}))
    (folder / "s1.jsonl").write_text("\n".join(lines) + "\n", encoding="utf-8")


def test_encode_cwd_matches_claude_code():
    assert encode_cwd("/home/kenmini/share/saitensamurai_dev") == "-home-kenmini-share-saitensamurai-dev"
    assert encode_cwd("C:\\Users\\t\\Documents\\採点") == "C--Users-t-Documents---"


def test_suspicious_rules():
    flagged = [
        ("Read", {"file_path": "C:/w/projects/期末/scans/p001.png"}),
        ("Read", {"file_path": "/w/projects/期末/答案.pdf"}),
        ("Grep", {"path": "projects/期末/除外"}),
        ("Bash", {"command": "cat projects/期末/scans/_saiten_grading_results/01_Results/x.json"}),
        ("Bash", {"command": "mkdir x && cp projects/期末/scans/p001.png /tmp/"}),
        ("Bash", {"command": "curl -s http://127.0.0.1:5123/img?i=0"}),
        ("WebFetch", {"url": "http://localhost:5123/"}),
    ]
    allowed = [
        ("Read", {"file_path": "/w/projects/期末/_llm_view/pages/p001.jpg"}),
        ("Bash", {"command": "cd app && uv run python -m saiten_llm pdf ../projects/期末"}),
        ("Bash", {"command": "mkdir -p projects/期末 && mv 期末.pdf projects/期末/"}),
        ("Bash", {"command": "ls projects/期末"}),
    ]
    for name, inp in flagged:
        assert suspicious(name, inp), (name, inp)
    for name, inp in allowed:
        assert suspicious(name, inp) is None, (name, inp)


def test_audit_reports_findings(tmp_path, capsys):
    ws, base = tmp_path / "ws", tmp_path / "transcripts"
    ws.mkdir()
    _transcript(base, ws.resolve(), [
        ("Bash", {"command": "uv run python -m saiten_llm orient ../projects/a --apply"}),
        ("Read", {"file_path": str(ws / "projects/a/_llm_view/pages/p.jpg")}),
        ("Read", {"file_path": str(ws / "projects/a/scans/p.png")}),
    ])
    assert main(["audit", str(ws), "--transcripts", str(base)]) == 2
    out = capsys.readouterr().out
    assert "ツール呼び出し 3 回" in out and "scans/p.png" in out


def test_audit_clean_and_missing(tmp_path, capsys):
    ws, base = tmp_path / "ws", tmp_path / "transcripts"
    ws.mkdir()
    assert main(["audit", str(ws), "--transcripts", str(base)]) == 2     # 記録が無い
    _transcript(base, ws.resolve(), [("Bash", {"command": "ls projects"})])
    assert main(["audit", str(ws), "--transcripts", str(base)]) == 0
    assert "形跡はありませんでした" in capsys.readouterr().out


def test_audit_without_tool_calls_is_not_trusted(tmp_path):
    ws, base = tmp_path / "ws", tmp_path / "transcripts"
    ws.mkdir()
    _transcript(base, ws.resolve(), [])
    assert main(["audit", str(ws), "--transcripts", str(base)]) == 2


def test_progress_file_keeps_memo(recognized_project):
    path = recognized_project / "進捗.md"
    assert main(["status", str(recognized_project), "--write"]) == 0
    text = path.read_text(encoding="utf-8")
    assert "- [x] マークの読み取り" in text and "- [ ] 配点を先生が決定" in text
    path.write_text(text.replace("（先生と決めたこと・気をつけることを書き足す）",
                                 "満点は50点"), encoding="utf-8")
    assert main(["status", str(recognized_project), "--write"]) == 0
    assert "満点は50点" in path.read_text(encoding="utf-8")


def test_guide_mentions_only_real_commands():
    """手順書に書いたコマンドが実在する（名前を変えたら手順書も直す）"""
    known = set(cli.COMMANDS) | set(cli.WORKSPACE_COMMANDS)
    used = set(re.findall(r"^\s*([a-z][a-z-]+)\s+\.\./projects", GUIDE.read_text(encoding="utf-8"),
                          re.MULTILINE))
    used |= set(re.findall(r"`([a-z][a-z-]+) \.\.\.", GUIDE.read_text(encoding="utf-8")))
    assert len(used) >= 12, used      # 抜き出しが空振りしていないこと
    assert used <= known, used - known
