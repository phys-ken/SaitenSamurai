"""workspace_cmds.py — 作業フォルダ全体に対するコマンド（試験フォルダではなく）。

  setup-workspace  projects/・templates/・.claude/settings.json などを用意し、
                   合成の答案で動作を確かめる
  audit            Claude Code の会話記録を調べ、エージェントが元の答案を
                   開こうとした形跡がないかを報告する
"""
import json
import re
import shutil
import tempfile
from pathlib import Path

TEMPLATE = Path(__file__).with_name("workspace_template")


def _say(*lines):
    for line in lines:
        print(line, flush=True)


# --- setup-workspace ----------------------------------------------------

def merge_deny(settings_path):
    """雛形の deny を、既存の settings.json に足す（既存の設定は消さない）。

    Returns: 足した規則の数
    """
    template = json.loads((TEMPLATE / "settings.json").read_text(encoding="utf-8"))
    wanted = template["permissions"]["deny"]
    if settings_path.exists():
        data = json.loads(settings_path.read_text(encoding="utf-8"))
    else:
        data = {}
    deny = data.setdefault("permissions", {}).setdefault("deny", [])
    added = [r for r in wanted if r not in deny]
    deny.extend(added)
    settings_path.parent.mkdir(parents=True, exist_ok=True)
    settings_path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return len(added)


def self_check():
    """合成の答案（実在の生徒はいない）で PDF→画像→向き揃え が動くか確かめる"""
    from saiten_llm.cli import main
    from saiten_llm.tools.make_sample_exam import make_exam
    with tempfile.TemporaryDirectory(prefix="saiten_selfcheck_") as tmp:
        root = Path(tmp) / "お試し"
        root.mkdir()
        make_exam(root / "お試し.pdf", students=3, seed=7, noise=False)
        (root / "お試し_truth.json").unlink()
        ok = main(["pdf", str(root)]) == 0 and main(["orient", str(root), "--apply"]) == 0
    return ok


def cmd_setup_workspace(args):
    ws = Path(args.workspace).resolve()
    if not ws.is_dir():
        _say(f"作業フォルダがありません: {ws}")
        return 1
    for name in ("projects", "templates"):
        (ws / name).mkdir(exist_ok=True)
    added = merge_deny(ws / ".claude" / "settings.json")
    _say(f"元の答案を読めないようにする設定: {ws / '.claude' / 'settings.json'}（{added} 件追加）")
    for name in ("AGENTS.md", "CLAUDE.md"):
        if not (ws / name).exists():
            shutil.copy2(TEMPLATE / name, ws / name)
            _say(f"{name} を置きました")
    if args.skip_check:
        return 0
    _say("合成の答案で動作を確かめています…")
    if not self_check():
        _say("動作の確認に失敗しました。表示された内容を先生に伝えてください。")
        return 1
    _say("動作を確かめました。",
         "設定を有効にするため、エージェントを一度終了して、作業フォルダで起動し直してください。")
    return 0


# --- doctor -------------------------------------------------------------

EXPECTED_BRANCH = "feat/llm-workflow"
_CLOUD_WORDS = ("onedrive", "dropbox", "google drive", "googledrive", "icloud", "box sync")


def _version(cmd):
    import subprocess
    try:
        out = subprocess.run(cmd, capture_output=True, text=True, timeout=20)
    except (OSError, subprocess.SubprocessError):
        return None
    return (out.stdout or out.stderr).strip().splitlines()[0] if out.returncode == 0 else None


def doctor_checks(ws):
    """環境の点検。Returns: [(重さ, 項目, 結果, 直し方)]  重さは ok / ng（必須）/ warn"""
    import importlib
    import sys
    from saiten_llm import REPO_ROOT
    checks = []

    def add(ok, item, result, fix="", required=True):
        checks.append(("ok" if ok else ("ng" if required else "warn"), item, result, fix))

    v = sys.version_info
    add(v >= (3, 10), "Python", f"{v.major}.{v.minor}.{v.micro}", "uv venv --python 3.12 で作り直す")
    git = _version(["git", "--version"])
    add(bool(git), "Git", git or "見つかりません", "winget install --id Git.Git -e")
    uv = _version(["uv", "--version"])
    add(bool(uv), "uv", uv or "見つかりません",
        'powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"')
    branch = _version(["git", "-C", str(REPO_ROOT), "rev-parse", "--abbrev-ref", "HEAD"])
    add(branch == EXPECTED_BRANCH, "採点侍のブランチ", branch or "Git の管理下にありません",
        f"app で git checkout {EXPECTED_BRANCH}")
    for module, label, required in (
            ("cv2", "画像処理（OpenCV）", True), ("numpy", "numpy", True),
            ("pandas", "pandas", True), ("openpyxl", "Excel（openpyxl）", True),
            ("sklearn", "読み取り（scikit-learn）", True), ("PIL", "画像（Pillow）", True),
            ("fitz", "PDF（PyMuPDF）", True), ("webview", "採点侍の画面（pywebview）", True),
            ("matplotlib", "分析の図（matplotlib）", False), ("reportlab", "PDF レポート（reportlab）", False)):
        try:
            importlib.import_module(module)
            add(True, label, "入っています")
        except Exception as e:  # noqa: BLE001  入っていない・壊れている
            add(False, label, f"読み込めません（{type(e).__name__}）",
                "app で uv pip install -r requirements.txt -r webui/requirements.txt", required)
    from constants import find_japanese_font
    font = find_japanese_font()
    add(bool(font), "日本語フォント", font or "見つかりません",
        "採点済み答案の文字が小さくなります。Windows なら通常は入っています", required=False)

    settings = ws / ".claude" / "settings.json"
    wanted = json.loads((TEMPLATE / "settings.json").read_text(encoding="utf-8"))["permissions"]["deny"]
    try:
        have = json.loads(settings.read_text(encoding="utf-8")).get("permissions", {}).get("deny", [])
        missing = [r for r in wanted if r not in have]
    except (OSError, ValueError):
        missing = wanted
    add(not missing, "元の答案を読めなくする設定", "そろっています" if not missing else
        f"{len(missing)} 件足りません", "app で uv run python -m saiten_llm setup-workspace ..")
    add((ws / "AGENTS.md").exists(), "入口の AGENTS.md", "あります" if (ws / "AGENTS.md").exists()
        else "ありません", "先生に AGENTS.md を置いてもらう")
    for name in ("projects", "templates"):
        d = ws / name
        try:
            probe = d / ".saiten_write_test"
            probe.write_text("x", encoding="utf-8")
            probe.unlink()
            add(True, f"{name}/ に書き込み", "できます")
        except OSError:
            add(False, f"{name}/ に書き込み", "できません", "setup-workspace を実行する")
    cloud = [w for w in _CLOUD_WORDS if w in str(ws).lower()]
    add(not cloud, "クラウド同期の外にあるか",
        "外にあります" if not cloud else f"{cloud[0]} の下にあります（答案がクラウドへ送られます）",
        "同期していないフォルダ（例: ドキュメント直下で同期を切ったもの）へ移す")
    add(REPO_ROOT.parent == ws, "採点侍の場所", str(REPO_ROOT),
        "作業フォルダの app/ に置く（AGENTS.md の手順）", required=False)
    return checks


def cmd_doctor(args):
    ws = Path(args.workspace).resolve()
    if not ws.is_dir():
        _say(f"作業フォルダがありません: {ws}")
        return 1
    checks = doctor_checks(ws)
    mark = {"ok": "✓", "ng": "✗", "warn": "△"}
    for level, item, result, fix in checks:
        _say(f"{mark[level]} {item}: {result}" + (f"\n    → {fix}" if level != "ok" and fix else ""))
    ng = [c for c in checks if c[0] == "ng"]
    if ng:
        _say("", f"✗ が {len(ng)} 件あります。直すまで採点を始めないでください（先生に伝えて許可を得てから直す）。")
        return 1
    _say("", "環境は整っています。")
    return 0


# --- audit --------------------------------------------------------------

def encode_cwd(path):
    """Claude Code が会話記録のフォルダ名にするときの変換（英数字以外を - に）"""
    return re.sub(r"[^A-Za-z0-9]", "-", str(path))


_PROTECTED = re.compile(r"(^|[\\/])(scans|除外)([\\/]|$)|\.pdf\b", re.IGNORECASE)
_LOCAL_URL = re.compile(r"(127\.0\.0\.1|localhost)", re.IGNORECASE)
# 中身を読まないコマンド（PDF を試験フォルダへ移すのは手順書の指示どおり）
_NON_READING = {"mv", "move", "Move-Item", "mkdir", "md", "ls", "dir", "cd"}


def find_tool_uses(transcript):
    """会話記録（JSONL）から、エージェントのツール呼び出しを取り出す"""
    with open(transcript, encoding="utf-8", errors="replace") as f:
        for line in f:
            try:
                entry = json.loads(line)
            except ValueError:
                continue
            msg = entry.get("message") or {}
            if entry.get("type") != "assistant" or not isinstance(msg.get("content"), list):
                continue
            for block in msg["content"]:
                if isinstance(block, dict) and block.get("type") == "tool_use":
                    yield entry.get("timestamp", ""), block.get("name", ""), block.get("input") or {}


def suspicious(name, tool_input):
    """元の答案に触れようとした呼び出しなら、その理由を返す"""
    if name in ("Read", "Grep", "Glob", "NotebookRead", "Edit", "Write"):
        for key in ("file_path", "path", "pattern", "notebook_path"):
            value = str(tool_input.get(key, ""))
            if value and _PROTECTED.search(value):
                return f"{name} {value}"
    if name == "Bash":
        cmd = str(tool_input.get("command", ""))
        first_line = cmd.strip().splitlines()[0][:160] if cmd.strip() else ""
        for segment in re.split(r"&&|\|\||[;|\n]", cmd):
            words = segment.split()
            if not words or "saiten_llm" in segment:
                continue    # 採点用のコマンド（元画像を読むのはコマンドの中だけ）
            if words[0] in _NON_READING and _PROTECTED.search(segment):
                continue    # 移す・フォルダを作る・名前を並べるだけ（中身は読まない）
            if _PROTECTED.search(segment):
                return f"Bash {first_line}"
            if _LOCAL_URL.search(segment) and words[0] in ("curl", "wget", "Invoke-WebRequest", "iwr"):
                return f"Bash {first_line}"
    if name == "WebFetch" and _LOCAL_URL.search(str(tool_input.get("url", ""))):
        return f"WebFetch {tool_input.get('url')}"
    return None


def cmd_audit(args):
    ws = Path(args.workspace).resolve()
    base = Path(args.transcripts) if args.transcripts else Path.home() / ".claude" / "projects"
    folder = base / encode_cwd(ws)
    sessions = sorted(folder.glob("*.jsonl")) if folder.is_dir() else []
    if not sessions:
        _say(f"会話記録が見つかりません: {folder}",
             "（Claude Code をこの作業フォルダで起動した記録が無い。Antigravity の記録は調べられません）")
        return 2
    checked, findings = 0, []
    for s in sessions:
        for ts, name, tool_input in find_tool_uses(s):
            checked += 1
            reason = suspicious(name, tool_input)
            if reason:
                findings.append((s.name, ts, reason))
    _say(f"会話 {len(sessions)} 件・ツール呼び出し {checked} 回を調べました（{folder}）。")
    if checked == 0:
        _say("ツール呼び出しが1回も見つかりません。会話記録の形式が変わった可能性があるので、"
             "この点検は当てになりません。")
        return 2
    if not findings:
        _say("元の答案（scans/・除外/・PDF）や黒塗り画面の画像に触れようとした形跡はありませんでした。")
        return 0
    _say(f"⚠ 元の答案などに触れようとした形跡が {len(findings)} 件あります"
         "（設定で止められた呼び出しも含みます）:")
    for session, ts, reason in findings:
        _say(f"  {ts} {session}: {reason}")
    return 2
