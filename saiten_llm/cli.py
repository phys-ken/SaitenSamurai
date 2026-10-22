"""cli.py — LLM が採点侍を動かすためのコマンド。

使い方（採点侍のフォルダで）:
    uv run python -m saiten_llm <コマンド> <試験フォルダ> [オプション]

コマンドは画像を出力しない。結果は文字で返し、LLM に見せてよい画像は
_llm_view/ にだけ書く。終了コード: 0=成功 / 1=失敗 / 2=先生の確認が必要
"""
import argparse
import subprocess
import sys
import webbrowser
from pathlib import Path

from saiten_llm import (REPO_ROOT, cmd_desc, cmd_finalize, cmd_mark, cmd_points,
                        workspace_cmds)
from saiten_llm.workspace import Project, now_iso

EXIT_OK, EXIT_ERROR, EXIT_NEEDS_TEACHER = 0, 1, 2


def _say(*lines):
    for line in lines:
        print(line)


# --- status -------------------------------------------------------------

def cmd_status(project, args):
    from saiten_llm.mask import mask_status
    scans = project.scan_files()
    orient = project.data.get("orientation") or {}
    view = project.data.get("llm_view") or {}
    _say(f"試験フォルダ: {project.root}",
         f"PDF: {', '.join(p.name for p in project.pdfs()) or 'なし'}",
         f"答案の画像: {len(scans)} 枚",
         f"向き: {'確認済み（' + orient['checked_at'] + '）' if orient.get('checked_at') else '未確認'}",
         f"黒塗り: {_MASK_LABEL[mask_status(project)]}",
         f"LLM用の写し: {str(view.get('pages')) + ' 枚（' + view['built_at'] + '）' if view else 'まだ作っていません'}")
    from saiten_llm import progress
    _say("", *progress.checklist(project))
    if args.write:
        _say("", f"書き出しました: {progress.write(project)}")
    return EXIT_OK


_MASK_LABEL = {
    "none": "未設定",
    "draft": "範囲は保存済み・先生の承認待ち",
    "approved": "先生が承認済み",
    "stale": "承認の後で答案が変わった（もう一度確かめる必要あり）",
}


# --- pdf ----------------------------------------------------------------

def cmd_pdf(project, args):
    from constants import extract_pdf_to_images
    pdfs = project.pdfs()
    if not pdfs:
        _say(f"試験フォルダに PDF がありません: {project.root}")
        return EXIT_ERROR
    existing = project.scan_files()
    if existing and not args.force:
        _say(f"scans/ にはすでに {len(existing)} 枚の画像があります。",
             "PDF から作り直すときは --force を付けてください（向きの修正は消えます）。")
        return EXIT_ERROR
    for p in existing:
        p.unlink()
    project.scans.mkdir(exist_ok=True)
    for pdf in pdfs:
        extract_pdf_to_images(pdf, output_folder=project.scans, dpi=args.dpi)
        _say(f"画像にしました: {pdf.name}")
    project.data["pdf"] = {"files": [p.name for p in pdfs], "dpi": args.dpi,
                           "pages": len(project.scan_files()), "extracted_at": now_iso()}
    project.data.pop("orientation", None)
    project.save()
    _say(f"答案の画像: {len(project.scan_files())} 枚 → {project.scans}",
         "次は orient で向きを確かめてください。")
    return EXIT_OK


# --- orient -------------------------------------------------------------

_METHOD_LABEL = {"marker": "マーカー", "similarity": "基準の答案との比較",
                 "reference": "基準の答案", "none": "判定できず"}


def cmd_orient(project, args):
    from omr_engine import find_sheet_markers, imread_unicode
    from saiten_llm.orient import apply_rotation, check_folder
    files = project.scan_files()
    if not files:
        _say("答案の画像がありません。先に pdf を実行してください。")
        return EXIT_ERROR

    if args.flip_all:
        for p in files:
            apply_rotation(p, 180)
        _say(f"全 {len(files)} 枚を上下逆にしました。")

    report = check_folder(files)
    pages = report["pages"]
    to_rotate = {n: r["rotation"] for n, r in pages.items()
                 if r.get("rotation") and "error" not in r}
    review = [n for n, r in pages.items() if r.get("needs_review") and not r.get("blank")]
    blank = [n for n, r in pages.items() if r.get("blank")]
    errors = [n for n, r in pages.items() if "error" in r]
    methods = {}
    for r in pages.values():
        methods[r.get("method", "error")] = methods.get(r.get("method", "error"), 0) + 1

    _say(f"答案 {len(files)} 枚の向きを調べました（基準: {report['reference']}）。",
         "判定の方法: " + "、".join(f"{_METHOD_LABEL.get(m, m)} {n}枚" for m, n in methods.items()))
    for name, rot in to_rotate.items():
        _say(f"  回す必要あり: {name} → 時計回りに {rot}°")
    for name in review:
        r = pages[name]
        _say(f"  要確認（向きを決めきれない）: {name}"
             + (f"（差 {r['margin']}）" if r.get("margin") is not None else ""))
    for name in blank:
        _say(f"  白紙の可能性: {name}")
    for name in errors:
        _say(f"  読めない画像: {name}")

    if to_rotate and args.apply:
        for name, rot in to_rotate.items():
            apply_rotation(project.scans / name, rot)
        _say(f"{len(to_rotate)} 枚を回しました。")
    elif to_rotate:
        _say("まだ回していません。回すときは --apply を付けて実行してください。")

    # 傾き補正ができる用紙か（マーカーの有無）を数える
    applied = args.apply or not to_rotate
    aligned = None
    if applied:
        aligned = sum(1 for p in files
                      if (img := imread_unicode(str(p))) is not None
                      and find_sheet_markers(img) is not None)
        _say(f"四隅のマーカーが見つかった答案: {aligned} / {len(files)} 枚")
        if 0 < aligned < len(files):
            _say("  ⚠ マーカーが見つからない答案が混ざっています。読み取りや記述の範囲がずれる"
                 "おそれがあるので、先生に確認してください。")

    # どの答案を回したかは、何度実行しても積み上げて残す（角度は合計）
    rotated = dict((project.data.get("orientation") or {}).get("rotated", {}))
    if args.flip_all:
        for p in files:
            rotated[p.name] = (rotated.get(p.name, 0) + 180) % 360
    if args.apply:
        for name, rot in to_rotate.items():
            rotated[name] = (rotated.get(name, 0) + rot) % 360
    rotated = {n: r for n, r in rotated.items() if r}
    project.data["orientation"] = {
        "checked_at": now_iso(), "applied": applied, "reference": report["reference"],
        "rotated": rotated, "pending": {} if applied else to_rotate,
        "needs_review": review, "blank": blank, "errors": errors,
        "with_markers": aligned,
    }
    project.save()
    if review or blank or errors or (aligned is not None and 0 < aligned < len(files)):
        _say("先生の確認が必要な答案があります。上の一覧を先生に伝えてください。",
             "1枚目（基準）が正しい向きかどうかは、黒塗りの画面で先生が確かめます。")
        return EXIT_NEEDS_TEACHER
    if not applied:
        return EXIT_NEEDS_TEACHER
    _say("向きは揃っています。次は mask-server で黒塗りの範囲を先生に決めてもらってください。")
    return EXIT_OK


# --- mask-server / llm-view ----------------------------------------------

def cmd_mask_server(project, args):
    from saiten_llm.mask import build_llm_view
    from saiten_llm.mask_server import MaskServer
    orient = project.data.get("orientation") or {}
    if not orient.get("applied"):
        _say("先に orient で向きを揃えてください（回す答案が残っています）。")
        return EXIT_ERROR
    server = MaskServer(project, timeout_minutes=args.timeout)
    # バックグラウンドで実行されても LLM にすぐ URL が届くよう、ここだけ flush する
    print("黒塗りの範囲を決める画面を用意しました。先生にこのリンクを開いてもらってください:")
    print(f"  {server.url}")
    print(f"（このPCからしか開けません。承認されるか {args.timeout} 分たつと止まります）",
          flush=True)
    if not args.no_open:
        try:
            webbrowser.open(server.url)
        except Exception:
            pass
    if not server.serve_until_done():
        _say("承認されないまま時間切れになりました。範囲は下書きとして保存されています。")
        return EXIT_NEEDS_TEACHER
    count = build_llm_view(project)
    _say("先生が黒塗りを承認しました。",
         f"LLM用の写しを作りました: {project.llm_view / 'pages'}（{count} 枚）")
    return EXIT_OK


def cmd_llm_view(project, args):
    from saiten_llm.mask import build_llm_view
    try:
        count = build_llm_view(project)
    except PermissionError as e:
        _say(str(e))
        return EXIT_NEEDS_TEACHER
    _say(f"LLM用の写しを作りました: {project.llm_view / 'pages'}（{count} 枚）")
    return EXIT_OK


# --- open-ui ------------------------------------------------------------

def cmd_open_ui(project, args):
    """採点侍の画面を、この試験の状態で開く（先生の番の始まり）"""
    from api import uilock
    from saiten_llm.engine import session_path
    if uilock.holder(project.scans):
        _say("採点侍の画面はすでにこの試験を開いています。")
        return EXIT_NEEDS_TEACHER
    if not session_path(project).exists():
        _say("まだ採点侍の設定がありません。先に mode でモードを決めてください。")
        return EXIT_ERROR
    cmd = [sys.executable, str(REPO_ROOT / "webui" / "app.py"), "--folder", str(project.scans)]
    proc = subprocess.Popen(cmd, cwd=str(REPO_ROOT))
    _say(f"採点侍を開きました（プロセス {proc.pid}）。",
         "先生が画面を閉じるまで、LLM は採点のファイルに書き込みません。")
    return EXIT_OK


def _args_status(p):
    p.add_argument("--write", action="store_true", help="進捗.md を書き直す（メモ欄は残す）")


def _args_pdf(p):
    p.add_argument("--force", action="store_true", help="scans/ を作り直す")
    p.add_argument("--dpi", type=int, default=200)


def _args_orient(p):
    p.add_argument("--apply", action="store_true", help="回す必要のある答案を回す")
    p.add_argument("--flip-all", action="store_true",
                   help="全答案を上下逆にする（基準の答案が逆さだったとき）")


def _args_mask_server(p):
    p.add_argument("--timeout", type=int, default=120, help="待つ時間（分）")
    p.add_argument("--no-open", action="store_true", help="ブラウザを自動で開かない")


def _args_facts(p):
    p.add_argument("--templates", help="座標ファイルを控えるフォルダ（既定: 作業フォルダの templates/）")


def _args_mode(p):
    p.add_argument("--set", required=True, choices=list(cmd_mark.MODES))
    p.add_argument("--format", default="standard", choices=list(cmd_mark.FORMATS))
    p.add_argument("--skip", type=int, help="ID欄（学年・クラス・番号など）の行数")
    p.add_argument("--teacher-approved", action="store_true",
                   help="先生がこのモードを承認した（無いと実行しない）")


def _args_coord(p):
    p.add_argument("file", help="座標ファイル（Mark2 の座標 Excel）")
    p.add_argument("--skip", type=int, help="ID欄（学年・クラス・番号など）の行数")
    p.add_argument("--templates", help="座標ファイルを控えるフォルダ（既定: 作業フォルダの templates/）")


def _args_answer_key(p):
    src = p.add_mutually_exclusive_group(required=True)
    src.add_argument("--from-xlsx", help="先生が埋めた answer_key.xlsx（書き換えずに使う）")
    src.add_argument("--from-json", help="LLM が書き写した正答の JSON")
    src.add_argument("--from-sheet", help="正答をマークした用紙のファイル名（読み取り済みのもの）")


def _args_set_aside(p):
    p.add_argument("filename", help="scans/ の中のファイル名")
    p.add_argument("--reason", default="", help="外す理由（例: 正答用の用紙）")


def _args_desc_set(p):
    p.add_argument("--input", required=True,
                   help='{"questions": [{"name": "問1", "region": [x1, y1, x2, y2]}]}'
                        "（範囲は _llm_view/pages の画像の画素）")
    p.add_argument("--replace", action="store_true", help="今の記述問題を消して作り直す（得点が無いときだけ）")


def _args_crops(p):
    p.add_argument("--question", help="問題ID（D1 など）。省略で全問")
    p.add_argument("--per-sheet", type=int, default=15, help="1枚に並べる人数（10〜20）")


def _args_judge(p):
    p.add_argument("--question", required=True, help="問題ID（D1 など）")
    p.add_argument("--input", required=True,
                   help='{"番号": {"verdict": "正|部分|誤|保留", "reason": "..."}}')


def _args_review(p):
    p.add_argument("--question", help="問題ID。省略で全問")


def _args_confirm(p):
    p.add_argument("what", choices=["question", "regions", "answer-key", "marks"],
                   help="何を確定したか")
    p.add_argument("--question", help="問題ID（question のとき）")


def _args_stats(p):
    p.add_argument("--total", type=int, default=100, help="満点（既定 100）")


def _args_set_points(p):
    p.add_argument("--input", required=True,
                   help='{"marks": {"1": 2, "2": 3}, "descriptive": {"D1": 10}}'
                        '（マークの配点を answer_key.xlsx のまま使うなら "marks": "keep"）')
    p.add_argument("--teacher-approved", action="store_true",
                   help="先生がこの配点を承認した（無いと実行しない）")


COMMANDS = {
    "status": (cmd_status, "試験フォルダの進み具合を表示する", _args_status),
    "pdf": (cmd_pdf, "試験フォルダの PDF を scans/ に画像として書き出す", _args_pdf),
    "orient": (cmd_orient, "答案の向きを確かめる（--apply で回す）", _args_orient),
    "mask-server": (cmd_mask_server, "黒塗りの範囲を先生に決めてもらう画面を開き、承認まで待つ",
                    _args_mask_server),
    "llm-view": (cmd_llm_view, "承認済みの黒塗りで LLM 用の写しを作り直す", None),
    "facts": (cmd_mark.cmd_facts, "モードを決めるための事実（マーカー・マーク欄）を集める", _args_facts),
    "mode": (cmd_mark.cmd_mode, "先生が承認したモードを設定する", _args_mode),
    "coord": (cmd_mark.cmd_coord, "座標ファイルを設定し、黒塗りした1枚目に枠を重ねる", _args_coord),
    "recognize": (cmd_mark.cmd_recognize, "マークを読み取り、確認の要る答えを数える", None),
    "answer-key": (cmd_mark.cmd_answer_key, "正答・配点（answer_key.xlsx）を用意してチェックする",
                   _args_answer_key),
    "set-aside": (cmd_mark.cmd_set_aside, "答案を集計から外す（正答用の用紙など）", _args_set_aside),
    "prepare": (cmd_desc.cmd_prepare, "記述のみのモードで答案を準備する（傾き補正）", None),
    "desc-set": (cmd_desc.cmd_desc_set, "記述問題と範囲の案を書き込む（仮の2点満点）", _args_desc_set),
    "crops": (cmd_desc.cmd_crops, "同じ問題の切り抜きを番号付きで並べた画像を作る", _args_crops),
    "judge": (cmd_desc.cmd_judge, "LLM の判定（正/部分/誤/保留）を書き込む", _args_judge),
    "review": (cmd_desc.cmd_review, "判定の状況・保留・先生の修正を一覧にする", _args_review),
    "confirm": (cmd_desc.cmd_confirm, "先生がチャットで伝えた確認を記録する", _args_confirm),
    "accuracy": (cmd_desc.cmd_accuracy, "LLM の判定と先生の確定の一致を数える", None),
    "stats": (cmd_points.cmd_stats, "正答率・識別力と配点案を出す（配点は先生が決める）", _args_stats),
    "set-points": (cmd_points.cmd_set_points, "先生が承認した配点を書き込み、点数に換算する",
                   _args_set_points),
    "finalize": (cmd_finalize.cmd_finalize, "確認がそろっていれば個票（採点済み答案）と集計を作る", None),
    "open-ui": (cmd_open_ui, "採点侍の画面をこの試験の状態で開く", None),
}


def _args_setup_workspace(p):
    p.add_argument("--skip-check", action="store_true", help="合成の答案での動作確認を省く")


def _args_audit(p):
    p.add_argument("--transcripts", help="会話記録のフォルダ（既定: ~/.claude/projects）")


# 試験フォルダではなく、作業フォルダ全体に対するコマンド
WORKSPACE_COMMANDS = {
    "setup-workspace": (workspace_cmds.cmd_setup_workspace,
                        "作業フォルダ（projects/・templates/・設定）を用意して動作を確かめる",
                        _args_setup_workspace),
    "doctor": (workspace_cmds.cmd_doctor,
               "環境を点検する（採点を始める前に必ず実行する）", lambda p: None),
    "audit": (workspace_cmds.cmd_audit,
              "会話記録を調べ、元の答案を開こうとした形跡がないか報告する", _args_audit),
}


def build_parser():
    ap = argparse.ArgumentParser(prog="saiten_llm",
                                 description="LLM と協働で採点するための採点侍コマンド")
    sub = ap.add_subparsers(dest="command", required=True)
    for name, (_fn, help_text, add_args) in COMMANDS.items():
        p = sub.add_parser(name, help=help_text, description=help_text)
        p.add_argument("project", help="試験フォルダ")
        if add_args:
            add_args(p)
    for name, (_fn, help_text, add_args) in WORKSPACE_COMMANDS.items():
        p = sub.add_parser(name, help=help_text, description=help_text)
        p.add_argument("workspace", help="作業フォルダ（AGENTS.md を置いたフォルダ）")
        add_args(p)
    return ap


def main(argv=None):
    from saiten_llm.engine import ScreenOpenError
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")   # Windows の端末でも日本語を崩さない
    args = build_parser().parse_args(argv)
    if args.command in WORKSPACE_COMMANDS:
        return WORKSPACE_COMMANDS[args.command][0](args)
    try:
        project = Project(args.project)
    except FileNotFoundError as e:
        print(e)
        return EXIT_ERROR
    fn, _help, _add = COMMANDS[args.command]
    try:
        return fn(project, args)
    except ScreenOpenError as e:
        print(e)
        return EXIT_NEEDS_TEACHER
    except (RuntimeError, ValueError, FileNotFoundError, TimeoutError) as e:
        print(f"失敗しました: {e}")
        return EXIT_ERROR
