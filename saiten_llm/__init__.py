"""saiten_llm — LLM（Claude Code など）と協働で採点するための CLI。

LLM はこのパッケージのコマンドを通して採点侍を動かす。
元の答案画像を読むのはこのパッケージのコードだけで、LLM に渡すのは
黒塗りした写し（<試験フォルダ>/_llm_view/）だけにする。

設計の決定事項は利用者向け手順書 saiten_llm/GUIDE.md にまとめてある。
"""
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent

# main_src（採点ロジック）と webui（Bridge）を import できるようにする
for _p in (REPO_ROOT / "main_src", REPO_ROOT / "webui"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))
