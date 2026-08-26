"""preview.py — 開発用: Xvfb 上で webui を起動してスクリーンショットを撮る。

使い方: xvfb-run -a -s "-screen 0 1400x900x24" python webui/tools/preview.py out.png
tk 時代のキャプチャ評価サイクル（撮る→目視→直す）を webui でも回すための道具。
"""
import sys
import threading
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import webview  # noqa: E402
from app import create_app  # noqa: E402


def _setup_demo(bridge):
    """合成答案でデータソース選択＋認識実行まで自動で流す（目視評価用）。

    ダイアログ応答を偽装して bridge の正規経路をそのまま通す。
    """
    import tempfile
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent / "tests"))
    import cv2
    from test_multidigit_image_e2e import make_coord_xlsx, make_sheet, sym_positions

    tmp = Path(tempfile.mkdtemp(prefix="saiten_demo_"))
    coord = make_coord_xlsx(tmp / "coord.xlsx", 3)
    scans = tmp / "scans"
    scans.mkdir()
    s1 = {q: p for q, p in zip([1, 2, 3], sym_positions('-24'))}
    s2 = {1: sym_positions('9')[0], 3: sym_positions('4')[0]}
    for name, filled in [("s1.png", s1), ("s2.png", s2)]:
        cv2.imwrite(str(scans / name), make_sheet(filled, with_markers=True))

    # 正答データも用意（採点・集計のデモ用）
    from test_multi_digit_mode import _create_answer_key
    key = tmp / "answer_key.xlsx"
    _create_answer_key(key, [
        {'問題番号': 1, '正答': '-24', '配点': 3, '観点': 1},
    ])
    file_queue = [str(coord), str(key)]
    real_folder = bridge._win.open_folder_dialog
    real_file = bridge._win.open_file_dialog
    bridge._win.open_folder_dialog = lambda **kw: str(scans)
    bridge._win.open_file_dialog = lambda **kw: file_queue.pop(0) if file_queue else None
    time.sleep(2.5)  # 画面初期化を待ってから UI 操作
    bridge._win.eval_js("document.querySelector('.mode-card.math').click()")
    time.sleep(0.5)
    bridge.set_skip_questions(0)
    bridge._win.eval_js("document.querySelector('[data-action=select_image_folder]').click()")
    time.sleep(0.6)
    bridge._win.eval_js("document.querySelector('[data-action=select_coord_file]').click()")
    time.sleep(0.6)
    bridge._win.eval_js("document.querySelector('[data-action=select_answer_key]').click()")
    time.sleep(0.6)
    bridge._win.eval_js("document.getElementById('btn-run-recognition').click()")
    time.sleep(3)
    if "--full" in sys.argv:
        bridge._win.eval_js("document.getElementById('btn-run-scoring').click()")
        time.sleep(4)
        bridge._win.eval_js("document.getElementById('btn-run-summary').click()")
        time.sleep(6)
    if "--checker" in sys.argv:
        bridge._win.eval_js("document.getElementById('btn-open-checker').click()")
        time.sleep(2.5)
    bridge._win.open_folder_dialog = real_folder
    bridge._win.open_file_dialog = real_file


def _setup_desc_demo(bridge, out_prefix):
    """記述のみモード: 画像準備→問題設定→採点グリッド→一枚採点をキャプチャ"""
    import tempfile
    import cv2
    import numpy as np
    from PIL import ImageGrab

    tmp = Path(tempfile.mkdtemp(prefix="saiten_desc_"))
    scans = tmp / "scans"
    scans.mkdir()
    for i, name in enumerate(["s1.png", "s2.png", "s3.png"]):
        img = np.full((842, 595, 3), 255, dtype=np.uint8)
        cv2.putText(img, f"Answer {i+1}", (80, 200), cv2.FONT_HERSHEY_SIMPLEX,
                    1.4, (40, 40, 40), 3)
        cv2.line(img, (60, 320), (520, 320 + i * 30), (60, 60, 60), 2)
        cv2.putText(img, "x = " + str(3 + i), (100, 500),
                    cv2.FONT_HERSHEY_SIMPLEX, 1.2, (30, 30, 90), 3)
        cv2.imwrite(str(scans / name), img)

    bridge._win.open_folder_dialog = lambda **kw: str(scans)
    time.sleep(2.5)
    bridge._win.eval_js("document.querySelector('.mode-card[data-mode=descriptive_only]').click()")
    time.sleep(0.5)
    bridge._win.eval_js("document.querySelector('[data-action=select_image_folder]').click()")
    time.sleep(0.6)
    bridge._win.eval_js("document.getElementById('btn-run-recognition').click()")  # =画像準備
    time.sleep(2)
    bridge.add_descriptive_question("問1", 5, 1, [60, 120, 540, 250])
    bridge.add_descriptive_question("問2", 3, 1, [60, 400, 540, 560])
    # bridge を直接叩いた変更は UI に伝わらないので再同期してからボタンを押す
    bridge._win.eval_js("window.__refreshState()")
    time.sleep(0.5)
    ImageGrab.grab(xdisplay="").save(f"{out_prefix}_main.png")
    print(f"saved: {out_prefix}_main.png")
    bridge._win.eval_js("document.getElementById('btn-desc-scoring').click()")
    time.sleep(2.5)
    ImageGrab.grab(xdisplay="").save(f"{out_prefix}_grid.png")
    print(f"saved: {out_prefix}_grid.png")
    bridge.set_descriptive_score("s1.png", "D1", 5)
    # ドキュメント用に手書きコメントも載せる
    bridge.set_handwriting("s1.png", 595, 842, [
        {"color": "#c73e2e", "width": 3,
         "points": [[90, 300, 0.3], [170, 260, 0.7], [260, 310, 0.9], [340, 265, 0.6]]},
        {"color": "#1d5fa8", "width": 2,
         "points": [[110, 640, 0.5], [420, 640, 0.5]]},
    ])
    bridge._win.eval_js("window.__refreshState()")
    time.sleep(0.5)
    bridge._win.eval_js("document.getElementById('btn-single-sheet').click()")
    time.sleep(2.5)
    ImageGrab.grab(xdisplay="").save(f"{out_prefix}_single.png")
    print(f"saved: {out_prefix}_single.png")


def _shoot_docs(bridge, out_dir):
    """ドキュメントサイト用スクリーンショット一式（docs/images/webui/）"""
    from PIL import ImageGrab
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)

    def snap(name, wait=0.8):
        time.sleep(wait)
        ImageGrab.grab(xdisplay="").save(str(out / name))
        print("saved:", out / name)

    def js(code, wait=0.6):
        bridge._win.eval_js(code)
        time.sleep(wait)

    time.sleep(2.5)
    snap("01_mode_select.png", wait=0.2)

    # 数学マーク採点デモ一式（フォルダ・座標・正答・読み取り）
    _setup_demo(bridge)
    time.sleep(1.0)
    js("window.scrollTo(0, 0)")
    snap("02_main.png")

    js("document.getElementById('output-settings').open = true")
    js("document.getElementById('output-settings-panel').scrollIntoView()")
    snap("03_output_settings.png")

    js("document.getElementById('btn-render-settings').click()", wait=2.0)
    snap("04_render_settings.png")
    js("document.getElementById('btn-rs-close').click()")

    js("document.getElementById('btn-open-map').click()")
    snap("05_files_map.png")
    js("document.getElementById('btn-map-close').click()")

    js("document.getElementById('btn-open-checker').click()", wait=2.0)
    snap("06_checker.png")
    js("document.getElementById('btn-close-checker').click()", wait=1.0)


def main():
    out = Path(sys.argv[1] if len(sys.argv) > 1 else "webui_preview.png")
    demo = "--demo" in sys.argv
    window, bridge = create_app()

    def shoot():
        if "--docs" in sys.argv:
            _shoot_docs(bridge, str(out))
            window.destroy()
            return
        if "--descdemo" in sys.argv:
            _setup_desc_demo(bridge, str(out.with_suffix("")))
            window.destroy()
            return
        if demo:
            _setup_demo(bridge)
            time.sleep(2)
        else:
            time.sleep(4)  # 描画待ち
        try:
            from PIL import ImageGrab
            img = ImageGrab.grab(xdisplay="")
            img.save(out)
            print(f"saved: {out} ({img.size[0]}x{img.size[1]})")
        finally:
            window.destroy()

    threading.Thread(target=shoot, daemon=True).start()
    webview.start()


if __name__ == "__main__":
    main()
