"""手書きコメント（L2）: マウストグル描画・保存・undo・マークのみ入口。"""

TINY_PNG = ("data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ"
            "AAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==")

API = """(() => {
  window.__hw = {};
  const state = () => ({
    app_mode: 'mark_only', mark_format: 'standard', skip_questions: 4,
    image_folder: '/tmp/scans', image_count: 2,
    coord_file: '/tmp/c.xlsx', coord_summary: {answer_rows: 3, marks_per_row: 10, warning: null},
    answer_key: null, key_summary: null,
    omr_mode: 'kmeans', color_threshold: 0.1, area_threshold: 0.4,
    omr_result: '/tmp/r.xlsx', job: {running: false, kind: null, current: 0, total: 0},
    checker: null, descriptive: null,
  });
  return {
    ping: async () => ({ok: true}),
    set_mode: async () => ({ok: true, state: state()}),
    get_app_info: async () => ({ok: true, app_version: 't', webui_version: 't'}),
    get_state: async () => ({ok: true, state: state()}),
    get_progress: async () => ({ok: true, progress:
      {prepared: true, read: true, scored: false, summarized: false}}),
    list_sheet_files: async () => ({ok: true, files: ['s1.png', 's2.png']}),
    list_sheet_overview: async () => ({ok: true, items: [
      {filename: 's1.png', done: 0, total: 0,
       handwriting: (window.__hw['s1.png'] ?? []).length > 0},
      {filename: 's2.png', done: 0, total: 0,
       handwriting: (window.__hw['s2.png'] ?? []).length > 0},
    ]}),
    get_sheet_image: async (f) => ({ok: true, filename: f ?? 's1.png',
                                    data_url: '%(png)s'}),
    get_handwriting: async (f) => ({ok: true,
      strokes: window.__hw[f] ?? []}),
    set_handwriting: async (f, w, h, strokes) => {
      window.__hw[f] = strokes;
      window.__hwLastSize = [w, h];
      return {ok: true, stroke_count: strokes.length};
    },
    select_image_folder: async () => ({ok: true, cancelled: true}),
  };
})()""" % {"png": TINY_PNG}

from conftest import enter_mode


def _open_annotate(page):
    enter_mode(page)
    page.click("#btn-sheet-review")
    page.wait_for_selector("#sheet-list-view", state="visible")
    page.locator("#sheet-list .sheet-row").first.click()
    page.wait_for_selector("#single-sheet-view", state="visible")
    page.wait_for_selector("#hw-canvas")


def _draw_line(page):
    box = page.locator("#hw-canvas").bounding_box()
    x0, y0 = box["x"] + box["width"] * 0.3, box["y"] + box["height"] * 0.3
    page.mouse.move(x0, y0)
    page.mouse.down()
    page.mouse.move(x0 + 80, y0 + 10, steps=5)
    page.mouse.up()


def test_mark_mode_entry_and_mouse_toggle_drawing(open_app):
    page = open_app(API)
    _open_annotate(page)
    # トグルオフではマウスで描けない（パン扱い）
    _draw_line(page)
    assert page.evaluate("window.__hw['s1.png'] ?? null") is None
    # ✎コメントをオンにすると描けて、保存される
    page.click("#btn-hw-toggle")
    _draw_line(page)
    page.wait_for_function("(window.__hw['s1.png'] ?? []).length === 1")
    stroke = page.evaluate("window.__hw['s1.png'][0]")
    assert stroke["color"] == "#c73e2e" and len(stroke["points"]) >= 2
    # 全消去ボタンが有効になっている
    assert not page.locator("#btn-hw-clear").is_disabled()


def test_undo_redo_roundtrip(open_app):
    page = open_app(API)
    _open_annotate(page)
    page.click("#btn-hw-toggle")
    _draw_line(page)
    page.wait_for_function("(window.__hw['s1.png'] ?? []).length === 1")
    page.click("#btn-hw-undo")
    page.wait_for_function("(window.__hw['s1.png'] ?? []).length === 0")
    page.click("#btn-hw-redo")
    page.wait_for_function("(window.__hw['s1.png'] ?? []).length === 1")


def test_strokes_are_per_sheet(open_app):
    page = open_app(API)
    _open_annotate(page)
    page.click("#btn-hw-toggle")
    _draw_line(page)
    page.wait_for_function("(window.__hw['s1.png'] ?? []).length === 1")
    page.keyboard.press("ArrowRight")
    page.wait_for_function(
        "document.querySelector('#single-sheet-name').textContent.includes('s2.png')")
    # 次の答案では空から始まり、書けば別キーに保存される
    _draw_line(page)
    page.wait_for_function("(window.__hw['s2.png'] ?? []).length === 1")
    assert page.evaluate("window.__hw['s1.png'].length") == 1


def test_c_key_toggles_and_close_returns_main(open_app):
    page = open_app(API)
    _open_annotate(page)
    page.keyboard.press("c")
    page.wait_for_selector("#btn-hw-toggle.active")
    page.keyboard.press("c")
    page.wait_for_selector("#btn-hw-toggle:not(.active)")
    page.click("#btn-single-close")
    page.wait_for_selector("#sheet-list-view", state="visible")
    page.click("#btn-sheet-list-close")
    page.wait_for_selector("#sheet-list-view", state="hidden")
    assert page.locator("#datasource-panel").is_visible()


def test_eraser_drag_erases_and_undo_restores(open_app):
    """S6: 消しゴムはドラッグ中も例外なく効き、Ctrl+Z で戻せる"""
    page = open_app(API)
    _open_annotate(page)
    errors = []
    page.on("pageerror", lambda e: errors.append(str(e)))
    page.click("#btn-hw-toggle")
    _draw_line(page)
    page.wait_for_function("(window.__hw['s1.png'] ?? []).length === 1")
    page.click("#btn-hw-eraser")
    # ストロークの上をドラッグで通過して消す（押下点は線から外し、move 中に消させる）
    box = page.locator("#hw-canvas").bounding_box()
    y = box["y"] + box["height"] * 0.3 + 5
    page.mouse.move(box["x"] + box["width"] * 0.2, y - 60)
    page.mouse.down()
    page.mouse.move(box["x"] + box["width"] * 0.35, y, steps=12)
    page.mouse.up()
    page.wait_for_function("(window.__hw['s1.png'] ?? []).length === 0")
    assert errors == [], f"消しゴム中に例外: {errors}"
    page.keyboard.press("Control+z")
    page.wait_for_function("(window.__hw['s1.png'] ?? []).length === 1")
