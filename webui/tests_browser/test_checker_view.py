"""マークチェックビューのUI振る舞い（L2: モックbridge）。"""

TINY_PNG = ("data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ"
            "AAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==")

API = """({
  ping: async () => ({ok: true}),
  set_mode: async () => ({ok: true, state: (await window.__mockApi.get_state()).state}),
  get_app_info: async () => ({ok: true, app_version: 't', webui_version: 't'}),
  get_state: async () => ({ok: true, state: {
    app_mode: 'mark_only', mark_format: 'multi_digit', skip_questions: 0,
    image_folder: '/tmp/scans', image_count: 2,
    coord_file: '/tmp/c.xlsx', coord_summary: {answer_rows: 3, marks_per_row: 15, warning: null},
    answer_key: null, key_summary: null,
    omr_mode: 'kmeans', color_threshold: 0.1, area_threshold: 0.4,
    omr_result: '/tmp/r.xlsx', job: {running: false, kind: null, current: 0, total: 0},
    checker: null,
  }}),
  open_mark_checker: async () => ({ok: true, state: {checker: window.__checkerState}}),
  close_mark_checker: async () => ({ok: true, state: {checker: null}}),
  get_checker_entries: async (cat, page, size) => {
    const all = window.__entries;
    const items = cat === '__errors__' ? all.filter(e => e.error_type)
      : cat ? all.filter(e => e.category === cat) : all;
    return {ok: true, items: items.slice(page*size, (page+1)*size), total: items.length,
            page, page_size: size};
  },
  get_entry_image: async () => ({ok: true, data_url: '%(png)s'}),
  set_correction: async (id, v) => {
    if (v && !'-0123456789abcd'.includes(v.toLowerCase())) {
      return {ok: false, error: 'マーク記号1文字（- 0〜9 a〜d)または -1 を入力してください'};
    }
    const e = window.__entries.find(x => x.id === id);
    e.after = v.toLowerCase();
    window.__checkerState.corrected = window.__entries.filter(x => x.after).length;
    return {ok: true, state: {checker: window.__checkerState},
            entry: {id, before: e.before, after: e.after, category: e.category}};
  },
  apply_corrections: async () => {
    window.__entries.forEach(e => { if (e.after) { e.before = e.after; e.after = ''; e.error_type = ''; e.category = e.before; } });
    window.__checkerState = Object.assign({}, window.__checkerState,
      {corrected: 0, error_count: 0});
    return {ok: true, applied: 2, backup: '/tmp/backup/x.xlsx',
            state: {checker: window.__checkerState}};
  },
  select_image_folder: async () => ({ok: true, cancelled: true}),
  select_coord_file: async () => ({ok: true, cancelled: true}),
  select_answer_key: async () => ({ok: true, cancelled: true}),
  select_omr_result: async () => ({ok: true, cancelled: true}),
  set_skip_questions: async () => ({ok: true, cancelled: true}),
  set_omr_mode: async () => ({ok: true, cancelled: true}),
  set_thresholds: async () => ({ok: true, cancelled: true}),
  run_recognition: async () => ({ok: true, cancelled: true}),
  cancel_job: async () => ({ok: true, cancelled: true}),
})""" % {"png": TINY_PNG}

SEED = """
window.__entries = [
  {id: 0, filename: 's1.png', question_no: 1, before: '', after: '',
   category: 'ノーマーク', error_type: 'NoMark'},
  {id: 1, filename: 's1.png', question_no: 2, before: '2;8', after: '',
   category: '複数マーク', error_type: 'DoubleMark'},
  {id: 2, filename: 's2.png', question_no: 1, before: '-', after: '',
   category: '-', error_type: ''},
];
window.__checkerState = {open: true, total: 3, corrected: 0, error_count: 2,
  categories: [
    {name: 'ノーマーク', count: 1, is_error: true},
    {name: '複数マーク', count: 1, is_error: true},
    {name: '-', count: 1, is_error: false},
  ]};
"""


from conftest import enter_mode


def _open_checker(open_app):
    page = open_app(API)
    page.evaluate(SEED)
    enter_mode(page, ".mode-card.math")
    page.click("#btn-open-checker")
    page.wait_for_selector("#checker-view:not([hidden])")
    return page


def test_opens_with_error_tab_active(open_app):
    page = _open_checker(open_app)
    assert "要確認 (2)" in page.locator(".tab.active").inner_text()
    assert page.locator(".entry-card").count() == 2   # エラー2件のみ表示
    assert page.locator("main > .panel").first.is_hidden()


def test_correction_input_validates_and_marks_card(open_app):
    page = _open_checker(open_app)
    card = page.locator(".entry-card").first
    inp = card.locator("input")
    inp.fill("x")
    inp.dispatch_event("change")
    page.wait_for_function("document.querySelector('#log').textContent.includes('❌')")

    inp.fill("A")
    inp.dispatch_event("change")
    page.wait_for_selector(".entry-card.corrected")
    assert "訂正 1 件" in page.locator("#checker-summary").inner_text()
    assert not page.locator("#btn-apply-corrections").is_disabled()


def test_apply_flow_updates_summary_and_log(open_app):
    page = _open_checker(open_app)
    inp = page.locator(".entry-card").first.locator("input")
    inp.fill("9")
    inp.dispatch_event("change")
    page.wait_for_selector(".entry-card.corrected")
    page.click("#btn-apply-corrections")
    page.wait_for_function("document.querySelector('#log').textContent.includes('反映しました')")
    assert "要確認 0 件" in page.locator("#checker-summary").inner_text()


def test_close_restores_panels(open_app):
    page = _open_checker(open_app)
    page.click("#btn-close-checker")
    page.wait_for_selector("#checker-view", state="hidden")
    assert page.locator("main > .panel").first.is_visible()


def test_keyboard_correction_advances_cursor(open_app):
    page = _open_checker(open_app)
    # カーソルは先頭カード。記号キー1打で訂正が入り、次のカードへ進む
    page.wait_for_selector(".entry-card.cursor[data-index='0']")
    page.keyboard.press("a")
    page.wait_for_selector(".entry-card.corrected")
    page.wait_for_selector(".entry-card.cursor[data-index='1']")
    assert "訂正 1 件" in page.locator("#checker-summary").inner_text()
    # BackSpace で訂正取消（カーソルは動かない）
    page.keyboard.press("ArrowLeft")
    page.wait_for_selector(".entry-card.cursor[data-index='0']")
    page.keyboard.press("Backspace")
    page.wait_for_function(
        "document.querySelector('#checker-summary').textContent.includes('訂正 0 件')")
    # Enter で入力欄にフォーカス（-1 などの特殊値用）
    page.keyboard.press("Enter")
    page.wait_for_function(
        "document.activeElement.tagName === 'INPUT'")
