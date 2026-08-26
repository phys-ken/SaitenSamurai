"""ステッパー・ℹポップオーバー・用語とファイルの地図（L2）。"""

API = """(() => {
  window.__progress = {prepared: true, read: true, scored: false, summarized: false};
  const state = () => ({
    app_mode: 'mark_only', mark_format: 'multi_digit', skip_questions: 0,
    image_folder: '/tmp/scans', image_count: 3,
    coord_file: '/tmp/c.xlsx', coord_summary: {answer_rows: 3, marks_per_row: 15, warning: null},
    answer_key: null, key_summary: null,
    omr_mode: 'kmeans', color_threshold: 0.1, area_threshold: 0.4,
    omr_result: '/tmp/r.xlsx', job: {running: false, kind: null, current: 0, total: 0},
    checker: null,
  });
  return {
    ping: async () => ({ok: true}),
    set_mode: async () => ({ok: true, state: state()}),
    get_app_info: async () => ({ok: true, app_version: 't', webui_version: 't'}),
    get_state: async () => ({ok: true, state: state()}),
    get_progress: async () => ({ok: true, progress: window.__progress}),
    select_image_folder: async () => ({ok: true, cancelled: true}),
  };
})()"""

from conftest import enter_mode


def test_stepper_marks_done_and_current(open_app):
    page = open_app(API)
    enter_mode(page)
    page.wait_for_selector("#stepper:not([hidden])")
    page.wait_for_selector("#stepper li[data-step='prepared'].done")
    assert page.locator("#stepper li[data-step='read']").get_attribute("class") \
        .find("done") >= 0
    # 現在地は「採点」に朱
    page.wait_for_selector("#stepper li[data-step='scored'].current")
    assert "current" not in (
        page.locator("#stepper li[data-step='summarized']").get_attribute("class") or "")
    # 進行が進むと表示も進む
    page.evaluate("window.__progress.scored = true")
    page.evaluate("window.dispatchEvent(new Event('noop'))")
    page.click("#stepper li[data-step='prepared']")   # クリック→render は走らないため
    # get_state 経由で再描画させる
    page.evaluate("window.__refreshState()")
    page.wait_for_selector("#stepper li[data-step='scored'].done")
    page.wait_for_selector("#stepper li[data-step='summarized'].current")


def test_info_popover_opens_and_links_to_map(open_app):
    page = open_app(API)
    enter_mode(page)
    page.hover(".info-btn[data-info='coord-file']")
    page.wait_for_selector(".info-popover:not([hidden])")
    text = page.locator(".info-popover").inner_text()
    assert "座標ファイル" in text and "mark_areas.xlsx" in text
    # リンクから地図が開く
    page.click(".info-popover a")
    page.wait_for_selector("#map-overlay:not([hidden])")
    assert "00_Processing" in page.locator(".map-body").inner_text()
    page.click("#btn-map-close")
    page.wait_for_selector("#map-overlay", state="hidden")


def test_map_button_in_topbar(open_app):
    page = open_app(API)
    page.wait_for_function(
        "document.querySelector('#bridge-status').dataset.state === 'ok'")
    page.click("#btn-open-map")
    page.wait_for_selector("#map-overlay:not([hidden])")
    assert "消してよい" in page.locator(".map-body").inner_text()
