"""表示項目の設定パネルと合計点ヒントの state 連動（L2）。"""

API = """(() => {
  window.__rs = {};
  window.__defaults = {
    mark_result_offset: 0.0, show_correct_answer: true, show_ox_mark: true,
    show_score: true, show_aspect: true, show_all_correct_star: true,
    mark_result_bg_white: false, total_show_max: true, total_show_aspects: true,
    descriptive_opacity: 0.5, descriptive_show_mark: true,
    descriptive_show_score: true, descriptive_show_aspect: true,
  };
  const state = () => ({
    app_mode: 'mark_only', mark_format: 'standard', skip_questions: 4,
    image_folder: '/tmp/scans', image_count: 3,
    coord_file: '/tmp/c.xlsx', coord_summary: {answer_rows: 3, marks_per_row: 15, warning: null},
    answer_key: null, key_summary: null,
    omr_mode: 'kmeans', color_threshold: 0.1, area_threshold: 0.4,
    omr_result: null, job: {running: false, kind: null, current: 0, total: 0},
    checker: null,
    total_display_region: [300, 700, 560, 800],
    rendering_settings: window.__rs,
  });
  return {
    ping: async () => ({ok: true}),
    set_mode: async () => ({ok: true, state: state()}),
    get_app_info: async () => ({ok: true, app_version: 't', webui_version: 't'}),
    get_state: async () => ({ok: true, state: state()}),
    get_rendering_settings: async () =>
      ({ok: true, settings: {...window.__defaults, ...window.__rs},
        defaults: window.__defaults}),
    set_rendering_settings: async (ov) => {
      Object.assign(window.__rs, ov);
      return {ok: true, state: state()};
    },
    select_image_folder: async () => ({ok: true, cancelled: true}),
  };
})()"""

from conftest import enter_mode


def test_total_hint_restored_from_state(open_app):
    """設定済みの合計点位置が、クリックしていなくてもヒントに出る（回帰）"""
    page = open_app(API)
    enter_mode(page)
    assert "設定済み (300, 700, 560, 800)" in \
        page.locator("#total-position-hint").inner_text()


def test_dialog_sections_follow_mode_and_saves_change(open_app):
    page = open_app(API)
    enter_mode(page)   # mark_only
    page.click("#btn-render-settings")
    page.wait_for_selector("#rs-overlay:not([hidden])")
    sections = page.locator(".rs-section h3")
    texts = " ".join(sections.nth(i).inner_text() for i in range(sections.count()))
    assert "合計点の表示" in texts and "マーク設問への描き込み" in texts
    assert "記述設問への描き込み" not in texts   # マークのみモードでは出さない

    # 「満点も表示する」を外す → bridge に保存される
    row = page.locator(".rs-row", has_text="満点も表示する")
    row.locator("input").uncheck()
    page.wait_for_function("window.__rs.total_show_max === false")
    page.click("#btn-rs-close")
    page.wait_for_selector("#rs-overlay", state="hidden")
