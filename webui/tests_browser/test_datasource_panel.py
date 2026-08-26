"""データソースパネルのUI振る舞い（L2: モックbridge）。"""

BASE_STATE = """{
  app_mode: 'mark_only', mark_format: 'standard', skip_questions: 4,
  image_folder: null, image_count: 0,
  coord_file: null, coord_summary: null,
  answer_key: null, key_summary: null,
  omr_mode: 'kmeans', color_threshold: 0.1, area_threshold: 0.4,
  omr_result: null, job: {running: false, kind: null, current: 0, total: 0},
}"""

API_TEMPLATE = """({
  ping: async () => ({ok: true}),
  set_mode: async () => ({ok: true, state: (await window.__mockApi.get_state()).state}),
  get_app_info: async () => ({ok: true, app_version: 't', webui_version: 't'}),
  get_state: async () => ({ok: true, state: %(base)s}),
  set_skip_questions: async (n) => (%(skip_impl)s),
  select_image_folder: async () => (%(folder_impl)s),
  select_coord_file: async () => (%(coord_impl)s),
  select_answer_key: async () => ({ok: true, cancelled: true}),
})"""


def make_api(folder_impl="{ok: true, cancelled: true}",
             coord_impl="{ok: true, cancelled: true}",
             skip_impl="{ok: true, state: Object.assign(%s, {skip_questions: Number(n)})}" % BASE_STATE):
    return API_TEMPLATE % {"base": BASE_STATE, "folder_impl": folder_impl,
                           "coord_impl": coord_impl, "skip_impl": skip_impl}


from conftest import enter_mode


def test_initial_state_shows_placeholders(open_app):
    page = open_app(make_api())
    enter_mode(page)
    # 未選択プレースホルダ（CSSの ::before で表示）
    assert page.locator("#row-image-folder .ds-value").inner_text() == ""
    assert int(page.locator("#skip-input").input_value()) == 4


def test_folder_selection_updates_row_and_log(open_app):
    folder_impl = ("{ok: true, state: Object.assign(%s, "
                   "{image_folder: '/tmp/scans', image_count: 32})}") % BASE_STATE
    page = open_app(make_api(folder_impl=folder_impl))
    enter_mode(page)
    page.click("[data-action='select_image_folder']")
    page.wait_for_selector("#summary-image-folder:not([hidden])")
    assert "画像 32 枚" in page.locator("#summary-image-folder").inner_text()
    assert "/tmp/scans" in page.locator("#row-image-folder .ds-value").inner_text()
    assert "画像フォルダを選択しました" in page.locator("#log").inner_text()


def test_coord_warning_is_visible(open_app):
    coord_impl = ("{ok: true, state: Object.assign(%s, {coord_file: '/tmp/c.xlsx', "
                  "coord_summary: {answer_rows: 8, marks_per_row: 10, "
                  "warning: '数学マーク採点（複数桁）モードですが標準テンプレート相当です'}})}"
                  ) % BASE_STATE
    page = open_app(make_api(coord_impl=coord_impl))
    enter_mode(page)
    page.click("[data-action='select_coord_file']")
    summary = page.locator("#summary-coord-file")
    page.wait_for_selector("#summary-coord-file:not([hidden])")
    assert "⚠" in summary.inner_text()
    assert "warn" in summary.get_attribute("class")


def test_error_is_logged_and_state_kept(open_app):
    folder_impl = "{ok: false, error: '選択したフォルダに画像（jpg/png）がありません'}"
    page = open_app(make_api(folder_impl=folder_impl))
    enter_mode(page)
    page.on("dialog", lambda d: d.accept())
    page.click("[data-action='select_image_folder']")
    page.wait_for_function("document.querySelector('#log').textContent.includes('❌')")
    assert "画像（jpg/png）がありません" in page.locator("#log").inner_text()
    # 行は未選択のまま
    assert page.locator("#row-image-folder .ds-value").inner_text() == ""
