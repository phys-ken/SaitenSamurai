"""モード選択画面の「採点再開（セッション復元）」ボタン（L2）。"""

API = """({
  ping: async () => ({ok: true}),
  get_app_info: async () => ({ok: true, app_version: 't', webui_version: 't'}),
  get_state: async () => ({ok: true, state: {
    app_mode: 'mark_only', mark_format: 'multi_digit', skip_questions: 0,
    image_folder: null, image_count: 0,
    coord_file: null, coord_summary: null, answer_key: null, key_summary: null,
    omr_mode: 'kmeans', color_threshold: 0.1, area_threshold: 0.4,
    omr_result: null, job: {running: false, kind: null, current: 0, total: 0},
    checker: null,
  }}),
  restore_session: async () => ({ok: true, warnings: ['座標ファイルが見つかりません（c.xlsx）。選び直してください'],
    state: {
      app_mode: 'mark_only', mark_format: 'multi_digit', skip_questions: 0,
      image_folder: '/tmp/scans', image_count: 3,
      coord_file: null, coord_summary: null,
      answer_key: '/tmp/key.xlsx',
      key_summary: {ok: true, errors: [], warnings: [], stats_line: '4問 / 5点'},
      omr_mode: 'kmeans', color_threshold: 0.1, area_threshold: 0.4,
      omr_result: '/tmp/r.xlsx', job: {running: false, kind: null, current: 0, total: 0},
      checker: null,
    }}),
})"""


def test_resume_button_restores_and_enters_main(open_app):
    page = open_app(API)
    page.wait_for_function(
        "document.querySelector('#bridge-status').dataset.state === 'ok'")
    btn = page.locator("#btn-resume-session")
    assert btn.is_visible()   # モード選択画面に出ている
    btn.click()
    page.wait_for_selector("#mode-select", state="hidden")
    # 復元された state が描画されている
    assert "/tmp/scans" in page.locator("#row-image-folder .ds-value").inner_text()
    assert "数学マーク採点" in page.locator("#mode-badge").inner_text()
    log = page.locator("#log").inner_text()
    assert "セッションを復元しました" in log
    assert "座標ファイルが見つかりません" in log   # warnings がログに出る


CANCEL_API = """({
  ping: async () => ({ok: true}),
  get_app_info: async () => ({ok: true, app_version: 't', webui_version: 't'}),
  get_state: async () => ({ok: true, state: {
    app_mode: 'mark_only', mark_format: 'standard', skip_questions: 4,
    image_folder: null, image_count: 0,
    coord_file: null, coord_summary: null, answer_key: null, key_summary: null,
    omr_mode: 'kmeans', color_threshold: 0.1, area_threshold: 0.4,
    omr_result: null, job: {running: false, kind: null, current: 0, total: 0},
    checker: null,
  }}),
  restore_session: async () => ({ok: true, cancelled: true, state: null}),
})"""


def test_resume_cancel_stays_on_mode_select(open_app):
    page = open_app(CANCEL_API)
    page.wait_for_function(
        "document.querySelector('#bridge-status').dataset.state === 'ok'")
    page.click("#btn-resume-session")
    page.wait_for_timeout(300)
    assert page.locator("#mode-select").is_visible()
