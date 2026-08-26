"""Step1（認識実行）パネルのUI振る舞い（L2: モックbridge＋イベント注入）。"""

READY_STATE = """{
  app_mode: 'mark_only', mark_format: 'standard', skip_questions: 4,
  image_folder: '/tmp/scans', image_count: 2,
  coord_file: '/tmp/c.xlsx',
  coord_summary: {answer_rows: 8, marks_per_row: 10, warning: null},
  answer_key: null, key_summary: null,
  omr_mode: 'kmeans', color_threshold: 0.1, area_threshold: 0.4,
  omr_result: null, job: {running: false, kind: null, current: 0, total: 0},
}"""

API = """({
  ping: async () => ({ok: true}),
  set_mode: async () => ({ok: true, state: (await window.__mockApi.get_state()).state}),
  get_app_info: async () => ({ok: true, app_version: 't', webui_version: 't'}),
  get_state: async () => ({ok: true, state: window.__state ?? %(ready)s}),
  run_recognition: async () => {
    window.__state = Object.assign(%(ready)s, {job: {running: true, kind: 'recognition', current: 0, total: 2}});
    return {ok: true, state: window.__state};
  },
  cancel_job: async () => ({ok: true, state: window.__state}),
  set_omr_mode: async (m) => ({ok: true, state: Object.assign(%(ready)s, {omr_mode: m})}),
  set_thresholds: async () => ({ok: true, state: %(ready)s}),
  select_omr_result: async () => ({ok: true, cancelled: true}),
  select_image_folder: async () => ({ok: true, cancelled: true}),
  select_coord_file: async () => ({ok: true, cancelled: true}),
  select_answer_key: async () => ({ok: true, cancelled: true}),
  set_skip_questions: async () => ({ok: true, state: %(ready)s}),
})""" % {"ready": READY_STATE}


from conftest import enter_mode


def _boot(open_app):
    page = open_app(API)
    enter_mode(page)
    return page


def test_run_button_enabled_when_ready(open_app):
    page = _boot(open_app)
    assert not page.locator("#btn-run-recognition").is_disabled()
    assert page.locator("#threshold-inputs").is_hidden()  # kmeans では非表示


def test_threshold_inputs_appear_in_threshold_mode(open_app):
    page = _boot(open_app)
    page.select_option("#omr-mode", "threshold")
    page.wait_for_selector("#threshold-inputs:not([hidden])")


def test_run_shows_progress_then_done(open_app):
    page = _boot(open_app)
    page.click("#btn-run-recognition")
    # 実行中: ボタン無効・中断表示
    page.wait_for_selector("#btn-cancel:not([hidden])")
    assert page.locator("#btn-run-recognition").is_disabled()

    # Python からの進捗 push を模擬
    page.evaluate("window.saitenEvents({type:'progress', kind:'recognition', current:1, total:2})")
    assert page.locator("#job-status").inner_text() == "1 / 2"

    # 完了 push（get_state が返す state を完了後に差し替え）
    page.evaluate("""
      window.__state = Object.assign(%s, {omr_result: '/tmp/x/Mark2-Result-A0.40-C0.10-x.xlsx'});
      window.saitenEvents({type:'job_done', kind:'recognition', ok:true,
                          message:'認識完了: 成功 2 件 / エラー 0 件'});
    """ % READY_STATE)
    page.wait_for_function("document.querySelector('#log').textContent.includes('認識完了')")
    page.wait_for_selector("#btn-cancel[hidden]", state="attached")
    assert "Mark2-Result-" in page.locator("#row-omr-result .ds-value").inner_text()
