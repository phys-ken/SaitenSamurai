"""記述のみモードでマーク系UIが隠れること（L2）。"""
DESC_STATE = """{
  app_mode: 'descriptive_only', mark_format: 'standard', skip_questions: 4,
  image_folder: '/tmp/scans', image_count: 3,
  coord_file: null, coord_summary: null,
  answer_key: null, key_summary: null,
  omr_mode: 'kmeans', color_threshold: 0.1, area_threshold: 0.4,
  omr_result: null, job: {running: false, kind: null, current: 0, total: 0},
  checker: null,
  descriptive: {questions: [], scored_counts: {}, prepared_count: 3},
}"""

API = """({
  ping: async () => ({ok: true}),
  set_mode: async () => ({ok: true, state: %(s)s}),
  get_app_info: async () => ({ok: true, app_version: 't', webui_version: 't'}),
  get_state: async () => ({ok: true, state: %(s)s}),
})""" % {"s": DESC_STATE}


def test_desc_only_hides_mark_rows(open_app):
    page = open_app(API)
    errors = []
    page.on("console", lambda m: errors.append(m.text) if m.type == "error" else None)
    page.on("pageerror", lambda e: errors.append(str(e)))
    page.wait_for_function("document.querySelector('#bridge-status').dataset.state === 'ok'")
    page.click(".mode-card[data-mode=descriptive_only]")
    page.wait_for_selector("#mode-select", state="hidden")
    if errors:
        raise AssertionError("JS errors: " + " | ".join(errors))
    assert page.locator("#row-coord-file").is_hidden()
    assert page.locator("#row-answer-key").is_hidden()
    assert page.locator("#row-omr-result").is_hidden()
    assert page.locator("#btn-run-recognition").inner_text() == "▶ 画像準備"
    assert page.locator(".desc-only-ui").first.is_visible()
