"""殻（index.html + bridge.js + main.js）の振る舞いを固定する。"""


OK_API = """({
  ping: async () => ({ok: true, pong: true}),
  get_app_info: async () => ({ok: true, app_version: '9.9.9-test', webui_version: '0.0.t'}),
  get_state: async () => ({ok: true, state: {
    app_mode: 'mark_only', mark_format: 'standard', skip_questions: 4,
    image_folder: null, image_count: 0, coord_file: null, coord_summary: null,
    answer_key: null, key_summary: null,
  }}),
})"""

FAIL_API = """({
  ping: async () => ({ok: false, error: 'ブリッジ初期化に失敗しました'}),
  get_app_info: async () => ({ok: true, app_version: 'x', webui_version: 'x'}),
})"""


def test_shell_shows_version_and_connected(open_app):
    page = open_app(OK_API)
    status = page.locator("#bridge-status")
    status.wait_for(state="visible")
    page.wait_for_function("document.querySelector('#bridge-status').dataset.state === 'ok'")
    assert status.inner_text() == "接続済み"
    assert "9.9.9-test" in page.locator("#version").inner_text()
    assert "webui 0.0.t" in page.locator("#version").inner_text()


def test_shell_reports_bridge_error(open_app):
    """{ok:false} は握りつぶされず、接続エラーとして利用者に見える"""
    page = open_app(FAIL_API)
    page.wait_for_function("document.querySelector('#bridge-status').dataset.state === 'error'")
    assert "ブリッジ初期化に失敗しました" in page.locator("#bridge-status").inner_text()
