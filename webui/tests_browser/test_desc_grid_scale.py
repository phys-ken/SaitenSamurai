"""記述グリッドの400人スケール動作（L2）: 仮想スクロール・キー採点・自動送り。"""

TINY_PNG = ("data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ"
            "AAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==")

N = 400

API = """(() => {
  const N = %(n)d;
  window.__desc = {
    questions: [
      {id: 'D1', name: '問1', max_score: 5, aspect: 1, region: [10, 10, 200, 80]},
      {id: 'D2', name: '問2', max_score: 12, aspect: 1, region: [10, 90, 200, 160]},
    ],
    scores: {},
  };
  const files = Array.from({length: N}, (_, i) =>
    's' + String(i + 1).padStart(3, '0') + '.png');
  files.forEach(f => { window.__desc.scores[f] = {}; });
  window.__cropCalls = 0;
  const scoredCounts = () => {
    const c = {};
    for (const q of window.__desc.questions) {
      c[q.id] = Object.values(window.__desc.scores)
        .filter(s => s[q.id] !== undefined && s[q.id] !== null).length;
    }
    return c;
  };
  const state = () => ({
    app_mode: 'descriptive_only', mark_format: 'standard', skip_questions: 4,
    image_folder: '/tmp/scans', image_count: N,
    coord_file: null, coord_summary: null, answer_key: null, key_summary: null,
    omr_mode: 'kmeans', color_threshold: 0.1, area_threshold: 0.4,
    omr_result: null, job: {running: false, kind: null, current: 0, total: 0},
    checker: null,
    descriptive: {questions: window.__desc.questions,
                  scored_counts: scoredCounts(), prepared_count: N},
  });
  return {
    ping: async () => ({ok: true}),
    set_mode: async () => ({ok: true, state: state()}),
    get_app_info: async () => ({ok: true, app_version: 't', webui_version: 't'}),
    get_state: async () => ({ok: true, state: state()}),
    start_descriptive_scoring: async () => ({ok: true, crops: {}}),
    list_descriptive_targets: async (qid) => ({ok: true, items:
      files.map(f => ({filename: f, score: window.__desc.scores[f][qid] ?? null}))}),
    get_descriptive_crop: async () => {
      window.__cropCalls++;
      return {ok: true, data_url: '%(png)s'};
    },
    set_descriptive_score: async (f, qid, v) => {
      if (v === null) delete window.__desc.scores[f][qid];
      else window.__desc.scores[f][qid] = v;
      return {ok: true, state: state()};
    },
    select_image_folder: async () => ({ok: true, cancelled: true}),
  };
})()""" % {"n": N, "png": TINY_PNG}

from conftest import enter_mode

DESC_CARD = ".mode-card[data-mode=descriptive_only]"


def _open_grid(page):
    enter_mode(page, DESC_CARD)
    page.click("#btn-desc-scoring")
    page.wait_for_selector("#desc-scoring-view", state="visible")
    page.wait_for_selector("#desc-grid .entry-card")


def test_virtual_scroll_mounts_subset_of_400(open_app):
    page = open_app(API)
    _open_grid(page)
    mounted = page.locator("#desc-grid .entry-card").count()
    assert 0 < mounted < 100, f"400件全部DOMに載っている（{mounted}件）"
    # 画像取得も表示分＋先読みに留まる
    calls = page.evaluate("window.__cropCalls")
    assert calls < 100, f"クロップを{calls}回取得（遅延取得が働いていない）"
    # 末尾までスクロールすると最後のカードが現れ、DOM数は増えない
    page.evaluate(
        "const v = document.querySelector('#desc-grid'); v.scrollTop = v.scrollHeight")
    page.wait_for_selector("#desc-grid .entry-card[data-index='399']")
    assert page.locator("#desc-grid .entry-card").count() < 100


def test_keyboard_scoring_auto_advances(open_app):
    page = open_app(API)
    _open_grid(page)
    # カーソルは先頭。数字キーで採点 → 次の未採点（s002）へ朱枠が移る
    page.keyboard.press("4")
    page.wait_for_function("window.__desc.scores['s001.png'].D1 === 4")
    page.wait_for_selector("#desc-grid .entry-card.cursor[data-index='1']")
    page.keyboard.press("0")
    page.wait_for_function("window.__desc.scores['s002.png'].D1 === 0")
    page.wait_for_selector("#desc-grid .entry-card.cursor[data-index='2']")
    # BackSpace で未採点に戻す（カーソルは進まない）
    page.keyboard.press("ArrowLeft")
    page.wait_for_selector("#desc-grid .entry-card.cursor[data-index='1']")
    page.keyboard.press("Backspace")
    page.wait_for_function("window.__desc.scores['s002.png'].D1 === undefined")


def test_two_digit_scores_for_high_max(open_app):
    page = open_app(API)
    _open_grid(page)
    # 問2（満点12）タブへ
    page.locator("#desc-q-tabs .tab", has_text="問2").click()
    page.wait_for_function(
        "document.querySelector('#desc-q-tabs .tab.active').textContent.includes('問2')")
    page.wait_for_selector("#desc-grid .entry-card")
    page.keyboard.press("1")
    page.keyboard.press("2")   # 「12」→ 満点なので即確定
    page.wait_for_function("window.__desc.scores['s001.png'].D2 === 12")
    # 1桁のまま止めた場合はタイムアウト確定（500ms）
    page.keyboard.press("7")
    page.wait_for_function("window.__desc.scores['s002.png'].D2 === 7",
                           timeout=3000)


def test_jump_to_next_unscored(open_app):
    page = open_app(API)
    _open_grid(page)
    # 先頭2人をキーで採点（カーソルは自動で2へ）→ 先頭に戻ってからジャンプ
    page.keyboard.press("3")
    page.wait_for_function("window.__desc.scores['s001.png'].D1 === 3")
    page.keyboard.press("3")
    page.wait_for_function("window.__desc.scores['s002.png'].D1 === 3")
    page.keyboard.press("ArrowLeft")
    page.keyboard.press("ArrowLeft")
    page.wait_for_selector("#desc-grid .entry-card.cursor[data-index='0']")
    page.click("#btn-jump-unscored")
    page.wait_for_selector("#desc-grid .entry-card.cursor[data-index='2']")


# ── 一枚採点（キーボード動線・ズーム） ─────────────────────────
SHEET_API = API.replace(
    "select_image_folder: async () => ({ok: true, cancelled: true}),",
    """select_image_folder: async () => ({ok: true, cancelled: true}),
    list_sheet_files: async () => ({ok: true, files: files}),
    get_sheet_image: async (f) => ({ok: true, filename: f ?? files[0],
                                    data_url: '%(png)s'}),""" % {"png": TINY_PNG})


def _open_single(page):
    enter_mode(page, DESC_CARD)
    page.click("#btn-desc-scoring")
    page.wait_for_selector("#desc-scoring-view", state="visible")
    page.click("#btn-single-sheet")
    page.wait_for_selector("#single-sheet-view", state="visible")
    page.wait_for_selector("#annotation-layer .region-box")


def test_single_sheet_key_scoring_advances_to_next_sheet(open_app):
    page = open_app(SHEET_API)
    _open_single(page)
    assert "s001.png" in page.locator("#single-sheet-name").inner_text()
    # 問1に3点 → フォーカスが問2へ、問2に9点 → 全問済みで次の答案へ自動送り
    page.keyboard.press("3")
    page.wait_for_function("window.__desc.scores['s001.png'].D1 === 3")
    page.keyboard.press("9")
    page.wait_for_function("window.__desc.scores['s001.png'].D2 === 9")
    page.wait_for_function(
        "document.querySelector('#single-sheet-name').textContent.includes('s002.png')")
    assert "全問採点済み 1 枚" in page.locator("#single-sheet-name").inner_text()
    # ←で戻ると採点済みの朱枠とラベルが見える
    page.keyboard.press("ArrowLeft")
    page.wait_for_function(
        "document.querySelector('#single-sheet-name').textContent.includes('s001.png')")
    # ヘッダ更新→画像ロード→オーバーレイ敷設の順なので、敷設完了を待つ
    page.wait_for_function(
        "document.querySelectorAll('#annotation-layer .region-box.scored').length === 2")


def test_single_sheet_zoom_buttons(open_app):
    page = open_app(SHEET_API)
    _open_single(page)
    fit = page.evaluate("document.getElementById('sheet-image').clientWidth")
    page.click("#btn-zoom-100")
    page.wait_for_function(
        "document.getElementById('zoom-label').textContent === '100%'")
    natural = page.evaluate("document.getElementById('sheet-image').clientWidth")
    assert natural == page.evaluate(
        "document.getElementById('sheet-image').naturalWidth")
    page.click("#btn-zoom-fit")
    page.wait_for_function(
        f"document.getElementById('sheet-image').clientWidth === {fit}")


def test_jump_to_unfinished_sheet(open_app):
    page = open_app(SHEET_API)
    _open_single(page)
    # s001 を全問採点 → s002 に自動送りされた状態から、s001 に戻ってジャンプ
    page.keyboard.press("2")
    page.keyboard.press("8")
    page.wait_for_function(
        "document.querySelector('#single-sheet-name').textContent.includes('s002.png')")
    page.keyboard.press("ArrowLeft")
    page.wait_for_function(
        "document.querySelector('#single-sheet-name').textContent.includes('s001.png')")
    page.click("#btn-sheet-unfinished")
    page.wait_for_function(
        "document.querySelector('#single-sheet-name').textContent.includes('s002.png')")


def test_filter_and_sort_controls(open_app):
    page = open_app(API)
    _open_grid(page)
    # 2人採点してから「未採点」フィルタ → 表示件数が減る
    page.keyboard.press("5")
    page.wait_for_function("window.__desc.scores['s001.png'].D1 === 5")
    page.keyboard.press("0")
    page.wait_for_function("window.__desc.scores['s002.png'].D1 === 0")
    page.select_option("#desc-filter", "unscored")
    page.wait_for_function(
        "document.querySelector('#desc-grid .entry-card') &&"
        "document.querySelector('#desc-grid .entry-card').textContent.includes('s003.png')")
    # 「満点」フィルタは s001 だけ
    page.select_option("#desc-filter", "full")
    page.wait_for_function(
        "document.querySelectorAll('#desc-grid .entry-card').length === 1")
    assert "s001.png" in page.locator("#desc-grid .entry-card").inner_text()
    # 得点降順ソート（全て表示に戻して）
    page.select_option("#desc-filter", "all")
    page.select_option("#desc-sort", "score_desc")
    page.wait_for_function(
        "document.querySelector('#desc-grid .entry-card[data-index=\\'0\\']')"
        "?.textContent.includes('s001.png')")


def test_m_key_gives_full_score_and_space_skips(open_app):
    page = open_app(API)
    _open_grid(page)
    page.keyboard.press("m")
    page.wait_for_function("window.__desc.scores['s001.png'].D1 === 5")
    # Space は採点せずに次へ
    cur = page.locator("#desc-grid .entry-card.cursor")
    idx_before = cur.get_attribute("data-index")
    page.keyboard.press(" ")
    page.wait_for_function(
        f"document.querySelector('#desc-grid .entry-card.cursor').dataset.index !== '{idx_before}'")


def test_digit_buffer_discarded_on_tab_switch(open_app):
    """S7: 打ちかけの2桁バッファはタブ切替で破棄され、誤確定しない"""
    page = open_app(API)
    _open_grid(page)
    page.locator("#desc-q-tabs .tab", has_text="問2").click()   # 満点12
    page.wait_for_function(
        "document.querySelector('#desc-q-tabs .tab.active').textContent.includes('問2')")
    page.wait_for_selector("#desc-grid .entry-card")
    page.keyboard.press("1")                     # バッファに「1」
    page.locator("#desc-q-tabs .tab", has_text="問1").click()   # 500ms 以内に切替
    page.wait_for_function(
        "document.querySelector('#desc-q-tabs .tab.active').textContent.includes('問1')")
    page.wait_for_timeout(700)                   # 旧タイマーが生きていれば確定してしまう
    assert page.evaluate(
        "Object.values(window.__desc.scores).every(s => Object.keys(s).length === 0)"), \
        "打ちかけの得点がタブ切替後に確定された（S7）"
