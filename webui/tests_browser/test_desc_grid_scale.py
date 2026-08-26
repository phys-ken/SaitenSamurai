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
