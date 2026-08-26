"""記述採点3ビュー（設定/グリッド/一枚採点）のUI振る舞い（L2: モックbridge）。"""

TINY_PNG = ("data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ"
            "AAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==")

# 可変モック: window.__desc に問題定義と得点を持たせ、set_descriptive_score で更新する
API = """(() => {
  window.__desc = {
    questions: [
      {id: 'D1', name: '問1', max_score: 5, aspect: 1, region: [10, 10, 80, 40]},
      {id: 'D2', name: '問2', max_score: 3, aspect: 1, region: [10, 50, 80, 90]},
    ],
    scores: {'s1.png': {}, 's2.png': {}},
  };
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
    image_folder: '/tmp/scans', image_count: 2,
    coord_file: null, coord_summary: null, answer_key: null, key_summary: null,
    omr_mode: 'kmeans', color_threshold: 0.1, area_threshold: 0.4,
    omr_result: null, job: {running: false, kind: null, current: 0, total: 0},
    checker: null,
    descriptive: {questions: window.__desc.questions,
                  scored_counts: scoredCounts(), prepared_count: 2},
  });
  return {
    ping: async () => ({ok: true}),
    set_mode: async () => ({ok: true, state: state()}),
    get_app_info: async () => ({ok: true, app_version: 't', webui_version: 't'}),
    get_state: async () => ({ok: true, state: state()}),
    get_sheet_image: async (f) => ({ok: true, filename: f ?? 's1.png', data_url: '%(png)s'}),
    list_sheet_files: async () => ({ok: true, files: ['s1.png', 's2.png']}),
    start_descriptive_scoring: async () => ({ok: true, crops: {}}),
    list_descriptive_targets: async (qid) => ({ok: true, items:
      ['s1.png', 's2.png'].map(f => ({filename: f,
        score: window.__desc.scores[f][qid] ?? null}))}),
    get_descriptive_crop: async () => ({ok: true, data_url: '%(png)s'}),
    set_descriptive_score: async (f, qid, v) => {
      if (v !== null) {
        const q = window.__desc.questions.find(x => x.id === qid);
        if (v < 0 || v > q.max_score) return {ok: false, error: '範囲外の得点です'};
      }
      if (v === null) delete window.__desc.scores[f][qid];
      else window.__desc.scores[f][qid] = v;
      return {ok: true, state: state()};
    },
    delete_descriptive_question: async (qid) => {
      window.__desc.questions = window.__desc.questions.filter(q => q.id !== qid);
      for (const f in window.__desc.scores) delete window.__desc.scores[f][qid];
      return {ok: true, state: state()};
    },
    select_image_folder: async () => ({ok: true, cancelled: true}),
  };
})()""" % {"png": TINY_PNG}

from conftest import enter_mode

DESC_CARD = ".mode-card[data-mode=descriptive_only]"


def test_config_view_lists_questions_and_deletes(open_app):
    page = open_app(API)
    enter_mode(page, DESC_CARD)
    page.click("#btn-desc-config")
    page.wait_for_selector("#desc-config-view", state="visible")
    rows = page.locator("#desc-table tbody tr")
    assert rows.count() == 2
    assert "問1" in rows.nth(0).inner_text()
    assert "2 問 / 対象画像 2 枚" in page.locator("#desc-config-summary").inner_text()
    # 削除（confirm は Playwright 既定で dismiss → accept を設定）
    page.on("dialog", lambda d: d.accept())
    rows.nth(1).locator("button", has_text="削除").click()
    page.wait_for_function(
        "document.querySelectorAll('#desc-table tbody tr').length === 1")
    # 閉じるとメインに戻る
    page.click("#btn-desc-config-close")
    # View Transition は1フレーム遅れて反映されるため待機型で確認
    page.wait_for_selector("#desc-config-view", state="hidden")
    assert page.locator("#datasource-panel").is_visible()


def test_grid_view_scores_and_updates_tabs(open_app):
    page = open_app(API)
    enter_mode(page, DESC_CARD)
    page.click("#btn-desc-scoring")
    page.wait_for_selector("#desc-scoring-view", state="visible")
    tabs = page.locator("#desc-q-tabs .tab")
    assert tabs.count() == 2
    assert "問1 (0/2)" in tabs.nth(0).inner_text()
    cards = page.locator("#desc-grid .entry-card")
    assert cards.count() == 2
    # s1.png の問1に 4 点を付ける → タブとカードが更新される
    cards.nth(0).locator(".score-btn", has_text="4").first.click()
    page.wait_for_function(
        "document.querySelector('#desc-q-tabs .tab').textContent.includes('(1/2)')")
    assert "4 点" in page.locator("#desc-grid .entry-card").nth(0).inner_text()
    assert "採点済み 1 / 4" in page.locator("#desc-scoring-summary").inner_text()
    # 「未」で未採点に戻せる
    page.locator("#desc-grid .entry-card").nth(0).locator(".score-btn.clear").click()
    page.wait_for_function(
        "document.querySelector('#desc-q-tabs .tab').textContent.includes('(0/2)')")


def test_single_sheet_view_overlays_and_navigation(open_app):
    page = open_app(API)
    enter_mode(page, DESC_CARD)
    page.click("#btn-desc-scoring")
    page.wait_for_selector("#desc-scoring-view", state="visible")
    page.click("#btn-single-sheet")
    page.wait_for_selector("#single-sheet-view", state="visible")
    assert "s1.png（1 / 2）" in page.locator("#single-sheet-name").inner_text()
    assert page.locator("#btn-sheet-prev").is_disabled()
    # 領域オーバーレイが問題数ぶん載る
    assert page.locator("#annotation-layer .region-box").count() == 2
    # サイドパネルで採点 → ラベルに点数が反映される
    side1 = page.locator("#sheet-side .side-q").nth(0)
    side1.locator(".score-btn", has_text="5").first.click()
    page.wait_for_function(
        "document.querySelector('#annotation-layer .region-label').textContent.includes('5点')")
    assert "5点" in page.locator("#sheet-side .side-q .score-mark").first.inner_text()
    # 次の答案へ → ファイル名が変わり、得点は答案ごとに独立
    page.click("#btn-sheet-next")
    page.wait_for_function(
        "document.querySelector('#single-sheet-name').textContent.includes('s2.png')")
    assert page.locator("#btn-sheet-next").is_disabled()
    assert "未" in page.locator("#annotation-layer .region-label").first.inner_text()
    # 問題別に戻る
    page.click("#btn-single-close")
    page.wait_for_selector("#desc-scoring-view", state="visible")
