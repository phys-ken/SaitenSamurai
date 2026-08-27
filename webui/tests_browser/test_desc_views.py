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
    list_sheet_overview: async () => ({ok: true, items:
      ['s1.png', 's2.png'].map(f => ({filename: f,
        done: Object.keys(window.__desc.scores[f] ?? {}).length,
        total: window.__desc.questions.length, handwriting: false}))}),
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

from conftest import enter_mode, wait_transition

DESC_CARD = ".mode-card[data-mode=descriptive_only]"


def test_config_view_lists_questions_and_deletes(open_app):
    page = open_app(API)
    enter_mode(page, DESC_CARD)
    page.click("#btn-desc-config")
    page.wait_for_selector("#desc-config-view", state="visible")
    rows = page.locator("#desc-table tbody tr")
    assert rows.count() == 2
    # 名前・配点・観点はその場編集の input になっている
    assert rows.nth(0).locator("input.desc-edit").nth(0).input_value() == "問1"
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
    """得点パレット選択→カードクリックで付与（tk のアクティブ得点方式）"""
    page = open_app(API)
    enter_mode(page, DESC_CARD)
    page.click("#btn-desc-scoring")
    page.wait_for_selector("#desc-scoring-view", state="visible")
    wait_transition(page)
    tabs = page.locator("#desc-q-tabs .tab")
    assert tabs.count() == 2
    assert "問1 (0/2)" in tabs.nth(0).inner_text()
    cards = page.locator("#desc-grid .entry-card")
    assert cards.count() == 2
    # パレットで 4 を選び、カードをクリックして付与
    page.locator("#score-palette .palette-btn", has_text="4").first.click()
    cards.nth(0).click()
    page.wait_for_function(
        "document.querySelector('#desc-q-tabs .tab').textContent.includes('(1/2)')")
    first = page.locator("#desc-grid .entry-card[data-index='0']")
    assert "4点" in first.inner_text()
    assert "sc-partial" in (first.get_attribute("class") or "")   # 中間点=橙
    assert "採点済み 1 / 4" in page.locator("#desc-scoring-summary").inner_text()
    # パレット「未」で未採点に戻せる
    page.locator("#score-palette .palette-btn", has_text="未").first.click()
    page.locator("#desc-grid .entry-card[data-index='0']").click()
    page.wait_for_function(
        "document.querySelector('#desc-q-tabs .tab').textContent.includes('(0/2)')")


def test_single_sheet_view_overlays_and_navigation(open_app):
    page = open_app(API)
    enter_mode(page, DESC_CARD)
    page.click("#btn-sheet-review")
    page.wait_for_selector("#sheet-list-view", state="visible")
    page.locator("#sheet-list .sheet-row").first.click()
    page.wait_for_selector("#single-sheet-view", state="visible")
    assert "s1.png（1 / 2）" in page.locator("#single-sheet-name").inner_text()
    assert page.locator("#btn-sheet-prev").is_disabled()
    # 領域オーバーレイは画像ロード後に敷かれるため待機型で確認
    page.wait_for_function(
        "document.querySelectorAll('#annotation-layer .region-box').length === 2")
    # サイドパネルの「満点」で採点 → ラベルと4色背景に反映される
    side1 = page.locator("#sheet-side .side-q").nth(0)
    side1.locator("button", has_text="満点").click()
    page.wait_for_function(
        "document.querySelector('#annotation-layer .region-label').textContent.includes('5点')")
    assert "5点" in side1.locator(".score-badge").inner_text()
    assert "sc-full" in (page.locator("#sheet-side .side-q").nth(0)
                         .get_attribute("class") or "")
    # 次の答案へ → ファイル名が変わり、得点は答案ごとに独立
    page.click("#btn-sheet-next")
    page.wait_for_function(
        "document.querySelector('#single-sheet-name').textContent.includes('s2.png')")
    assert page.locator("#btn-sheet-next").is_disabled()
    page.wait_for_function(
        "document.querySelector('#annotation-layer .region-label')?.textContent.includes('未')")
    # 答案一覧に戻る
    page.click("#btn-single-close")
    page.wait_for_selector("#sheet-list-view", state="visible")


def _inject_big_sheet(page):
    page.evaluate("""() => {
      const c = document.createElement('canvas');
      c.width = 595; c.height = 842;
      const g = c.getContext('2d');
      g.fillStyle = '#fff'; g.fillRect(0, 0, 595, 842);
      const url = c.toDataURL('image/png');
      window.__mockApi.get_sheet_image =
        async (f) => ({ok: true, filename: f ?? 's1.png', data_url: url});
      window.__calls = [];
      const origAdd = window.__mockApi.add_descriptive_question;
      window.__mockApi.add_descriptive_question = async (...a) => {
        window.__calls.push(['add', ...a]);
        return {ok: true, question_id: 'D9'};
      };
      window.__mockApi.update_descriptive_region = async (...a) => {
        window.__calls.push(['update', ...a]);
        return {ok: true};
      };
    }""")


def _drag_on_config(page, x0, y0, x1, y1):
    box = page.locator("#config-image").bounding_box()
    page.mouse.move(box["x"] + x0, box["y"] + y0)
    page.mouse.down()
    page.mouse.move(box["x"] + x1, box["y"] + y1, steps=5)
    page.mouse.up()


def test_config_drag_adds_question_and_rearm_updates(open_app):
    """tk方式: ドラッグで即追加（自動命名・既定配点）、領域再指定は次のドラッグ"""
    page = open_app(API)
    enter_mode(page, DESC_CARD)
    _inject_big_sheet(page)
    page.click("#btn-desc-config")
    page.wait_for_selector("#desc-config-view", state="visible")
    page.fill("#new-q-max", "7")
    page.wait_for_function(
        "document.getElementById('config-image').naturalWidth === 595")
    wait_transition(page)
    # 既存2問の領域枠がオーバーレイに出る
    page.wait_for_function(
        "document.querySelectorAll('#config-overlay .config-box').length === 2")

    _drag_on_config(page, 100, 300, 300, 380)
    page.wait_for_function("(window.__calls ?? []).length === 1")
    call0 = page.evaluate("window.__calls[0]")
    assert call0[0] == "add"
    assert call0[1] == "記述3"        # 既存2問 → 自動で3番
    assert call0[2] == "7"            # 既定配点欄の値
    region = call0[4]
    assert region[2] > region[0] and region[3] > region[1]

    # 領域再指定: ボタン → 次のドラッグが update になる
    page.locator("#desc-table tbody tr").first \
        .locator("button", has_text="領域再指定").click()
    page.wait_for_function(
        "document.getElementById('config-hint').textContent.includes('D1')")
    _drag_on_config(page, 60, 100, 260, 160)
    page.wait_for_function("(window.__calls ?? []).length === 2")
    call1 = page.evaluate("window.__calls[1]")
    assert call1[0] == "update" and call1[1] == "D1"
