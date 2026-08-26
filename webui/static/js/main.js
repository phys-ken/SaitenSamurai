import { call } from './bridge.js';
import { openChecker, wireChecker } from './checker.js';
import { pickRegion } from './region-picker.js';
import { wireDescriptive } from './descriptive.js';

// ---------------------------------------------------------------
// ログ
// ---------------------------------------------------------------
function log(message) {
  const el = document.getElementById('log');
  el.textContent += (el.textContent ? '\n' : '') + message;
  el.scrollTop = el.scrollHeight;
}

// ---------------------------------------------------------------
// state → 画面（UIは state の再描画だけを行う）
// ---------------------------------------------------------------
/** 長いパスは末尾（ファイル名側）を残して省略。RTL トリックは
    先頭の '/' が末尾に見える bidi の副作用があるため使わない */
function trimPath(path, max = 64) {
  if (!path || path.length <= max) return path ?? '';
  return '…' + path.slice(-(max - 1));
}

function setRow(rowId, summaryId, path, summaryHtmlBuilder) {
  const value = document.querySelector(`#${rowId} .ds-value`);
  value.textContent = trimPath(path);
  value.title = path ?? '';
  value.classList.toggle('selected', Boolean(path));
  const summary = summaryId ? document.getElementById(summaryId) : null;
  if (!summary) return;
  const built = path ? summaryHtmlBuilder() : null;
  if (built) {
    summary.textContent = built.text;
    summary.className = `ds-summary ${built.kind ?? ''}`;
    summary.hidden = false;
  } else {
    summary.hidden = true;
  }
}

export function render(state) {
  setRow('row-image-folder', 'summary-image-folder', state.image_folder,
    () => ({ text: `画像 ${state.image_count} 枚`, kind: '' }));

  setRow('row-coord-file', 'summary-coord-file', state.coord_file, () => {
    const s = state.coord_summary;
    if (!s) return null;
    if (s.warning) return { text: `⚠ ${s.warning}`, kind: 'warn' };
    return { text: `解答欄 ${s.answer_rows} 行 / 1行あたり ${s.marks_per_row} マーク`, kind: '' };
  });

  setRow('row-answer-key', 'summary-answer-key', state.answer_key, () => {
    const k = state.key_summary;
    if (!k) return null;
    if (!k.ok) return { text: `❌ ${k.errors.join('\n')}`, kind: 'err' };
    if (k.warnings.length) {
      return { text: `⚠ ${k.warnings.join('\n')}\n✓ ${k.stats_line}`, kind: 'warn' };
    }
    return { text: `✓ ${k.stats_line}`, kind: '' };
  });

  document.getElementById('skip-input').value = state.skip_questions;

  // --- Step 1 ---
  document.getElementById('omr-mode').value = state.omr_mode;
  document.getElementById('threshold-inputs').hidden = state.omr_mode !== 'threshold';
  document.getElementById('color-th').value = state.color_threshold;
  document.getElementById('area-th').value = state.area_threshold;
  setRow('row-omr-result', null, state.omr_result, () => null);

  // --- モード別の表示分岐 ---
  const isDescMode = state.app_mode === 'mark_and_descriptive' ||
                     state.app_mode === 'descriptive_only';
  const isDescOnly = state.app_mode === 'descriptive_only';
  currentMode.descOnly = isDescOnly;
  document.querySelectorAll('.desc-only-ui').forEach((n) => { n.hidden = !isDescMode; });
  document.getElementById('row-coord-file').hidden = isDescOnly;
  document.getElementById('summary-coord-file').hidden =
    isDescOnly || !state.coord_file || !state.coord_summary;
  document.getElementById('row-answer-key').hidden = isDescOnly;
  document.getElementById('summary-answer-key').hidden =
    isDescOnly || !state.answer_key || !state.key_summary;
  document.querySelector('.ds-options').hidden = isDescOnly;
  document.getElementById('omr-mode').parentElement.hidden = isDescOnly;
  document.getElementById('row-omr-result').hidden = isDescOnly;
  document.getElementById('btn-open-checker').parentElement.hidden = isDescOnly;
  document.getElementById('btn-run-recognition').textContent =
    isDescOnly ? '▶ 画像準備' : '▶ 認識実行';
  if (state.descriptive) {
    const d = state.descriptive;
    const total = d.questions.reduce((n, q) => n + (d.scored_counts[q.id] ?? 0), 0);
    document.getElementById('desc-progress-hint').textContent =
      d.questions.length === 0 ? '記述問題は未設定です'
        : `${d.questions.length} 問 / 採点済み ${total} / ${d.questions.length * d.prepared_count}`;
  }

  const running = state.job?.running;
  document.getElementById('btn-run-recognition').disabled =
    Boolean(running) || !state.image_folder ||
    (!isDescOnly && !state.coord_file);
  document.getElementById('btn-open-checker').disabled =
    Boolean(running) || !state.omr_result;
  document.getElementById('btn-total-position').disabled = Boolean(running) || !state.image_folder;
  const hasPrepared = (state.descriptive?.prepared_count ?? 0) > 0;
  document.getElementById('btn-desc-config').disabled =
    Boolean(running) || !hasPrepared;
  document.getElementById('btn-desc-scoring').disabled =
    Boolean(running) || !(state.descriptive?.questions?.length);
  const markReady = state.image_folder && state.coord_file &&
    state.answer_key && state.omr_result && state.key_summary?.ok;
  const descReady = Boolean(state.descriptive?.questions?.length);
  const scoringReady = isDescOnly ? descReady
    : (isDescMode ? (markReady && descReady) : markReady);
  document.getElementById('btn-run-scoring').disabled =
    Boolean(running) || !scoringReady;
  document.getElementById('btn-run-summary').disabled =
    Boolean(running) || !scoringReady;
  document.getElementById('btn-cancel').hidden = !running;
  document.getElementById('job-progress').hidden = !running;
}

// ---------------------------------------------------------------
// 操作
// ---------------------------------------------------------------
const ACTION_LOG_LABEL = {
  select_image_folder: '画像フォルダ',
  select_coord_file: '座標ファイル',
  select_answer_key: '正答データ',
  select_omr_result: 'OMR結果',
};

async function runAction(name) {
  try {
    const res = await call(name);
    if (res.cancelled) return;
    render(res.state);
    log(`✓ ${ACTION_LOG_LABEL[name]}を選択しました`);
  } catch (e) {
    log(`❌ ${e.message}`);
    alert(e.message);
  }
}

// ---------------------------------------------------------------
// Python からの push イベント（進捗・完了）
// ---------------------------------------------------------------
window.saitenEvents = (ev) => {
  if (ev.type === 'progress') {
    const bar = document.getElementById('job-progress');
    bar.hidden = false;
    bar.max = ev.total;
    bar.value = ev.current;
    document.getElementById('job-status').textContent = `${ev.current} / ${ev.total}`;
  } else if (ev.type === 'job_done') {
    document.getElementById('job-status').textContent = '';
    log(ev.ok ? `✓ ${ev.message}` : `❌ ${ev.message}`);
    call('get_state').then((res) => render(res.state));
  }
};

const JOB_START_LOG = {
  run_recognition: '▶ 認識を開始しました',
  run_prepare_images: '▶ 画像準備を開始しました',
  run_scoring: '▶ 採点を開始しました',
  run_summary: '▶ 集計を開始しました',
};

async function runJob(name) {
  try {
    const res = await call(name);
    render(res.state);
    log(JOB_START_LOG[name]);
  } catch (e) {
    log(`❌ ${e.message}`);
    alert(e.message);
  }
}
let currentMode = { descOnly: false };
const runRecognition = () => runJob(
  currentMode.descOnly ? 'run_prepare_images' : 'run_recognition');

function wireEvents() {
  document.getElementById('btn-run-recognition')
    .addEventListener('click', runRecognition);
  document.getElementById('btn-run-scoring')
    .addEventListener('click', () => runJob('run_scoring'));
  document.getElementById('btn-total-position').addEventListener('click', async () => {
    try {
      const sheet = await call('get_sheet_image');
      const existing = (await call('get_total_display_region')).region;
      const region = await pickRegion(sheet.data_url, {
        title: `合計点を表示する位置をドラッグで指定（${sheet.filename}）`,
        existing,
      });
      if (region === null) return;
      await call('set_total_display_region', region);
      document.getElementById('total-position-hint').textContent =
        `設定済み (${region.join(', ')})`;
      log('📐 合計点表示位置を設定しました');
    } catch (e) {
      log(`❌ ${e.message}`);
      alert(e.message);
    }
  });
  document.getElementById('btn-run-summary')
    .addEventListener('click', () => runJob('run_summary'));
  document.getElementById('btn-open-checker').addEventListener('click', async () => {
    try {
      await openChecker(log);
    } catch (e) {
      log(`❌ ${e.message}`);
      alert(e.message);
    }
  });
  wireChecker(log);
  wireDescriptive(log, render);
  document.getElementById('btn-cancel')
    .addEventListener('click', async () => { await call('cancel_job'); log('⏹ 中断を要求しました'); });
  document.getElementById('omr-mode').addEventListener('change', async (ev) => {
    const res = await call('set_omr_mode', ev.target.value);
    render(res.state);
  });
  for (const id of ['color-th', 'area-th']) {
    document.getElementById(id).addEventListener('change', async () => {
      try {
        const res = await call('set_thresholds',
          document.getElementById('color-th').value,
          document.getElementById('area-th').value);
        render(res.state);
      } catch (e) {
        log(`❌ ${e.message}`);
        const res = await call('get_state');
        render(res.state);
      }
    });
  }

  document.querySelectorAll('[data-action]').forEach((btn) => {
    btn.addEventListener('click', () => runAction(btn.dataset.action));
  });
  document.getElementById('skip-input').addEventListener('change', async (ev) => {
    try {
      const res = await call('set_skip_questions', ev.target.value);
      render(res.state);
    } catch (e) {
      log(`❌ ${e.message}`);
      const res = await call('get_state');
      render(res.state);  // 不正値は元に戻す
    }
  });
}

// ---------------------------------------------------------------
// 起動
// ---------------------------------------------------------------
const MODE_LABEL = {
  'mark_only/standard': 'マーク採点',
  'mark_and_descriptive/standard': 'マーク採点＋記述採点',
  'descriptive_only/standard': '記述採点',
  'mark_only/multi_digit': '数学マーク採点（複数桁）',
  'mark_and_descriptive/multi_digit': '数学マーク採点＋記述採点',
};

function showMain(state) {
  document.getElementById('mode-select').hidden = true;
  document.querySelectorAll('main > .panel').forEach((p) => { p.hidden = false; });
  const label = MODE_LABEL[`${state.app_mode}/${state.mark_format}`] ?? '';
  let badge = document.getElementById('mode-badge');
  if (!badge) {
    badge = document.createElement('span');
    badge.id = 'mode-badge';
    badge.className = 'mode-badge';
    document.querySelector('.topbar h1').appendChild(badge);
  }
  badge.textContent = `— ${label}`;
  render(state);
}

function wireModeSelect() {
  document.querySelectorAll('.mode-card').forEach((card) => {
    card.addEventListener('click', async () => {
      try {
        const res = await call('set_mode', card.dataset.mode, card.dataset.format);
        showMain(res.state);
        log(`モード: ${MODE_LABEL[`${res.state.app_mode}/${res.state.mark_format}`]}`);
      } catch (e) {
        alert(e.message);
      }
    });
  });
}

async function init() {
  const status = document.getElementById('bridge-status');
  try {
    await call('ping');
    const info = await call('get_app_info');
    document.getElementById('version').textContent =
      `v${info.app_version} (webui ${info.webui_version})`;
    status.textContent = '接続済み';
    status.dataset.state = 'ok';
    const res = await call('get_state');
    render(res.state);
    wireEvents();
    wireModeSelect();
    // モード選択が出ている間はメインパネルを隠す
    document.querySelectorAll('main > .panel').forEach((p) => { p.hidden = true; });
  } catch (e) {
    status.textContent = `接続エラー: ${e.message}`;
    status.dataset.state = 'error';
  }
}

init();

// テスト・デモ用: bridge 側を直接操作した後に UI を state と同期させるフック
window.__refreshState = async () => {
  const res = await call('get_state');
  render(res.state);
};
