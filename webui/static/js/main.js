import { call } from './bridge.js';
import { openChecker, wireChecker } from './checker.js';
import { pickRegion } from './region-picker.js';
import { wireDescriptive } from './descriptive.js';
import { withTransition } from './transitions.js';
import { wireRenderSettings } from './render-settings.js';
import { wireInfo } from './info.js';

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

let progressCache = { at: 0, value: null };

/** ジョブ完了・ビュー遷移などで進行度を確実に取り直したいときに呼ぶ */
export function invalidateProgress() {
  progressCache.at = 0;
}

async function updateStepper(state) {
  const stepper = document.getElementById('stepper');
  // モード選択中はステッパーを出さない（A6）。表示は showMain が行う
  document.getElementById('step-read-label').textContent =
    state.app_mode === 'descriptive_only' ? '画像準備' : '読み取り';
  try {
    // 進行度はフォルダ走査を伴うため短時間メモ化する（A2）
    const now = Date.now();
    if (!progressCache.value || now - progressCache.at > 3000) {
      progressCache = { at: now, value: (await call('get_progress')).progress };
    }
    const p = progressCache.value;
    const order = ['prepared', 'read', 'scored', 'summarized'];
    let currentSet = false;
    for (const li of stepper.querySelectorAll('li')) {
      const done = p[li.dataset.step];
      li.classList.toggle('done', Boolean(done));
      const isCurrent = !done && !currentSet;
      li.classList.toggle('current', isCurrent);
      if (isCurrent) currentSet = true;
    }
    applyNextAction(state, p);
  } catch { /* 進行度が取れなくてもUIは動かす */ }
}

/** 「次に押すべきボタン」をひとつだけ朱にする（#3: マストの明示） */
function applyNextAction(state, p) {
  document.querySelectorAll('.btn-next').forEach((b) => b.classList.remove('btn-next'));
  const running = state.job?.running;
  if (running) return;
  const isDescOnly = state.app_mode === 'descriptive_only';
  const isDescMode = isDescOnly || state.app_mode === 'mark_and_descriptive';
  const d = state.descriptive;
  const descDone = d && d.questions.length > 0 &&
    d.questions.reduce((n, q) => n + (d.scored_counts[q.id] ?? 0), 0)
      >= d.questions.length * d.prepared_count;
  let id = null;
  if (!state.image_folder) id = null;                       // フォルダ選択は data-action
  else if (!isDescOnly && !state.coord_file) id = null;     // 同上（下で解決）
  else if (!p.read) id = 'btn-run-recognition';
  else if (!isDescOnly && !state.key_summary?.ok) id = null;
  else if (isDescMode && !(d?.questions?.length)) id = 'btn-desc-config';
  else if (isDescMode && !descDone) id = 'btn-desc-scoring';
  else if (!p.scored) id = 'btn-run-scoring';
  else if (!p.summarized) id = 'btn-run-summary';
  let el2 = id && document.getElementById(id);
  if (!el2) {
    // データソース系はボタン要素を直接指す
    if (!state.image_folder) {
      el2 = document.querySelector('[data-action=select_image_folder]');
    } else if (!isDescOnly && !state.coord_file) {
      el2 = document.querySelector('[data-action=select_coord_file]');
    } else if (!isDescOnly && !state.key_summary?.ok && p.read) {
      el2 = document.querySelector('[data-action=select_answer_key]');
    }
  }
  el2?.classList.add('btn-next');
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

  const folderEl = document.getElementById('current-folder');
  if (state.image_folder) {
    const name = state.image_folder.replace(/[\\/]+$/, '').split(/[\\/]/).pop();
    folderEl.textContent = name;
    folderEl.title = state.image_folder;
  } else {
    folderEl.textContent = '';
    folderEl.title = '';
  }

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
  currentMode.descMode = isDescMode;
  currentMode.appMode = state.app_mode;
  lastState = state;
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
  document.querySelectorAll('.mark-only-ui').forEach((n) => { n.hidden = isDescMode; });
  document.getElementById('btn-run-recognition').textContent =
    isDescOnly ? '▶ 画像を準備する' : '▶ 答案を読み取る';
  if (state.descriptive) {
    const d = state.descriptive;
    const total = d.questions.reduce((n, q) => n + (d.scored_counts[q.id] ?? 0), 0);
    document.getElementById('desc-progress-hint').textContent =
      d.questions.length === 0 ? '記述問題は未設定です'
        : `${d.questions.length} 問 / 採点済み ${total} / ${d.questions.length * d.prepared_count}`;
  }

  document.getElementById('total-position-hint').textContent =
    state.total_display_region
      ? `設定済み (${state.total_display_region.join(', ')})`
      : '未設定（自動配置）';

  document.getElementById('row-include-desc').hidden =
    state.app_mode !== 'mark_and_descriptive';
  document.getElementById('include-desc-check').checked =
    Boolean(state.include_descriptive_in_analysis);

  document.getElementById('name-trim-check').checked =
    Boolean(state.name_trim_enabled);
  document.getElementById('btn-name-position').disabled =
    !state.name_trim_enabled || !state.image_folder;
  document.getElementById('name-position-hint').textContent =
    !state.name_trim_enabled ? '' :
      (state.name_trim_region
        ? `指定済み (${state.name_trim_region.join(', ')})`
        : '未指定（指定すると集計に氏名画像が入ります）');

  const running = state.job?.running;
  // 実行中のフォルダ/ファイル変更はワーカーの読む state を壊すため止める（A4）
  document.querySelectorAll('[data-action]').forEach((b) => {
    b.disabled = Boolean(running);
  });
  document.getElementById('btn-select-pdf').disabled = Boolean(running);
  document.getElementById('btn-recheck-key').disabled =
    Boolean(running) || !state.answer_key;
  document.getElementById('skip-input').disabled = Boolean(running);
  document.getElementById('btn-run-recognition').disabled =
    Boolean(running) || !state.image_folder ||
    (!isDescOnly && !state.coord_file);
  document.getElementById('btn-open-checker').disabled =
    Boolean(running) || !state.omr_result;
  document.getElementById('btn-total-position').disabled = Boolean(running) || !state.image_folder;
  document.getElementById('btn-annotate').disabled =
    Boolean(running) || !state.image_folder ||
    !(isDescOnly ? (state.descriptive?.prepared_count ?? 0) > 0
                 : Boolean(state.omr_result));
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
  document.getElementById('job-panel').hidden = !running;
  if (running) {
    document.getElementById('job-kind').textContent =
      JOB_KIND_LABEL[state.job.kind] ?? '処理中';
  }
  updateStepper(state);
}

// ---------------------------------------------------------------
// 操作
// ---------------------------------------------------------------
const ACTION_LOG_LABEL = {
  select_image_folder: '答案画像フォルダ',
  select_coord_file: 'マーク位置の定義',
  select_answer_key: '正答・配点',
  select_omr_result: '読み取り結果',
};

async function runAction(name) {
  try {
    const res = await call(name);
    if (res.cancelled) return;
    render(res.state);
    log(`✓ ${ACTION_LOG_LABEL[name]}を選択しました`);
    if (res.auto_detected_answer_key) {
      log(`✓ 正答データを自動検出しました: ${res.auto_detected_answer_key}`);
    }
    if (res.session_found &&
        confirm('このフォルダには前回のセッション情報が見つかりました。\n前回の設定を復元しますか？')) {
      const restored = await call('restore_session', res.session_found);
      showMain(restored.state);
      log('✓ 前回のセッションを復元しました');
      for (const w of restored.warnings ?? []) log(`⚠ ${w}`);
    }
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
    document.getElementById('job-panel').hidden = false;
    document.getElementById('job-kind').textContent =
      JOB_KIND_LABEL[ev.kind] ?? '処理中';
    const bar = document.getElementById('job-progress');
    bar.max = ev.total;
    bar.value = ev.current;
    document.getElementById('job-status').textContent = `${ev.current} / ${ev.total}`;
  } else if (ev.type === 'job_done') {
    invalidateProgress();
    document.getElementById('job-panel').hidden = true;
    document.getElementById('job-status').textContent = '';
    log(ev.ok ? `✓ ${ev.message}` : `❌ ${ev.message}`);
    call('get_state').then((res) => render(res.state));
  }
};

const JOB_KIND_LABEL = {
  recognition: '答案の読み取り',
  pdf_import: 'PDF展開',
  prepare_images: '画像の準備',
  scoring: '採点済み答案の生成',
  summary: '集計',
};

const JOB_START_LOG = {
  select_pdf: '▶ PDF展開を開始しました',
  run_recognition: '▶ 答案の読み取りを開始しました',
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
let currentMode = { descOnly: false, descMode: false, appMode: 'mark_only' };
let lastState = null;
const currentModeName = () => currentMode.appMode;
const runRecognition = () => runJob(
  currentMode.descOnly ? 'run_prepare_images' : 'run_recognition');

/** 記述採点が未完了なら続行確認する（tk 版と同じ「未採点は0点」警告） */
function confirmDescriptiveComplete() {
  const d = lastState?.descriptive;
  if (!currentMode.descMode || !d || !d.questions.length) return true;
  const total = d.questions.length * d.prepared_count;
  const done = d.questions.reduce((n, q) => n + (d.scored_counts[q.id] ?? 0), 0);
  if (done >= total) return true;
  return confirm(`記述採点が完了していません（採点済み ${done} / ${total}）。\n` +
                 '未採点の問題は 0点 として扱われます。このまま続行しますか？');
}

function wireEvents() {
  document.getElementById('btn-recheck-key').addEventListener('click', async () => {
    try {
      const res = await call('recheck_answer_key');
      render(res.state);
      log('✓ 正答データを再チェックしました');
    } catch (e) { log(`❌ ${e.message}`); }
  });
  document.querySelectorAll('[data-open-folder]').forEach((btn) => {
    btn.addEventListener('click', async () => {
      try {
        await call('open_folder', btn.dataset.openFolder);
      } catch (e) { log(`❌ ${e.message}`); }
    });
  });
  document.getElementById('include-desc-check').addEventListener('change', async (ev) => {
    try {
      const res = await call('set_include_descriptive_in_analysis', ev.target.checked);
      render(res.state);
    } catch (e) { log(`❌ ${e.message}`); }
  });
  document.getElementById('btn-run-recognition')
    .addEventListener('click', runRecognition);
  document.getElementById('btn-run-scoring')
    .addEventListener('click', () => {
      if (!confirmDescriptiveComplete()) return;
      runJob('run_scoring');
    });
  document.getElementById('btn-total-position').addEventListener('click', async () => {
    try {
      const sheet = await call('get_sheet_image');
      const existing = (await call('get_total_display_region')).region;
      const region = await pickRegion(sheet.data_url, {
        title: `合計点を表示する位置をドラッグで指定（${sheet.filename}）`,
        existing,
      });
      if (region === null) return;
      const set = await call('set_total_display_region', region);
      render(set.state);
      log('✓ 合計点の表示位置を設定しました');
    } catch (e) {
      log(`❌ ${e.message}`);
      alert(e.message);
    }
  });
  document.getElementById('btn-run-summary')
    .addEventListener('click', () => {
      if (!confirmDescriptiveComplete()) return;
      runJob('run_summary');
    });
  document.getElementById('btn-open-checker').addEventListener('click', async () => {
    try {
      await openChecker(log);
    } catch (e) {
      log(`❌ ${e.message}`);
      alert(e.message);
    }
  });
  document.getElementById('btn-select-pdf').addEventListener('click', async () => {
    try {
      const res = await call('select_pdf');
      if (res.cancelled) return;
      render(res.state);
      log(JOB_START_LOG.select_pdf);
    } catch (e) {
      log(`❌ ${e.message}`);
      alert(e.message);
    }
  });
  document.getElementById('btn-calibrate').addEventListener('click', async () => {
    try {
      const res = await call('run_threshold_calibration');
      render(res.state);
      log(`🔧 自動調整: 濃さ ${res.color_threshold} / 面積 ${res.area_threshold}` +
          `（${res.image_count}枚から推定）`);
    } catch (e) {
      log(`❌ ${e.message}`);
      alert(e.message);
    }
  });
  document.getElementById('name-trim-check').addEventListener('change', async (ev) => {
    try {
      const res = await call('set_name_trim_enabled', ev.target.checked);
      render(res.state);
    } catch (e) { log(`❌ ${e.message}`); }
  });
  document.getElementById('btn-name-position').addEventListener('click', async () => {
    try {
      const sheet = await call('get_sheet_image');
      const state = (await call('get_state')).state;
      const region = await pickRegion(sheet.data_url, {
        title: `氏名欄をドラッグで指定（${sheet.filename}）`,
        existing: state.name_trim_region,
      });
      if (region === null) return;
      const res = await call('set_name_trim_region', region);
      render(res.state);
      log('📐 氏名欄の位置を設定しました');
    } catch (e) {
      log(`❌ ${e.message}`);
      alert(e.message);
    }
  });
  wireRenderSettings(log, () => {
    const badge = document.getElementById('mode-badge');
    return badge ? currentModeName() : 'mark_only';
  });
  document.querySelectorAll('#stepper li').forEach((li) => {
    li.addEventListener('click', () => {
      document.getElementById(li.dataset.target)
        ?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  });
  wireChecker(log, render);
  wireDescriptive(log, render);
  document.getElementById('btn-cancel')
    .addEventListener('click', async () => {
      try {
        await call('cancel_job');
        log('⏹ 中断を要求しました');
      } catch (e) { log(`❌ ${e.message}`); }
    });
  document.getElementById('omr-mode').addEventListener('change', async (ev) => {
    try {
      const res = await call('set_omr_mode', ev.target.value);
      render(res.state);
    } catch (e) { log(`❌ ${e.message}`); }
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
  withTransition(() => {
    document.getElementById('mode-select').hidden = true;
    document.getElementById('stepper').hidden = false;
    document.getElementById('main-cols').hidden = false;
    const label = MODE_LABEL[`${state.app_mode}/${state.mark_format}`] ?? '';
    let badge = document.getElementById('mode-badge');
    if (!badge) {
      badge = document.createElement('span');
      badge.id = 'mode-badge';
      badge.className = 'mode-badge';
      document.querySelector('.topbar h1').appendChild(badge);
    }
    badge.textContent = label;
    render(state);
  });
}

function wireModeSelect() {
  const resumeBtn = document.getElementById('btn-resume-session');
  resumeBtn.hidden = false;
  resumeBtn.addEventListener('click', async () => {
    try {
      const res = await call('restore_session');
      if (res.cancelled) return;
      showMain(res.state);
      log(`📂 セッションを復元しました（${MODE_LABEL[`${res.state.app_mode}/${res.state.mark_format}`]}）`);
      for (const w of res.warnings ?? []) log(`⚠ ${w}`);
    } catch (e) {
      alert(e.message);
    }
  });
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
    document.getElementById('version').textContent = `v${info.app_version}`;
    document.getElementById('version').title =
      `採点侍 v${info.app_version} / webui ${info.webui_version} / Python ${info.python}`;
    status.textContent = '接続済み';
    status.dataset.state = 'ok';
    const res = await call('get_state');
    render(res.state);
    wireEvents();
    wireModeSelect();
    wireInfo();
    // モード選択が出ている間はメインパネルを隠す
    document.getElementById('main-cols').hidden = true;
  } catch (e) {
    status.textContent = `接続エラー: ${e.message}`;
    status.dataset.state = 'error';
  }
}

init();

// テスト・デモ用: bridge 側を直接操作した後に UI を state と同期させるフック
window.__refreshState = async () => {
  invalidateProgress();
  const res = await call('get_state');
  render(res.state);
};
