/**
 * descriptive.js — 記述採点の3ビュー（設定 / 問題別グリッド / 一枚採点）。
 *
 * 一枚採点は「画像レイヤ + annotation-layer」の重ね構造。将来のコメント
 * モード（手書き描画）は annotation-layer に canvas を足すだけで載る設計。
 */
import { call } from './bridge.js';
import { pickRegion } from './region-picker.js';
import { withTransition } from './transitions.js';

let logFn = () => {};
let onStateUpdate = () => {};
const el = (id) => document.getElementById(id);

function showView(id) {
  withTransition(() => {
    document.querySelectorAll('main > .panel').forEach((p) => { p.hidden = true; });
    for (const v of ['desc-config-view', 'desc-scoring-view', 'single-sheet-view']) {
      el(v).hidden = v !== id;
    }
    if (id === null) {
      document.querySelectorAll('main > .panel').forEach((p) => { p.hidden = false; });
    }
  });
}

// ================================================================
// 設定ビュー
// ================================================================

async function renderConfigTable() {
  const res = await call('get_state');
  const desc = res.state.descriptive;
  const tbody = el('desc-table').querySelector('tbody');
  tbody.replaceChildren(...desc.questions.map((q) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${q.id}</td><td>${q.name}</td><td>${q.max_score}点</td>` +
      `<td>${q.aspect}</td><td>(${q.region.join(', ')})</td>`;
    const td = document.createElement('td');
    const reBtn = document.createElement('button');
    reBtn.className = 'btn';
    reBtn.textContent = '領域再指定';
    reBtn.addEventListener('click', async () => {
      const sheet = await call('get_sheet_image');
      const region = await pickRegion(sheet.data_url,
        { title: `${q.name} の領域を指定（${sheet.filename}）`, existing: q.region });
      if (region === null) return;
      await call('update_descriptive_region', q.id, region);
      await renderConfigTable();
    });
    const delBtn = document.createElement('button');
    delBtn.className = 'btn';
    delBtn.textContent = '削除';
    delBtn.addEventListener('click', async () => {
      if (!confirm(`${q.name}（${q.id}）を削除しますか？ 採点済みの得点も消えます`)) return;
      await call('delete_descriptive_question', q.id);
      logFn(`🗑 ${q.id} を削除しました`);
      await renderConfigTable();
    });
    td.append(reBtn, ' ', delBtn);
    tr.appendChild(td);
    return tr;
  }));
  el('desc-config-summary').textContent =
    `${desc.questions.length} 問 / 対象画像 ${desc.prepared_count} 枚`;
  onStateUpdate(res.state);
}

export async function openDescConfig() {
  showView('desc-config-view');
  await renderConfigTable();
}

// ================================================================
// 問題別グリッド採点ビュー
// ================================================================

let currentQid = null;

function scoreButtons(maxScore, current, onSet) {
  const wrap = document.createElement('div');
  wrap.className = 'score-btns';
  for (let v = 0; v <= maxScore; v++) {
    const b = document.createElement('button');
    b.className = 'score-btn' + (current === v ? ' active' : '');
    b.textContent = v;
    b.addEventListener('click', () => onSet(v));
    wrap.appendChild(b);
  }
  const clear = document.createElement('button');
  clear.className = 'score-btn clear';
  clear.textContent = '未';
  clear.title = '未採点に戻す';
  clear.addEventListener('click', () => onSet(null));
  wrap.appendChild(clear);
  return wrap;
}

async function renderQTabs() {
  const res = await call('get_state');
  const desc = res.state.descriptive;
  el('desc-q-tabs').replaceChildren(...desc.questions.map((q) => {
    const b = document.createElement('button');
    const done = desc.scored_counts[q.id] ?? 0;
    b.className = 'tab' + (q.id === currentQid ? ' active' : '');
    b.textContent = `${q.name} (${done}/${desc.prepared_count})`;
    b.addEventListener('click', async () => {
      currentQid = q.id;
      await renderQTabs();
      await renderDescGrid();
    });
    return b;
  }));
  const total = desc.questions.reduce(
    (n, q) => n + (desc.scored_counts[q.id] ?? 0), 0);
  el('desc-scoring-summary').textContent =
    `採点済み ${total} / ${desc.questions.length * desc.prepared_count}`;
  onStateUpdate(res.state);
}

async function renderDescGrid() {
  const state = (await call('get_state')).state;
  const q = state.descriptive.questions.find((x) => x.id === currentQid);
  const targets = await call('list_descriptive_targets', currentQid);
  el('desc-grid').replaceChildren(...targets.items.map((t) => {
    const card = document.createElement('div');
    card.className = 'entry-card' + (t.score !== null ? ' scored' : '');
    const img = document.createElement('img');
    img.alt = t.filename;
    call('get_descriptive_crop', currentQid, t.filename)
      .then((r) => { img.src = r.data_url; })
      .catch(() => { img.alt = '画像なし'; });
    const meta = document.createElement('div');
    meta.className = 'entry-meta';
    meta.innerHTML = `<span>${t.filename}</span>` +
      (t.score === null
        ? '<span class="score-note">未採点</span>'
        : `<span class="score-note score-mark">${t.score} 点</span>`);
    card.append(img, meta,
      scoreButtons(q.max_score, t.score, async (v) => {
        try {
          await call('set_descriptive_score', t.filename, currentQid, v);
          await renderQTabs();
          await renderDescGrid();
        } catch (e) { logFn(`❌ ${e.message}`); }
      }));
    return card;
  }));
}

export async function openDescScoring() {
  await call('start_descriptive_scoring');
  const state = (await call('get_state')).state;
  if (!state.descriptive.questions.length) return;
  currentQid = currentQid ?? state.descriptive.questions[0].id;
  showView('desc-scoring-view');
  await renderQTabs();
  await renderDescGrid();
  logFn('✏ 記述採点を開きました');
}

// ================================================================
// 一枚採点ビュー
// ================================================================

let sheetFiles = [];
let sheetIndex = 0;
let focusedQid = null;

async function renderSingleSheet() {
  const filename = sheetFiles[sheetIndex];
  el('single-sheet-name').textContent =
    `${filename}（${sheetIndex + 1} / ${sheetFiles.length}）`;
  el('btn-sheet-prev').disabled = sheetIndex === 0;
  el('btn-sheet-next').disabled = sheetIndex >= sheetFiles.length - 1;

  const sheet = await call('get_sheet_image', filename);
  const img = el('sheet-image');
  await new Promise((resolve) => {
    img.onload = resolve;
    img.src = sheet.data_url;
  });

  const state = (await call('get_state')).state;
  const questions = state.descriptive.questions;
  const layer = el('annotation-layer');
  const side = el('sheet-side');
  const sx = img.clientWidth / img.naturalWidth;
  const sy = img.clientHeight / img.naturalHeight;

  const targetsByQ = {};
  for (const q of questions) {
    const t = await call('list_descriptive_targets', q.id);
    targetsByQ[q.id] = t.items.find((i) => i.filename === filename) ?? { score: null };
  }

  layer.replaceChildren(...questions.map((q) => {
    const [x1, y1, x2, y2] = q.region;
    const box = document.createElement('div');
    box.className = 'region-box' +
      (targetsByQ[q.id].score !== null ? ' scored' : '') +
      (q.id === focusedQid ? ' focused' : '');
    Object.assign(box.style, {
      left: `${x1 * sx}px`, top: `${y1 * sy}px`,
      width: `${(x2 - x1) * sx}px`, height: `${(y2 - y1) * sy}px`,
    });
    const label = document.createElement('span');
    label.className = 'region-label';
    const sc = targetsByQ[q.id].score;
    label.textContent = `${q.name}: ${sc === null ? '未' : sc + '点'}`;
    box.appendChild(label);
    box.addEventListener('click', () => {
      focusedQid = q.id;
      renderSingleSheet();
    });
    return box;
  }));

  side.replaceChildren(...questions.map((q) => {
    const div = document.createElement('div');
    div.className = 'side-q' + (q.id === focusedQid ? ' focused' : '');
    const h = document.createElement('h3');
    const sc = targetsByQ[q.id].score;
    h.textContent = `${q.name}（${q.max_score}点満点）` +
      (sc === null ? '' : ` — ${sc}点`);
    div.append(h, scoreButtons(q.max_score, sc, async (v) => {
      try {
        await call('set_descriptive_score', filename, q.id, v);
        focusedQid = q.id;
        await renderSingleSheet();
      } catch (e) { logFn(`❌ ${e.message}`); }
    }));
    return div;
  }));
}

export async function openSingleSheet() {
  const listing = await call('list_sheet_files');
  sheetFiles = listing.files;
  sheetIndex = 0;
  showView('single-sheet-view');
  await renderSingleSheet();
  logFn('📄 一枚採点モードを開きました');
}

// ================================================================
// 配線
// ================================================================

export function wireDescriptive(log, stateUpdate) {
  logFn = log;
  onStateUpdate = stateUpdate;

  el('btn-desc-config').addEventListener('click', () =>
    openDescConfig().catch((e) => { log(`❌ ${e.message}`); alert(e.message); }));
  el('btn-desc-config-close').addEventListener('click', () => showView(null));

  el('btn-add-question').addEventListener('click', async () => {
    try {
      const sheet = await call('get_sheet_image');
      const name = el('new-q-name').value || `問${Date.now() % 1000}`;
      const region = await pickRegion(sheet.data_url,
        { title: `${name} の採点領域をドラッグで指定（${sheet.filename}）` });
      if (region === null) return;
      const res = await call('add_descriptive_question',
        name, el('new-q-max').value, el('new-q-aspect').value, region);
      log(`＋ ${res.question_id}（${name}）を追加しました`);
      el('new-q-name').value = '';
      await renderConfigTable();
    } catch (e) { log(`❌ ${e.message}`); alert(e.message); }
  });

  el('btn-desc-scoring').addEventListener('click', () =>
    openDescScoring().catch((e) => { log(`❌ ${e.message}`); alert(e.message); }));
  el('btn-desc-scoring-close').addEventListener('click', () => showView(null));

  el('btn-single-sheet').addEventListener('click', () =>
    openSingleSheet().catch((e) => { log(`❌ ${e.message}`); alert(e.message); }));
  el('btn-single-close').addEventListener('click', () =>
    openDescScoring().catch((e) => { log(`❌ ${e.message}`); }));
  el('btn-sheet-prev').addEventListener('click', () => { sheetIndex--; renderSingleSheet(); });
  el('btn-sheet-next').addEventListener('click', () => { sheetIndex++; renderSingleSheet(); });
}
