/**
 * descriptive.js — 記述採点の3ビュー（設定 / 問題別グリッド / 一枚採点）。
 *
 * 一枚採点は「画像レイヤ + annotation-layer」の重ね構造。将来のコメント
 * モード（手書き描画）は annotation-layer に canvas を足すだけで載る設計。
 */
import { call } from './bridge.js';
import { pickRegion } from './region-picker.js';
import { withTransition } from './transitions.js';
import { createVirtualGrid, createImageLoader } from './vlist.js';

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
// 問題別グリッド採点ビュー（仮想スクロール＋キー採点）
// ================================================================

let currentQid = null;
let gridTargets = [];        // [{filename, score}] 現在の問題の全生徒
let grid = null;             // createVirtualGrid のハンドル
let gridCursor = 0;          // 採点カーソル（署名要素: 朱の採点枠）
let gridMax = 0;             // 現在の問題の満点
let digitBuf = '';
let digitTimer = null;

const cropLoader = createImageLoader(
  (qid, filename) => call('get_descriptive_crop', qid, filename)
    .then((r) => r.data_url),
  { concurrency: 6, cacheSize: 400 });

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

function openLightbox(src, alt) {
  const box = document.createElement('div');
  box.className = 'lightbox';
  const img = document.createElement('img');
  img.src = src;
  img.alt = alt ?? '';
  box.appendChild(img);
  const close = () => { box.remove(); document.removeEventListener('keydown', onKey); };
  const onKey = (e) => { if (e.key === 'Escape') { e.stopPropagation(); close(); } };
  box.addEventListener('click', close);
  document.addEventListener('keydown', onKey);
  document.body.appendChild(box);
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

/** 問題の領域縦横比からカードの固定高さを決める（仮想化の前提） */
function cardMetrics(q) {
  const [x1, y1, x2, y2] = q.region;
  const aspect = (y2 - y1) / Math.max(1, x2 - x1);
  const itemMinWidth = 240;
  const imgH = Math.min(260, Math.max(70, Math.round(itemMinWidth * aspect)));
  const perRow = 6;                                   // 240px に収まる点数ボタン数
  const btnRows = Math.ceil((q.max_score + 2) / perRow);
  const itemHeight = imgH + 24 + btnRows * 32 + 22;   // meta行 + ボタン行 + 余白
  return { itemMinWidth, imgH, itemHeight };
}

function buildCard(i, q, metrics) {
  const t = gridTargets[i];
  const card = document.createElement('div');
  card.className = 'entry-card'
    + (t.score !== null ? ' scored' : '')
    + (i === gridCursor ? ' cursor' : '');
  card.dataset.index = i;

  const img = document.createElement('img');
  img.className = 'crop-img';
  img.style.height = `${metrics.imgH}px`;
  img.alt = t.filename;
  cropLoader.load(`${currentQid}/${t.filename}`, currentQid, t.filename)
    .then((url) => { img.src = url; })
    .catch(() => { img.alt = '画像を読み込めません'; });
  img.addEventListener('click', async () => {
    gridCursor = i;
    updateCursorClasses();
    try {
      const url = await cropLoader.load(`${currentQid}/${t.filename}`, currentQid, t.filename);
      openLightbox(url, t.filename);
    } catch { /* 画像なしはカード表示のまま */ }
  });

  const meta = document.createElement('div');
  meta.className = 'entry-meta';
  meta.innerHTML = `<span>${t.filename}</span>` +
    (t.score === null
      ? '<span class="score-note">未採点</span>'
      : `<span class="score-note score-mark">${t.score} 点</span>`);

  card.append(img, meta, scoreButtons(q.max_score, t.score,
    (v) => setGridScore(i, v)));
  card.addEventListener('click', (e) => {
    if (e.target === card || e.target === meta) {
      gridCursor = i;
      updateCursorClasses();
    }
  });
  return card;
}

function updateCursorClasses() {
  if (!grid) return;
  document.querySelectorAll('#desc-grid .entry-card.cursor')
    .forEach((c) => c.classList.remove('cursor'));
  grid.itemAt(gridCursor)?.classList.add('cursor');
}

async function setGridScore(i, v) {
  const t = gridTargets[i];
  try {
    await call('set_descriptive_score', t.filename, currentQid, v);
  } catch (e) {
    logFn(`❌ ${e.message}`);
    return;
  }
  t.score = v;
  grid.refresh();
  await renderQTabs();
  if (v !== null) advanceToUnscored(i + 1);
}

/** from 以降（末尾で先頭に折返し）の最初の未採点へカーソルを送る */
function advanceToUnscored(from) {
  const n = gridTargets.length;
  for (let k = 0; k < n; k++) {
    const idx = (from + k) % n;
    if (gridTargets[idx].score === null) {
      gridCursor = idx;
      grid.scrollToIndex(idx);
      updateCursorClasses();
      return;
    }
  }
  logFn(`✓ ${currentQid} は全員採点済みです`);
}

async function renderDescGrid() {
  const state = (await call('get_state')).state;
  const q = state.descriptive.questions.find((x) => x.id === currentQid);
  if (!q) return;
  gridMax = q.max_score;
  const targets = await call('list_descriptive_targets', currentQid);
  gridTargets = targets.items;
  gridCursor = Math.min(gridCursor, Math.max(0, gridTargets.length - 1));
  const metrics = cardMetrics(q);
  grid?.destroy();
  grid = createVirtualGrid(el('desc-grid'), {
    count: gridTargets.length,
    itemMinWidth: metrics.itemMinWidth,
    itemHeight: metrics.itemHeight,
    gap: 10,
    renderItem: (i) => buildCard(i, q, metrics),
  });
  updateCursorClasses();
}

function commitDigits() {
  clearTimeout(digitTimer);
  digitTimer = null;
  if (digitBuf === '') return;
  const v = parseInt(digitBuf, 10);
  digitBuf = '';
  if (v <= gridMax) setGridScore(gridCursor, v);
  else logFn(`❌ ${v} 点は満点（${gridMax}点）を超えています`);
}

function gridKeyHandler(e) {
  if (el('desc-scoring-view').hidden) return;
  const tag = e.target.tagName;
  if (tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA') return;
  if (!grid || !gridTargets.length) return;

  if (e.key >= '0' && e.key <= '9') {
    e.preventDefault();
    if (gridMax <= 9) {
      const v = parseInt(e.key, 10);
      if (v <= gridMax) setGridScore(gridCursor, v);
      else logFn(`❌ ${v} 点は満点（${gridMax}点）を超えています`);
      return;
    }
    digitBuf += e.key;
    if (parseInt(digitBuf + '0', 10) > gridMax) commitDigits();
    else {
      clearTimeout(digitTimer);
      digitTimer = setTimeout(commitDigits, 500);
    }
    return;
  }
  if (e.key === 'Enter') { e.preventDefault(); commitDigits(); return; }
  if (e.key === 'Backspace' || e.key === 'Delete') {
    e.preventDefault();
    digitBuf = '';
    setGridScore(gridCursor, null);
    return;
  }
  const move = { ArrowLeft: -1, ArrowRight: 1,
                 ArrowUp: -grid.columns(), ArrowDown: grid.columns() }[e.key];
  if (move !== undefined) {
    e.preventDefault();
    gridCursor = Math.min(gridTargets.length - 1, Math.max(0, gridCursor + move));
    grid.scrollToIndex(gridCursor);
    updateCursorClasses();
  }
}

export async function openDescScoring() {
  await call('start_descriptive_scoring');
  const state = (await call('get_state')).state;
  if (!state.descriptive.questions.length) return;
  currentQid = currentQid ?? state.descriptive.questions[0].id;
  showView('desc-scoring-view');
  await renderQTabs();
  await renderDescGrid();
  logFn('記述採点を開きました（数字キーで得点、自動で次の未採点へ進みます）');
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
  document.addEventListener('keydown', gridKeyHandler);
  el('btn-jump-unscored').addEventListener('click', () => {
    if (grid && gridTargets.length) advanceToUnscored(gridCursor);
  });
  el('btn-desc-scoring-close').addEventListener('click', () => showView(null));

  el('btn-single-sheet').addEventListener('click', () =>
    openSingleSheet().catch((e) => { log(`❌ ${e.message}`); alert(e.message); }));
  el('btn-single-close').addEventListener('click', () =>
    openDescScoring().catch((e) => { log(`❌ ${e.message}`); }));
  el('btn-sheet-prev').addEventListener('click', () => { sheetIndex--; renderSingleSheet(); });
  el('btn-sheet-next').addEventListener('click', () => { sheetIndex++; renderSingleSheet(); });
}
