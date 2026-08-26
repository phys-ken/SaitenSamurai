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
  return withTransition(() => {
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
  await showView('desc-config-view');
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
  await showView('desc-scoring-view');
  await renderQTabs();
  await renderDescGrid();
  logFn('記述採点を開きました（数字キーで得点、自動で次の未採点へ進みます）');
}

// ================================================================
// 一枚採点ビュー（ズーム/パン＋キーボード動線）
// ================================================================

let sheetFiles = [];
let sheetIndex = 0;
let focusedQid = null;
let sheetQuestions = [];          // openSingleSheet で確定
let sheetScores = {};             // qid -> {filename -> score} 全読み込み・ローカル更新
let sheetZoom = null;             // null = 幅フィット / 数値 = 倍率
let sheetDigitBuf = '';
let sheetDigitTimer = null;

const sheetLoader = createImageLoader(
  (filename) => call('get_sheet_image', filename).then((r) => r.data_url),
  { concurrency: 3, cacheSize: 30 });

function sheetScore(qid, filename) {
  const v = sheetScores[qid]?.[filename];
  return v === undefined ? null : v;
}

function sheetCompleteCount() {
  return sheetFiles.filter((f) =>
    sheetQuestions.every((q) => sheetScore(q.id, f) !== null)).length;
}

function renderSheetHead() {
  const filename = sheetFiles[sheetIndex];
  el('single-sheet-name').textContent =
    `${filename}（${sheetIndex + 1} / ${sheetFiles.length}）` +
    ` ─ 全問採点済み ${sheetCompleteCount()} 枚`;
  el('btn-sheet-prev').disabled = sheetIndex === 0;
  el('btn-sheet-next').disabled = sheetIndex >= sheetFiles.length - 1;
}

/** 現在のズーム設定に合わせて画像幅を決め、領域オーバーレイを敷き直す */
function layoutSheet() {
  const img = el('sheet-image');
  const stage = el('sheet-stage');
  if (!img.naturalWidth) return;
  if (sheetZoom === null) {
    img.style.width = '100%';
  } else {
    img.style.width = `${img.naturalWidth * sheetZoom}px`;
  }
  const pct = Math.round((img.clientWidth / img.naturalWidth) * 100);
  el('zoom-label').textContent = `${pct}%`;

  const filename = sheetFiles[sheetIndex];
  const sx = img.clientWidth / img.naturalWidth;
  const sy = img.clientHeight / img.naturalHeight;
  const layer = el('annotation-layer');
  layer.style.width = `${img.clientWidth}px`;
  layer.style.height = `${img.clientHeight}px`;
  layer.replaceChildren(...sheetQuestions.map((q) => {
    const [x1, y1, x2, y2] = q.region;
    const sc = sheetScore(q.id, filename);
    const box = document.createElement('div');
    box.className = 'region-box' +
      (sc !== null ? ' scored' : '') +
      (q.id === focusedQid ? ' focused' : '');
    Object.assign(box.style, {
      left: `${x1 * sx}px`, top: `${y1 * sy}px`,
      width: `${(x2 - x1) * sx}px`, height: `${(y2 - y1) * sy}px`,
    });
    const label = document.createElement('span');
    label.className = 'region-label';
    label.textContent = `${q.name}: ${sc === null ? '未' : sc + '点'}`;
    box.appendChild(label);
    box.addEventListener('click', () => { focusedQid = q.id; layoutSheet(); renderSheetSide(); });
    return box;
  }));
}

function renderSheetSide() {
  const filename = sheetFiles[sheetIndex];
  const side = el('sheet-side');
  side.replaceChildren(...sheetQuestions.map((q) => {
    const div = document.createElement('div');
    div.className = 'side-q' + (q.id === focusedQid ? ' focused' : '');
    const h = document.createElement('h3');
    const sc = sheetScore(q.id, filename);
    h.textContent = `${q.name}（${q.max_score}点満点）`;
    if (sc !== null) {
      const mark = document.createElement('span');
      mark.className = 'score-mark';
      mark.textContent = ` ${sc}点`;
      h.appendChild(mark);
    }
    div.append(h, scoreButtons(q.max_score, sc,
      (v) => setSheetScore(q.id, v)));
    div.addEventListener('click', (e) => {
      if (e.target === div || e.target === h) {
        focusedQid = q.id;
        layoutSheet();
        renderSheetSide();
      }
    });
    return div;
  }));
}

async function setSheetScore(qid, v) {
  const filename = sheetFiles[sheetIndex];
  try {
    await call('set_descriptive_score', filename, qid, v);
  } catch (e) {
    logFn(`❌ ${e.message}`);
    return;
  }
  (sheetScores[qid] ??= {})[filename] = v === null ? undefined : v;
  if (v === null) delete sheetScores[qid][filename];
  if (v !== null) advanceSheetFocus(qid);
  layoutSheet();
  renderSheetSide();
  renderSheetHead();
}

/** 採点後の自動送り: 同じ答案の次の未採点問題 → なければ次の答案へ */
function advanceSheetFocus(fromQid) {
  const filename = sheetFiles[sheetIndex];
  const start = sheetQuestions.findIndex((q) => q.id === fromQid);
  for (let k = 1; k <= sheetQuestions.length; k++) {
    const q = sheetQuestions[(start + k) % sheetQuestions.length];
    if (sheetScore(q.id, filename) === null) {
      focusedQid = q.id;
      return;
    }
  }
  if (sheetIndex < sheetFiles.length - 1) {
    gotoSheet(sheetIndex + 1);
  } else {
    logFn('✓ 最後の答案まで採点しました');
  }
}

async function gotoSheet(index) {
  sheetIndex = Math.min(sheetFiles.length - 1, Math.max(0, index));
  const filename = sheetFiles[sheetIndex];
  renderSheetHead();
  const img = el('sheet-image');
  const url = await sheetLoader.load(filename, filename);
  await new Promise((resolve) => {
    img.onload = resolve;
    img.src = url;
  });
  // 次の答案を先読みしておく（ページ送りを待たせない）
  if (sheetIndex + 1 < sheetFiles.length) {
    const next = sheetFiles[sheetIndex + 1];
    sheetLoader.load(next, next).catch(() => {});
  }
  // フォーカスはこの答案の最初の未採点問題へ
  const firstUnscored = sheetQuestions.find((q) => sheetScore(q.id, filename) === null);
  focusedQid = (firstUnscored ?? sheetQuestions[0])?.id ?? null;
  el('sheet-stage').scrollTop = 0;
  layoutSheet();
  renderSheetSide();
}

/** 次の「未採点が残っている」答案へ */
function jumpToUnfinishedSheet() {
  const n = sheetFiles.length;
  for (let k = 1; k <= n; k++) {
    const idx = (sheetIndex + k) % n;
    const f = sheetFiles[idx];
    if (sheetQuestions.some((q) => sheetScore(q.id, f) === null)) {
      gotoSheet(idx);
      return;
    }
  }
  logFn('✓ すべての答案が採点済みです');
}

// ── ズーム/パン ──
function zoomBy(factor, cx, cy) {
  const stage = el('sheet-stage');
  const img = el('sheet-image');
  const cur = img.clientWidth / img.naturalWidth;
  const next = Math.min(4, Math.max(0.2, cur * factor));
  // カーソル位置を保ったまま拡縮する
  const rx = (stage.scrollLeft + cx) / img.clientWidth;
  const ry = (stage.scrollTop + cy) / img.clientHeight;
  sheetZoom = next;
  layoutSheet();
  stage.scrollLeft = rx * img.clientWidth - cx;
  stage.scrollTop = ry * img.clientHeight - cy;
}

function wireSheetStage() {
  const stage = el('sheet-stage');
  stage.addEventListener('wheel', (e) => {
    if (!e.ctrlKey) return;          // 通常ホイールはスクロールのまま
    e.preventDefault();
    const rect = stage.getBoundingClientRect();
    zoomBy(e.deltaY < 0 ? 1.2 : 1 / 1.2,
           e.clientX - rect.left, e.clientY - rect.top);
  }, { passive: false });

  // ドラッグでパン（領域クリックと区別するため少し動いてから発動）
  let pan = null;
  stage.addEventListener('pointerdown', (e) => {
    if (e.button !== 0) return;
    pan = { x: e.clientX, y: e.clientY,
            left: stage.scrollLeft, top: stage.scrollTop, moved: false };
  });
  stage.addEventListener('pointermove', (e) => {
    if (!pan) return;
    const dx = e.clientX - pan.x;
    const dy = e.clientY - pan.y;
    if (!pan.moved && Math.hypot(dx, dy) > 4) {
      pan.moved = true;
      stage.setPointerCapture(e.pointerId);
      stage.classList.add('panning');
    }
    if (pan.moved) {
      stage.scrollLeft = pan.left - dx;
      stage.scrollTop = pan.top - dy;
    }
  });
  const endPan = () => { pan = null; stage.classList.remove('panning'); };
  stage.addEventListener('pointerup', endPan);
  stage.addEventListener('pointercancel', endPan);

  el('btn-zoom-fit').addEventListener('click', () => { sheetZoom = null; layoutSheet(); });
  el('btn-zoom-100').addEventListener('click', () => { sheetZoom = 1; layoutSheet(); });
}

// ── キーボード ──
function commitSheetDigits() {
  clearTimeout(sheetDigitTimer);
  sheetDigitTimer = null;
  if (sheetDigitBuf === '') return;
  const v = parseInt(sheetDigitBuf, 10);
  sheetDigitBuf = '';
  const q = sheetQuestions.find((x) => x.id === focusedQid);
  if (!q) return;
  if (v <= q.max_score) setSheetScore(q.id, v);
  else logFn(`❌ ${v} 点は満点（${q.max_score}点）を超えています`);
}

function sheetKeyHandler(e) {
  if (el('single-sheet-view').hidden) return;
  const tag = e.target.tagName;
  if (tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA') return;
  if (!sheetFiles.length) return;

  if (e.key >= '0' && e.key <= '9') {
    e.preventDefault();
    const q = sheetQuestions.find((x) => x.id === focusedQid);
    if (!q) return;
    if (q.max_score <= 9) {
      const v = parseInt(e.key, 10);
      if (v <= q.max_score) setSheetScore(q.id, v);
      else logFn(`❌ ${v} 点は満点（${q.max_score}点）を超えています`);
      return;
    }
    sheetDigitBuf += e.key;
    if (parseInt(sheetDigitBuf + '0', 10) > q.max_score) commitSheetDigits();
    else {
      clearTimeout(sheetDigitTimer);
      sheetDigitTimer = setTimeout(commitSheetDigits, 500);
    }
    return;
  }
  if (e.key === 'Enter') { e.preventDefault(); commitSheetDigits(); return; }
  if (e.key === 'Backspace' || e.key === 'Delete') {
    e.preventDefault();
    sheetDigitBuf = '';
    if (focusedQid) setSheetScore(focusedQid, null);
    return;
  }
  if (e.key === 'ArrowLeft') { e.preventDefault(); gotoSheet(sheetIndex - 1); return; }
  if (e.key === 'ArrowRight') { e.preventDefault(); gotoSheet(sheetIndex + 1); return; }
  if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
    e.preventDefault();
    const i = sheetQuestions.findIndex((q) => q.id === focusedQid);
    const d = e.key === 'ArrowUp' ? -1 : 1;
    const next = sheetQuestions[
      (i + d + sheetQuestions.length) % sheetQuestions.length];
    focusedQid = next?.id ?? focusedQid;
    layoutSheet();
    renderSheetSide();
  }
}

export async function openSingleSheet() {
  const listing = await call('list_sheet_files');
  sheetFiles = listing.files;
  const state = (await call('get_state')).state;
  sheetQuestions = state.descriptive.questions;
  // 全問題の得点表を一括で読み込み、以後はローカルで持つ
  // （1枚ごとに全問題を照会すると 100問×移動のたびに往復してしまう）
  sheetScores = {};
  for (const q of sheetQuestions) {
    const t = await call('list_descriptive_targets', q.id);
    sheetScores[q.id] = {};
    for (const item of t.items) {
      if (item.score !== null) sheetScores[q.id][item.filename] = item.score;
    }
  }
  sheetZoom = null;
  await showView('single-sheet-view');
  await gotoSheet(sheetIndex);
  logFn('一枚採点を開きました（←→で答案を移動、数字キーで採点）');
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
  el('btn-sheet-prev').addEventListener('click', () => gotoSheet(sheetIndex - 1));
  el('btn-sheet-next').addEventListener('click', () => gotoSheet(sheetIndex + 1));
  el('btn-sheet-unfinished').addEventListener('click', jumpToUnfinishedSheet);
  document.addEventListener('keydown', sheetKeyHandler);
  wireSheetStage();
}
