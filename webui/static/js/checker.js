/**
 * checker.js — マークチェックビュー。
 * bridge の CheckerMixin API（open/get_entries/get_entry_image/
 * set_correction/apply_corrections/close）に対応する画面制御。
 *
 * 400人規模ではエラー・値ごとの件数が数千になり得るため、一覧は
 * 仮想スクロール（vlist.js）で表示し、画像は近づいたときだけ取得する。
 * キーボード: 記号キー1打で訂正＋次のカードへ。詳しくは画面のヒント行。
 */
import { call } from './bridge.js';
import { withTransition } from './transitions.js';
import { createVirtualGrid, createImageLoader } from './vlist.js';
import { openLightbox } from './lightbox.js';

let view = { category: '__errors__', sort: 'whiteness' };
let entries = [];          // 現在のタブの全件（メタデータのみ）
let grid = null;
let cursor = 0;
let cardWidth = 230;
let logFn = () => {};

const CARD_IMG_H = 72;
const CARD_H = CARD_IMG_H + 24 + 34 + 22;

const entryImageLoader = createImageLoader(
  (id) => call('get_entry_image', id).then((r) => r.data_url),
  { concurrency: 6, cacheSize: 400 });

function el(id) { return document.getElementById(id); }

export function isOpen() { return !el('checker-view').hidden; }

export async function openChecker(log) {
  logFn = log;
  const res = await call('open_mark_checker');
  const c = res.state.checker;
  // 要確認が0件なら値タブから開く（sort は保持する — S11）
  view.category = c.error_count > 0 ? '__errors__' : null;
  el('checker-sort').value = view.sort;
  cursor = 0;
  await withTransition(() => {
    document.querySelectorAll('main > .panel').forEach((p) => { p.hidden = true; });
    el('checker-view').hidden = false;
  });
  renderHead(c);
  renderTabs(c);
  await renderGrid();
  logFn(`マークチェックを開きました（全 ${c.total} 件 / 要確認 ${c.error_count} 件）`);
}

export async function closeChecker() {
  const st = (await call('get_state')).state;
  const pending = st.checker?.corrected ?? 0;
  if (pending > 0 &&
      !confirm(`「訂正をxlsxに反映」していない訂正が ${pending} 件あります。\n` +
               '訂正は保存済みなので、次回開いたときに続きから反映できます。\n' +
               'このまま閉じますか？')) {
    return;
  }
  await call('close_mark_checker');
  grid?.destroy();
  grid = null;
  entryImageLoader.clear();
  withTransition(() => {
    el('checker-view').hidden = true;
    document.querySelectorAll('main > .panel').forEach((p) => { p.hidden = false; });
  });
}

function renderHead(c) {
  el('checker-summary').textContent =
    `全 ${c.total} 件 / 要確認 ${c.error_count} 件 / 訂正 ${c.corrected} 件`;
  el('btn-apply-corrections').disabled = c.corrected === 0;
}

function renderTabs(c) {
  const tabs = [{ name: '__errors__', label: `⚠ 要確認 (${c.error_count})`, error: true },
                { name: null, label: `すべて (${c.total})`, error: false }];
  for (const cat of c.categories) {
    if (!cat.is_error) tabs.push({ name: cat.name, label: `${cat.name} (${cat.count})`, error: false });
  }
  el('checker-tabs').replaceChildren(...tabs.map((t) => {
    const b = document.createElement('button');
    b.className = 'tab' + (t.error ? ' error-cat' : '') +
      (t.name === view.category ? ' active' : '');
    b.textContent = t.label;
    b.addEventListener('click', async () => {
      view.category = t.name;
      cursor = 0;
      renderTabs(c);
      await renderGrid();
    });
    return b;
  }));
}

function updateCursorClasses() {
  if (!grid) return;
  document.querySelectorAll('#checker-grid .entry-card.cursor')
    .forEach((n) => n.classList.remove('cursor'));
  grid.itemAt(cursor)?.classList.add('cursor');
}

async function setCorrection(i, value) {
  const item = entries[i];
  try {
    const res = await call('set_correction', item.id, value);
    item.after = res.entry.after;
    renderHead(res.state.checker);
    grid.refresh();
    updateCursorClasses();
    return true;
  } catch (e) {
    logFn(`❌ ${e.message}`);
    return false;
  }
}

function entryCard(i) {
  const item = entries[i];
  const card = document.createElement('div');
  card.className = 'entry-card' +
    (item.error_type ? ' error' : '') + (item.after ? ' corrected' : '') +
    (i === cursor ? ' cursor' : '');
  card.dataset.entryId = item.id;
  card.dataset.index = i;

  const img = document.createElement('img');
  img.className = 'crop-img';
  img.style.height = `${CARD_IMG_H}px`;
  img.alt = `${item.filename} 問${item.question_no}`;
  card.appendChild(img);
  entryImageLoader.load(`e${item.id}`, item.id)
    .then((url) => { img.src = url; })
    .catch(() => { img.alt = '画像なし'; });
  img.addEventListener('click', async () => {
    cursor = i;
    updateCursorClasses();
    try {
      const url = await entryImageLoader.load(`e${item.id}`, item.id);
      openLightbox(url, `${item.filename} 問${item.question_no}`);
    } catch { /* 画像なしはカード表示のまま */ }
  });

  const meta = document.createElement('div');
  meta.className = 'entry-meta';
  meta.innerHTML = `<span>${item.filename} / 問${item.question_no}</span>` +
    `<span>${item.category}</span>`;
  card.appendChild(meta);

  const fix = document.createElement('div');
  fix.className = 'entry-fix';
  const before = item.before === '' ? '（無マーク）' : item.before;
  fix.innerHTML = `<span class="entry-before">${before}</span> →`;
  const input = document.createElement('input');
  input.value = item.after;
  input.placeholder = '訂正';
  input.addEventListener('focus', () => { cursor = i; updateCursorClasses(); });
  input.addEventListener('change', async () => {
    const ok = await setCorrection(i, input.value);
    if (!ok) {
      input.value = item.after;
      input.focus();
    }
  });
  fix.appendChild(input);
  card.appendChild(fix);
  card.addEventListener('click', (e) => {
    if (e.target.tagName !== 'INPUT') { cursor = i; updateCursorClasses(); }
  });
  return card;
}

async function renderGrid() {
  // メタデータは全件を一度に取得し（数千件でも軽い）、以後の描画は仮想化に任せる
  const res = await call('get_checker_entries', view.category, 0, 1000000);
  entries = res.items;
  if (view.sort === 'whiteness') {
    // 白い（=マークが薄い）順。tk 版の既定と同じで、怪しいものが先頭に来る
    entries.sort((a, b) => (b.whiteness ?? 0) - (a.whiteness ?? 0));
  } else {
    entries.sort((a, b) => a.filename.localeCompare(b.filename) ||
                           a.question_no - b.question_no);
  }
  el('btn-batch-minus1').hidden = view.category !== 'ノーマーク';
  cursor = Math.min(cursor, Math.max(0, entries.length - 1));
  grid?.destroy();
  grid = createVirtualGrid(el('checker-grid'), {
    count: entries.length,
    itemMinWidth: cardWidth,
    itemHeight: CARD_H,
    gap: 10,
    renderItem: entryCard,
  });
  updateCursorClasses();
}

function advance() {
  if (cursor < entries.length - 1) {
    cursor++;
    grid.scrollToIndex(cursor);
    updateCursorClasses();
  } else {
    logFn('✓ このタブの最後のマークです');
  }
}

function checkerKeyHandler(e) {
  if (el('checker-view').hidden) return;
  const tag = e.target.tagName;
  if (tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA') return;
  if (!grid || !entries.length) return;

  if (/^[0-9a-dA-D-]$/.test(e.key)) {
    e.preventDefault();
    setCorrection(cursor, e.key).then((ok) => { if (ok) advance(); });
    return;
  }
  if (e.key === 'Backspace' || e.key === 'Delete') {
    e.preventDefault();
    setCorrection(cursor, '');
    return;
  }
  if (e.key === 'Enter') {          // -1 などの特殊値は入力欄で
    e.preventDefault();
    grid.itemAt(cursor)?.querySelector('input')?.focus();
    return;
  }
  const move = { ArrowLeft: -1, ArrowRight: 1,
                 ArrowUp: -grid.columns(), ArrowDown: grid.columns() }[e.key];
  if (move !== undefined) {
    e.preventDefault();
    cursor = Math.min(entries.length - 1, Math.max(0, cursor + move));
    grid.scrollToIndex(cursor);
    updateCursorClasses();
  }
}

export function wireChecker(log) {
  logFn = log;
  el('btn-close-checker').addEventListener('click', closeChecker);
  el('checker-sort').addEventListener('change', async (ev) => {
    view.sort = ev.target.value;
    cursor = 0;
    await renderGrid();
  });
  el('checker-size').addEventListener('change', async (ev) => {
    cardWidth = parseInt(ev.target.value, 10);
    await renderGrid();
  });
  el('btn-batch-minus1').addEventListener('click', async () => {
    if (!confirm('未訂正のノーマーク全件を「無効回答(-1)」に設定します。よろしいですか？')) return;
    try {
      const res = await call('batch_correct_no_mark');
      log(`✓ ノーマーク ${res.applied} 件を -1 に設定しました`);
      renderHead(res.state.checker);
      await renderGrid();
    } catch (e) { log(`❌ ${e.message}`); }
  });
  document.addEventListener('keydown', checkerKeyHandler);
  el('btn-apply-corrections').addEventListener('click', async () => {
    try {
      const res = await call('apply_corrections');
      log(`訂正 ${res.applied} 件を xlsx に反映しました（バックアップ: ${res.backup}）`);
      const c = res.state.checker;
      renderHead(c);
      renderTabs(c);
      cursor = 0;
      await renderGrid();
    } catch (e) {
      log(`❌ ${e.message}`);
      alert(e.message);
    }
  });
}
