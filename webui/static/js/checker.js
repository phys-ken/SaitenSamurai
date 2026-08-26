/**
 * checker.js — マークチェックビュー。
 * bridge の CheckerMixin API（open/get_entries/get_entry_image/
 * set_correction/apply_corrections/close）に対応する画面制御。
 */
import { call } from './bridge.js';

const PAGE_SIZE = 24;
let view = { category: '__errors__', page: 0, total: 0 };
let logFn = () => {};

function el(id) { return document.getElementById(id); }

export function isOpen() { return !el('checker-view').hidden; }

export async function openChecker(log) {
  logFn = log;
  const res = await call('open_mark_checker');
  const c = res.state.checker;
  // 要確認が0件なら値タブから開く
  view = { category: c.error_count > 0 ? '__errors__' : null, page: 0, total: 0 };
  document.querySelectorAll('main > .panel').forEach((p) => { p.hidden = true; });
  el('checker-view').hidden = false;
  renderHead(c);
  renderTabs(c);
  await renderGrid();
  logFn(`🔍 マークチェックを開きました（全 ${c.total} 件 / 要確認 ${c.error_count} 件）`);
}

export async function closeChecker() {
  await call('close_mark_checker');
  el('checker-view').hidden = true;
  document.querySelectorAll('main > .panel').forEach((p) => { p.hidden = false; });
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
      view.page = 0;
      renderTabs(c);
      await renderGrid();
    });
    return b;
  }));
}

function entryCard(item) {
  const card = document.createElement('div');
  card.className = 'entry-card' +
    (item.error_type ? ' error' : '') + (item.after ? ' corrected' : '');
  card.dataset.entryId = item.id;

  const img = document.createElement('img');
  img.alt = `${item.filename} 問${item.question_no}`;
  card.appendChild(img);
  call('get_entry_image', item.id)
    .then((r) => { img.src = r.data_url; })
    .catch(() => { img.alt = '画像なし'; });

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
  input.addEventListener('change', async () => {
    try {
      const res = await call('set_correction', item.id, input.value);
      item.after = res.entry.after;
      input.value = res.entry.after;
      card.classList.toggle('corrected', Boolean(res.entry.after));
      renderHead(res.state.checker);
    } catch (e) {
      logFn(`❌ ${e.message}`);
      input.value = item.after;
      input.focus();
    }
  });
  fix.appendChild(input);
  card.appendChild(fix);
  return card;
}

async function renderGrid() {
  const res = await call('get_checker_entries', view.category, view.page, PAGE_SIZE);
  view.total = res.total;
  el('checker-grid').replaceChildren(...res.items.map(entryCard));
  const pages = Math.max(1, Math.ceil(res.total / PAGE_SIZE));
  el('pager-info').textContent = `${view.page + 1} / ${pages} ページ（${res.total} 件）`;
  el('pager-prev').disabled = view.page === 0;
  el('pager-next').disabled = view.page >= pages - 1;
}

export function wireChecker(log) {
  logFn = log;
  el('btn-close-checker').addEventListener('click', closeChecker);
  el('pager-prev').addEventListener('click', async () => { view.page--; await renderGrid(); });
  el('pager-next').addEventListener('click', async () => { view.page++; await renderGrid(); });
  el('btn-apply-corrections').addEventListener('click', async () => {
    try {
      const res = await call('apply_corrections');
      log(`💾 訂正 ${res.applied} 件を xlsx に反映しました（バックアップ: ${res.backup}）`);
      const c = res.state.checker;
      renderHead(c);
      renderTabs(c);
      view.page = 0;
      await renderGrid();
    } catch (e) {
      log(`❌ ${e.message}`);
      alert(e.message);
    }
  });
}
