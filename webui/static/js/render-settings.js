/**
 * render-settings.js — 「表示項目の設定」ダイアログ。
 *
 * 採点済み答案に何を描き込むかを選ぶ。変更はその場で bridge に保存され
 * （セッションにも残る）、次の「採点済み答案を生成」から反映される。
 * モードに応じて関係のない節は出さない。
 */
import { call } from './bridge.js';

let logFn = () => {};

// 節ごとの項目定義。key は constants.DEFAULT_RENDERING_SETTINGS と一致させる
const SECTIONS = [
  {
    title: '合計点の表示',
    modes: ['mark_only', 'mark_and_descriptive', 'descriptive_only'],
    items: [
      { key: 'total_show_max', label: '満点も表示する', type: 'bool',
        note: 'オフにすると「得点：82 / 100」ではなく「得点：82」になります' },
      { key: 'total_show_aspects', label: '観点別の行を表示する', type: 'bool',
        note: '「(観点①：40/50 …)」の行' },
    ],
  },
  {
    title: 'マーク設問への描き込み',
    modes: ['mark_only', 'mark_and_descriptive'],
    items: [
      { key: 'show_ox_mark', label: '○×△マーク', type: 'bool' },
      { key: 'show_score', label: '得点', type: 'bool' },
      { key: 'show_aspect', label: '観点番号', type: 'bool' },
      { key: 'show_correct_answer', label: '正答の選択肢番号', type: 'bool' },
      { key: 'show_all_correct_star', label: '全員正解（特例）の設問に★', type: 'bool',
        note: '不適切問題などを正答データの特例列で「全員正解」にした設問の印' },
      { key: 'mark_result_bg_white', label: '文字の背景を白塗りする', type: 'bool',
        note: '印字と重なって読みづらいときに' },
      { key: 'mark_result_offset', label: '描き込み位置のオフセット', type: 'offset',
        note: '0=標準 / 正で右へ・負で左へ（セル単位）。下のプレビューで確認' },
    ],
  },
  {
    title: '記述設問への描き込み',
    modes: ['mark_and_descriptive', 'descriptive_only'],
    items: [
      { key: 'descriptive_show_mark', label: '○×△マーク', type: 'bool' },
      { key: 'descriptive_show_score', label: '得点', type: 'bool' },
      { key: 'descriptive_show_aspect', label: '観点番号', type: 'bool' },
      { key: 'descriptive_opacity', label: '描き込みの濃さ', type: 'number',
        step: 0.05, min: 0, max: 1, note: '0=透明 〜 1=不透明' },
    ],
  },
];

const el = (id) => document.getElementById(id);

async function applyChange(key, value) {
  try {
    await call('set_rendering_settings', { [key]: value });
    logFn('✓ 表示項目の設定を保存しました');
    refreshPreview();
  } catch (e) {
    logFn(`❌ ${e.message}`);
  }
}

/** マーク描き込みのライブプレビュー（実際の描画関数の出力） */
let previewAvailable = false;
async function refreshPreview() {
  const box = el('rs-preview');
  if (!box || box.hidden) return;
  try {
    const res = await call('get_render_preview');
    box.querySelector('img').src = res.data_url;
    box.querySelector('.rs-preview-note').textContent = res.sample_note;
    previewAvailable = true;
  } catch (e) {
    if (!previewAvailable) {
      box.querySelector('.rs-preview-note').textContent = e.message;
      box.querySelector('img').removeAttribute('src');
    }
  }
}

function buildRow(item, settings) {
  if (item.type === 'offset') return buildOffsetRow(item, settings);
  const row = document.createElement('label');
  row.className = 'rs-row';
  if (item.type === 'bool') {
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = Boolean(settings[item.key]);
    cb.addEventListener('change', () => applyChange(item.key, cb.checked));
    row.append(cb, item.label);
  } else {
    const input = document.createElement('input');
    input.type = 'number';
    input.value = settings[item.key];
    if (item.step !== undefined) input.step = item.step;
    if (item.min !== undefined) input.min = item.min;
    if (item.max !== undefined) input.max = item.max;
    input.addEventListener('change', () => applyChange(item.key, input.value));
    row.append(item.label, input);
  }
  if (item.note) {
    const note = document.createElement('span');
    note.className = 'rs-note';
    note.textContent = item.note;
    row.appendChild(note);
  }
  return row;
}

function buildOffsetRow(item, settings) {
  const row = document.createElement('div');
  row.className = 'rs-row';
  row.append(item.label);
  let value = Number(settings[item.key]) || 0;
  const label = document.createElement('span');
  label.className = 'rs-offset-value score-mark';
  const show = () => { label.textContent = value.toFixed(1); };
  show();
  const stepBtn = (text, delta) => {
    const b = document.createElement('button');
    b.className = 'btn rs-step';
    b.textContent = text;
    b.addEventListener('click', () => {
      value = Math.round((value + delta) * 10) / 10;
      show();
      applyChange(item.key, value);
    });
    return b;
  };
  row.append(stepBtn('◀◀', -1), stepBtn('◀', -0.5), stepBtn('◀·', -0.1),
             label,
             stepBtn('·▶', 0.1), stepBtn('▶', 0.5), stepBtn('▶▶', 1));
  if (item.note) {
    const note = document.createElement('span');
    note.className = 'rs-note';
    note.textContent = item.note;
    row.appendChild(note);
  }
  return row;
}

export async function openRenderSettings(appMode) {
  const res = await call('get_rendering_settings');
  const body = el('rs-body');
  body.replaceChildren(...SECTIONS
    .filter((sec) => sec.modes.includes(appMode))
    .map((sec) => {
      const div = document.createElement('div');
      div.className = 'rs-section';
      const h = document.createElement('h3');
      h.textContent = sec.title;
      div.appendChild(h);
      for (const item of sec.items) div.appendChild(buildRow(item, res.settings));
      return div;
    }));
  if (appMode !== 'descriptive_only') {
    const prev = document.createElement('div');
    prev.className = 'rs-section';
    prev.id = 'rs-preview';
    prev.innerHTML = '<h3>プレビュー（実際の描画関数の出力）</h3>' +
      '<img class="rs-preview-img" alt="描き込みプレビュー">' +
      '<p class="rs-preview-note hint"></p>';
    body.appendChild(prev);
  }
  previewAvailable = false;
  el('rs-overlay').hidden = false;
  refreshPreview();
}

export function wireRenderSettings(log, getMode) {
  logFn = log;
  el('btn-render-settings').addEventListener('click', () =>
    openRenderSettings(getMode()).catch((e) => { log(`❌ ${e.message}`); }));
  el('btn-rs-close').addEventListener('click', () => { el('rs-overlay').hidden = true; });
  el('rs-overlay').addEventListener('click', (e) => {
    if (e.target === el('rs-overlay')) el('rs-overlay').hidden = true;
  });
}
