/**
 * handwriting.js — 答案ごとの確認・修正ビューの手書きコメント（webui/docs/handwriting-plan.md）。
 *
 * - ペン入力（pointerType=pen）は常時描画、マウスは「✎コメント」トグル（C）
 * - 筆跡は 00_Processing 実ピクセル座標のベクターで保持し、確定のたびに
 *   bridge へ保存。ズームは再描画（劣化なし）
 * - 消しゴムはストローク単位。戻す/やり直すはスナップショット方式
 * - 描画面（stage/canvas）は差し替え可能。1枚ずつビューは設問領域だけを
 *   表示するので、viewport（答案全体座標での表示範囲）を介して座標変換する。
 *   保存する基準寸法（natural）は常に答案全体。
 * - スタンプ（◎○✓波線）は通常のストロークとして生成して積むだけ。
 *   消しゴム・undo・焼き込みは既存経路がそのまま効く。
 */
import { call } from './bridge.js';

const COLORS = [
  { value: '#c73e2e', label: '朱' },
  { value: '#1d5fa8', label: '青' },
  { value: '#2b2825', label: '墨' },
];
const WIDTHS = [{ value: 2, label: '細' }, { value: 4, label: '太' }];
const STAMPS = [
  { id: 'circle2', label: '◎', title: 'よくできました' },
  { id: 'circle', label: '○', title: '良い' },
  { id: 'check', label: '✓', title: '確認' },
  { id: 'wave', label: '〰', title: '強調（波線）' },
];

let logFn = () => {};
let canvas = null;
let stage = null;
let filename = null;
let natural = { w: 595, h: 842 };   // 保存の基準寸法（常に答案全体）
let view = null;                     // 表示範囲 {x,y,w,h}（null=全体）
let strokes = [];
let undoStack = [];
let redoStack = [];
let drawing = null;          // 描画中のストローク
let commentMode = false;
let tool = { color: COLORS[0].value, width: WIDTHS[0].value,
             eraser: false, stamp: null };
const boundStages = new WeakSet();   // pointer リスナ二重登録防止

const el = (id) => document.getElementById(id);

// ── 描画 ─────────────────────────────────────────

function viewRect() {
  return view ?? { x: 0, y: 0, w: natural.w, h: natural.h };
}

function scaleFactor() {
  return canvas.width / viewRect().w;
}

function drawStroke(ctx, s, k) {
  const v = viewRect();
  ctx.strokeStyle = s.color;
  ctx.fillStyle = s.color;
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
  const pts = s.points;
  if (pts.length === 1) {
    const [x, y, p] = pts[0];
    const r = Math.max(0.6, (s.width * k * (0.5 + p * 0.8)) / 2);
    ctx.beginPath();
    ctx.arc((x - v.x) * k, (y - v.y) * k, r, 0, Math.PI * 2);
    ctx.fill();
    return;
  }
  for (let i = 1; i < pts.length; i++) {
    const [x1, y1, p1] = pts[i - 1];
    const [x2, y2, p2] = pts[i];
    ctx.lineWidth = Math.max(0.6, s.width * k * (0.5 + ((p1 + p2) / 2) * 0.8));
    ctx.beginPath();
    ctx.moveTo((x1 - v.x) * k, (y1 - v.y) * k);
    ctx.lineTo((x2 - v.x) * k, (y2 - v.y) * k);
    ctx.stroke();
  }
}

export function redrawHandwriting() {
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  const k = scaleFactor();
  for (const s of strokes) drawStroke(ctx, s, k);
  // 消しゴム中の drawing は points を持たない（S6: 描こうとすると例外→undo喪失）
  if (drawing && !drawing.eraser) drawStroke(ctx, drawing, k);
}

/** 画像の表示サイズ・実サイズに合わせてキャンバスを敷き直す。
 * naturalW/H は保存基準となる答案全体の寸法。viewport を渡すと
 * 「答案全体座標のうちこの範囲だけを表示している」ものとして座標変換する
 * （1枚ずつビューの設問クロップ用）。null なら全体表示。 */
export function layoutHandwriting(clientW, clientH, naturalW, naturalH,
                                  viewport = null) {
  if (!canvas) return;
  natural = { w: naturalW, h: naturalH };
  view = viewport;
  // 高DPI（Windows の 125/150% 表示など）でも筆跡がぼけないよう、
  // ビットマップは物理ピクセルで確保する（A1）。描画座標系は
  // canvas.width / natural.w のスケールなので追加の変換は不要
  const dpr = Math.min(3, window.devicePixelRatio || 1);
  canvas.width = Math.max(1, Math.round(clientW * dpr));
  canvas.height = Math.max(1, Math.round(clientH * dpr));
  canvas.style.width = `${clientW}px`;
  canvas.style.height = `${clientH}px`;
  redrawHandwriting();
}

// ── 保存・undo ───────────────────────────────────

async function save() {
  try {
    await call('set_handwriting', filename, natural.w, natural.h, strokes);
  } catch (e) {
    logFn(`❌ 筆跡の保存に失敗しました: ${e.message}`);
  }
}

function snapshot() {
  undoStack.push(JSON.stringify(strokes));
  if (undoStack.length > 50) undoStack.shift();
  redoStack = [];
  updateToolbar();
}

export function undoHandwriting() {
  if (!undoStack.length) return;
  redoStack.push(JSON.stringify(strokes));
  strokes = JSON.parse(undoStack.pop());
  redrawHandwriting();
  updateToolbar();
  save();
}

export function redoHandwriting() {
  if (!redoStack.length) return;
  undoStack.push(JSON.stringify(strokes));
  strokes = JSON.parse(redoStack.pop());
  redrawHandwriting();
  updateToolbar();
  save();
}

// ── 入力 ─────────────────────────────────────────

function toNatural(e) {
  const rect = canvas.getBoundingClientRect();
  const v = viewRect();
  return [
    v.x + ((e.clientX - rect.left) / rect.width) * v.w,
    v.y + ((e.clientY - rect.top) / rect.height) * v.h,
    e.pressure && e.pressure > 0 ? Math.min(1, e.pressure) : 0.5,
  ];
}

function shouldDraw(e) {
  if (e.button !== 0 && e.pointerType !== 'pen') return false;
  return e.pointerType === 'pen' || commentMode;
}

/** ストローク単位の消しゴム: ポインタ近傍を通るストロークを消す */
function eraseAt(e) {
  const [px, py] = toNatural(e);
  const threshold = 8 / scaleFactor();   // 画面上 8px 相当
  const before = strokes.length;
  strokes = strokes.filter((s) => !s.points.some(([x, y]) =>
    Math.hypot(x - px, y - py) < threshold + s.width));
  if (strokes.length !== before) redrawHandwriting();
  return strokes.length !== before;
}

/** スタンプ1個ぶんのストローク群（答案全体座標）を作る */
function stampStrokes(kind, cx, cy) {
  const r = 16;
  const mk = (pts) => ({ color: tool.color, width: tool.width,
                         points: pts.map(([x, y]) => [x, y, 0.6]) });
  const circle = (rad) => {
    const pts = [];
    for (let i = 0; i <= 24; i++) {
      const a = (i / 24) * Math.PI * 2;
      pts.push([cx + rad * Math.cos(a), cy + rad * Math.sin(a)]);
    }
    return pts;
  };
  switch (kind) {
    case 'circle': return [mk(circle(r))];
    case 'circle2': return [mk(circle(r)), mk(circle(r * 0.55))];
    case 'check':
      return [mk([[cx - r, cy], [cx - r * 0.25, cy + r * 0.75],
                  [cx + r, cy - r * 0.9]])];
    case 'wave': {
      const pts = [];
      for (let i = 0; i <= 24; i++) {
        const t = i / 24;
        pts.push([cx + (t - 0.5) * 4.4 * r,
                  cy + 0.5 * r * Math.sin(t * Math.PI * 3)]);
      }
      return [mk(pts)];
    }
    default: return [];
  }
}

function onPointerDown(e) {
  if (!shouldDraw(e)) return;
  e.stopPropagation();          // 描画中はパンを起動させない
  e.preventDefault();
  if (tool.stamp) {
    const [cx, cy] = toNatural(e);
    snapshot();
    strokes.push(...stampStrokes(tool.stamp, cx, cy));
    redrawHandwriting();
    updateToolbar();
    save();
    return;
  }
  canvas.setPointerCapture(e.pointerId);
  if (tool.eraser) {
    snapshot();
    drawing = { eraser: true, erased: eraseAt(e) };
    return;
  }
  snapshot();
  drawing = { color: tool.color, width: tool.width, points: [toNatural(e)] };
  redrawHandwriting();
}

function onPointerMove(e) {
  if (!drawing) return;
  e.stopPropagation();
  if (drawing.eraser) {
    drawing.erased = eraseAt(e) || drawing.erased;
    return;
  }
  drawing.points.push(toNatural(e));
  redrawHandwriting();
}

function onPointerUp(e) {
  if (!drawing) return;
  e.stopPropagation();
  if (drawing.eraser) {
    if (!drawing.erased) { undoStack.pop(); updateToolbar(); }  // 空振りは履歴に残さない
    drawing = null;
    save();
    return;
  }
  strokes.push({ color: drawing.color, width: drawing.width,
                 points: drawing.points });
  drawing = null;
  redrawHandwriting();
  updateToolbar();
  save();
}

// ── ツールバー ────────────────────────────────────

export function setCommentMode(on) {
  commentMode = on;
  el('btn-hw-toggle').classList.toggle('active', on);
  stage.classList.toggle('commenting', on);
}

function updateToolbar() {
  el('btn-hw-undo').disabled = undoStack.length === 0;
  el('btn-hw-redo').disabled = redoStack.length === 0;
  el('btn-hw-clear').disabled = strokes.length === 0;
  for (const b of document.querySelectorAll('#hw-toolbar .hw-color')) {
    b.classList.toggle('active', !tool.eraser && b.dataset.color === tool.color);
  }
  for (const b of document.querySelectorAll('#hw-toolbar .hw-stamp')) {
    b.classList.toggle('active', b.dataset.stamp === tool.stamp);
  }
  for (const b of document.querySelectorAll('#hw-toolbar .hw-width')) {
    b.classList.toggle('active', Number(b.dataset.width) === tool.width);
  }
  el('btn-hw-eraser').classList.toggle('active', tool.eraser);
}

/** 答案を切り替えたときに呼ぶ */
export async function loadHandwriting(fname) {
  filename = fname;
  strokes = [];
  undoStack = [];
  redoStack = [];
  drawing = null;
  try {
    const res = await call('get_handwriting', fname);
    strokes = res.strokes;
  } catch { /* 未保存なら空のまま */ }
  redrawHandwriting();
  updateToolbar();
}

/** 描画面を差し替える。答案ごとの確認＝#sheet-stage、1枚ずつ＝#one-stage。
 * hideClear: 全消去は「見えていない領域外の筆跡も消える」ため
 * 設問クロップ表示（1枚ずつ）では出さない。 */
export function attachHandwritingSurface(stageEl, canvasEl,
                                         { hideClear = false } = {}) {
  if (stage && stage !== stageEl) stage.classList.remove('commenting');
  stage = stageEl;
  canvas = canvasEl;
  if (!boundStages.has(stageEl)) {
    boundStages.add(stageEl);
    // capture 段で受けて、描画するときだけパン（bubble 段）を止める
    stageEl.addEventListener('pointerdown', onPointerDown, { capture: true });
    stageEl.addEventListener('pointermove', onPointerMove, { capture: true });
    stageEl.addEventListener('pointerup', onPointerUp, { capture: true });
    stageEl.addEventListener('pointercancel', onPointerUp, { capture: true });
  }
  el('btn-hw-clear').hidden = hideClear;
  stage.classList.toggle('commenting', commentMode);
}

export function wireHandwriting(log) {
  logFn = log;
  attachHandwritingSurface(el('sheet-stage'), el('hw-canvas'));

  el('btn-hw-toggle').addEventListener('click', () => setCommentMode(!commentMode));
  const colorWrap = el('hw-colors');
  for (const c of COLORS) {
    const b = document.createElement('button');
    b.className = 'btn hw-color';
    b.dataset.color = c.value;
    b.title = c.label;
    b.style.setProperty('--hw-color', c.value);
    b.addEventListener('click', () => {
      tool.color = c.value;
      tool.eraser = false;
      updateToolbar();
    });
    colorWrap.appendChild(b);
  }
  const widthWrap = el('hw-widths');
  for (const w of WIDTHS) {
    const b = document.createElement('button');
    b.className = 'btn hw-width';
    b.dataset.width = w.value;
    b.textContent = w.label;
    b.addEventListener('click', () => {
      tool.width = w.value;
      tool.eraser = false;
      updateToolbar();
    });
    widthWrap.appendChild(b);
  }
  const stampWrap = el('hw-stamps');
  for (const st of STAMPS) {
    const b = document.createElement('button');
    b.className = 'btn hw-stamp';
    b.dataset.stamp = st.id;
    b.textContent = st.label;
    b.title = `スタンプ: ${st.title}（クリックした場所に押します）`;
    b.addEventListener('click', () => {
      tool.stamp = (tool.stamp === st.id) ? null : st.id;
      tool.eraser = false;
      updateToolbar();
    });
    stampWrap.appendChild(b);
  }
  el('btn-hw-eraser').addEventListener('click', () => {
    tool.eraser = !tool.eraser;
    tool.stamp = null;
    updateToolbar();
  });
  el('btn-hw-undo').addEventListener('click', undoHandwriting);
  el('btn-hw-redo').addEventListener('click', redoHandwriting);
  el('btn-hw-clear').addEventListener('click', () => {
    if (!strokes.length) return;
    if (!confirm('この答案の手書きコメントをすべて消しますか？')) return;
    snapshot();
    strokes = [];
    redrawHandwriting();
    updateToolbar();
    save();
  });
  updateToolbar();
}

export function isCommentMode() { return commentMode; }
