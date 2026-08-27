/**
 * handwriting.js — 答案ごとの確認・修正ビューの手書きコメント（webui/docs/handwriting-plan.md）。
 *
 * - ペン入力（pointerType=pen）は常時描画、マウスは「✎コメント」トグル（C）
 * - 筆跡は 00_Processing 実ピクセル座標のベクターで保持し、確定のたびに
 *   bridge へ保存。ズームは再描画（劣化なし）
 * - 消しゴムはストローク単位。戻す/やり直すはスナップショット方式
 */
import { call } from './bridge.js';

const COLORS = [
  { value: '#c73e2e', label: '朱' },
  { value: '#1d5fa8', label: '青' },
  { value: '#2b2825', label: '墨' },
];
const WIDTHS = [{ value: 2, label: '細' }, { value: 4, label: '太' }];

let logFn = () => {};
let canvas = null;
let stage = null;
let filename = null;
let natural = { w: 595, h: 842 };
let strokes = [];
let undoStack = [];
let redoStack = [];
let drawing = null;          // 描画中のストローク
let commentMode = false;
let tool = { color: COLORS[0].value, width: WIDTHS[0].value, eraser: false };

const el = (id) => document.getElementById(id);

// ── 描画 ─────────────────────────────────────────

function scaleFactor() {
  return canvas.width / natural.w;
}

function drawStroke(ctx, s, k) {
  ctx.strokeStyle = s.color;
  ctx.fillStyle = s.color;
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
  const pts = s.points;
  if (pts.length === 1) {
    const [x, y, p] = pts[0];
    const r = Math.max(0.6, (s.width * k * (0.5 + p * 0.8)) / 2);
    ctx.beginPath();
    ctx.arc(x * k, y * k, r, 0, Math.PI * 2);
    ctx.fill();
    return;
  }
  for (let i = 1; i < pts.length; i++) {
    const [x1, y1, p1] = pts[i - 1];
    const [x2, y2, p2] = pts[i];
    ctx.lineWidth = Math.max(0.6, s.width * k * (0.5 + ((p1 + p2) / 2) * 0.8));
    ctx.beginPath();
    ctx.moveTo(x1 * k, y1 * k);
    ctx.lineTo(x2 * k, y2 * k);
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

/** layoutSheet から呼ぶ: 画像の表示サイズ・実サイズに合わせて敷き直す */
export function layoutHandwriting(clientW, clientH, naturalW, naturalH) {
  if (!canvas) return;
  natural = { w: naturalW, h: naturalH };
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
  return [
    ((e.clientX - rect.left) / rect.width) * natural.w,
    ((e.clientY - rect.top) / rect.height) * natural.h,
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

function onPointerDown(e) {
  if (!shouldDraw(e)) return;
  e.stopPropagation();          // 描画中はパンを起動させない
  e.preventDefault();
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

export function wireHandwriting(log) {
  logFn = log;
  canvas = el('hw-canvas');
  stage = el('sheet-stage');

  // capture 段で受けて、描画するときだけパン（bubble 段）を止める
  stage.addEventListener('pointerdown', onPointerDown, { capture: true });
  stage.addEventListener('pointermove', onPointerMove, { capture: true });
  stage.addEventListener('pointerup', onPointerUp, { capture: true });
  stage.addEventListener('pointercancel', onPointerUp, { capture: true });

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
  el('btn-hw-eraser').addEventListener('click', () => {
    tool.eraser = !tool.eraser;
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
