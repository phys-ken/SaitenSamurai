/**
 * region-picker.js — 画像上の矩形ドラッグ選択（汎用コンポーネント）。
 *
 * 合計点位置・氏名欄・記述問題の領域設定で共用する。表示は CSS で
 * 縮小されるため、選択結果は画像の natural サイズのピクセル座標へ
 * 変換して返す（tk 版と同じ「00_Processing 画像のピクセル座標」系）。
 *
 * 将来のコメントモードはこの上位版（自由描画レイヤ）として同じ
 * オーバーレイ構造に載せる想定。
 */

/**
 * @param {string} dataUrl 画像
 * @param {{title?: string, existing?: number[]}} opts
 * @returns {Promise<number[]|null>} [x1,y1,x2,y2]（natural座標）/ キャンセルは null
 */
export function pickRegion(dataUrl, opts = {}) {
  return new Promise((resolve) => {
    const overlay = document.createElement('div');
    overlay.className = 'picker-overlay';
    overlay.innerHTML = `
      <div class="picker-dialog">
        <div class="picker-head">
          <span>${opts.title ?? '領域をドラッグで指定してください'}</span>
          <button class="btn" data-act="cancel">キャンセル</button>
          <button class="btn btn-primary" data-act="ok" disabled>この領域に決定</button>
        </div>
        <div class="picker-stage">
          <img alt="answer sheet">
          <div class="picker-rect" hidden></div>
        </div>
      </div>`;
    document.body.appendChild(overlay);

    const img = overlay.querySelector('img');
    const rectEl = overlay.querySelector('.picker-rect');
    const okBtn = overlay.querySelector('[data-act="ok"]');
    let sel = null;     // natural座標 [x1,y1,x2,y2]

    function toNatural(ev) {
      const r = img.getBoundingClientRect();
      const sx = img.naturalWidth / r.width;
      const sy = img.naturalHeight / r.height;
      const x = Math.min(Math.max(ev.clientX - r.left, 0), r.width) * sx;
      const y = Math.min(Math.max(ev.clientY - r.top, 0), r.height) * sy;
      return [Math.round(x), Math.round(y)];
    }

    function showRect([x1, y1, x2, y2]) {
      const r = img.getBoundingClientRect();
      const stage = img.parentElement.getBoundingClientRect();
      const sx = r.width / img.naturalWidth;
      const sy = r.height / img.naturalHeight;
      Object.assign(rectEl.style, {
        left: `${r.left - stage.left + x1 * sx}px`,
        top: `${r.top - stage.top + y1 * sy}px`,
        width: `${(x2 - x1) * sx}px`,
        height: `${(y2 - y1) * sy}px`,
      });
      rectEl.hidden = false;
    }

    let dragStart = null;
    img.addEventListener('pointerdown', (ev) => {
      dragStart = toNatural(ev);
      img.setPointerCapture(ev.pointerId);
    });
    img.addEventListener('pointermove', (ev) => {
      if (!dragStart) return;
      const cur = toNatural(ev);
      sel = [Math.min(dragStart[0], cur[0]), Math.min(dragStart[1], cur[1]),
             Math.max(dragStart[0], cur[0]), Math.max(dragStart[1], cur[1])];
      showRect(sel);
    });
    img.addEventListener('pointerup', () => {
      dragStart = null;
      okBtn.disabled = !sel || (sel[2] - sel[0] < 4) || (sel[3] - sel[1] < 4);
    });

    function close(result) {
      overlay.remove();
      resolve(result);
    }
    overlay.querySelector('[data-act="cancel"]').addEventListener('click', () => close(null));
    okBtn.addEventListener('click', () => close(sel));
    img.addEventListener('load', () => {
      if (opts.existing) { sel = opts.existing; showRect(sel); okBtn.disabled = false; }
    });
    img.src = dataUrl;
  });
}
