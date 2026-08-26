/** lightbox.js — 画像のその場拡大（クリック/Escで閉じる）。共通部品 */
export function openLightbox(src, alt) {
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
