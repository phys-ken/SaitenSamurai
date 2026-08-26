/**
 * vlist.js — 依存なしの仮想グリッド。
 *
 * 400人×100問でも固まらないための基盤。スクロール領域に「見えている行
 * ＋前後 overscan 行」だけを DOM に置き、残りはスペーサーの高さで表現する。
 * カードの高さは一定（呼び出し側が問題ごとに計算して渡す）。
 *
 * 使い方:
 *   const grid = createVirtualGrid(viewportEl, {
 *     count, itemMinWidth, itemHeight, gap,
 *     renderItem: (index) => HTMLElement,   // 毎回作り直してよい（軽量前提）
 *   });
 *   grid.refresh();          // データ変更後の再描画（DOM上のカードのみ再生成）
 *   grid.scrollToIndex(i);   // i 番目が見える位置までスクロール
 *   grid.itemAt(i);          // DOM に載っていればその要素、なければ null
 *   grid.columns();          // 現在の列数（キーボードの上下移動用）
 *   grid.destroy();
 */
export function createVirtualGrid(viewport, opts) {
  const { itemMinWidth, itemHeight, gap = 10, renderItem, overscan = 2 } = opts;
  let count = opts.count;

  const spacer = document.createElement('div');
  spacer.className = 'vgrid-spacer';
  viewport.replaceChildren(spacer);

  let cols = 1;
  let itemW = itemMinWidth;
  const rowH = () => itemHeight + gap;
  const mounted = new Map();   // index -> element

  function layout() {
    const w = spacer.clientWidth;
    cols = Math.max(1, Math.floor((w + gap) / (itemMinWidth + gap)));
    itemW = (w - gap * (cols - 1)) / cols;
    const rows = Math.ceil(count / cols);
    spacer.style.height = `${Math.max(0, rows * rowH() - gap)}px`;
  }

  function visibleRange() {
    const top = viewport.scrollTop;
    const h = viewport.clientHeight;
    const r0 = Math.max(0, Math.floor(top / rowH()) - overscan);
    const r1 = Math.ceil((top + h) / rowH()) + overscan;
    return [r0 * cols, Math.min(count, r1 * cols)];
  }

  function place(el, i) {
    el.style.position = 'absolute';
    el.style.width = `${itemW}px`;
    el.style.height = `${itemHeight}px`;
    el.style.left = `${(i % cols) * (itemW + gap)}px`;
    el.style.top = `${Math.floor(i / cols) * rowH()}px`;
  }

  function render() {
    const [i0, i1] = visibleRange();
    for (const [i, el] of mounted) {
      if (i < i0 || i >= i1) {
        el.remove();
        mounted.delete(i);
      }
    }
    for (let i = i0; i < i1; i++) {
      if (!mounted.has(i)) {
        const el = renderItem(i);
        place(el, i);
        spacer.appendChild(el);
        mounted.set(i, el);
      } else {
        place(mounted.get(i), i);   // 列数変化に追従
      }
    }
  }

  function refresh() {
    for (const [, el] of mounted) el.remove();
    mounted.clear();
    layout();
    render();
  }

  function scrollToIndex(i) {
    const row = Math.floor(i / cols);
    const top = row * rowH();
    const bottom = top + itemHeight;
    if (top < viewport.scrollTop) {
      viewport.scrollTop = top;
    } else if (bottom > viewport.scrollTop + viewport.clientHeight) {
      viewport.scrollTop = bottom - viewport.clientHeight;
    }
    render();
  }

  const onScroll = () => render();
  viewport.addEventListener('scroll', onScroll, { passive: true });
  const ro = new ResizeObserver(() => { layout(); render(); });
  ro.observe(viewport);

  layout();
  render();

  return {
    refresh,
    render,
    scrollToIndex,
    itemAt: (i) => mounted.get(i) ?? null,
    columns: () => cols,
    setCount(n) { count = n; refresh(); },
    destroy() {
      viewport.removeEventListener('scroll', onScroll);
      ro.disconnect();
      for (const [, el] of mounted) el.remove();
      mounted.clear();
      spacer.remove();
    },
  };
}

/**
 * 画像の遅延取得キュー: 同時取得数を絞り、取得済みは LRU キャッシュから返す。
 * bridge 越しの get_*_image 系はプロセス間往復なので、無制限に投げない。
 */
export function createImageLoader(fetcher, { concurrency = 6, cacheSize = 400 } = {}) {
  const cache = new Map();     // key -> data_url（挿入順 = LRU近似）
  const pending = new Map();   // key -> Promise
  const queue = [];
  let active = 0;

  function pump() {
    while (active < concurrency && queue.length) {
      const job = queue.shift();
      active++;
      fetcher(...job.args)
        .then((url) => {
          if (cache.size >= cacheSize) {
            cache.delete(cache.keys().next().value);
          }
          cache.set(job.key, url);
          job.resolve(url);
        })
        .catch(job.reject)
        .finally(() => {
          active--;
          pending.delete(job.key);
          pump();
        });
    }
  }

  return {
    load(key, ...args) {
      if (cache.has(key)) {
        const url = cache.get(key);
        cache.delete(key);       // 触れたものを末尾へ（LRU）
        cache.set(key, url);
        return Promise.resolve(url);
      }
      if (pending.has(key)) return pending.get(key);
      const p = new Promise((resolve, reject) => {
        queue.push({ key, args, resolve, reject });
      });
      pending.set(key, p);
      pump();
      return p;
    },
    clear() { cache.clear(); },
  };
}
