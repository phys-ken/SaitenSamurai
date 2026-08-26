/**
 * transitions.js — 画面遷移の共通ヘルパー。
 * View Transitions API があれば軽いクロスフェード、なければ即時切替。
 * reduced-motion 指定時も即時切替にする。
 *
 * 戻り値: update が DOM に反映された後に解決する Promise。
 * 遷移後にレイアウト計測（画像サイズ・グリッド幅など）をする場合は
 * 必ず await すること — 反映前は clientWidth が 0 のままになる。
 */
export function withTransition(update) {
  if (document.startViewTransition &&
      !matchMedia('(prefers-reduced-motion: reduce)').matches) {
    const t = document.startViewTransition(update);
    // アニメーション中は ::view-transition がポインタを奪うため、
    // 終了が分かる印を付けておく（テスト・連打対策）
    document.documentElement.dataset.vtBusy = '1';
    t.finished.finally(() => {
      delete document.documentElement.dataset.vtBusy;
    });
    return t.updateCallbackDone;
  }
  update();
  return Promise.resolve();
}
