/**
 * transitions.js — 画面遷移の共通ヘルパー。
 * View Transitions API があれば軽いクロスフェード、なければ即時切替。
 * reduced-motion 指定時も即時切替にする。
 */
export function withTransition(update) {
  if (document.startViewTransition &&
      !matchMedia('(prefers-reduced-motion: reduce)').matches) {
    document.startViewTransition(update);
  } else {
    update();
  }
}
