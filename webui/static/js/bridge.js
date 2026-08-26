/**
 * bridge.js — Python API への唯一の入口。
 *
 * - 本番: pywebview が window.pywebview.api を注入する（pywebviewready イベント）
 * - L2テスト: Playwright が window.__mockApi を注入する（pywebview 不要）
 * どちらでも同じ `call(name, ...args)` で呼べるように吸収する。
 */

function realApiReady() {
  return new Promise((resolve) => {
    if (window.pywebview?.api) { resolve(window.pywebview.api); return; }
    window.addEventListener('pywebviewready', () => resolve(window.pywebview.api), { once: true });
  });
}

let apiPromise = null;

export function getApi() {
  if (window.__mockApi) return Promise.resolve(window.__mockApi);
  if (!apiPromise) apiPromise = realApiReady();
  return apiPromise;
}

/** Python 側 Bridge のメソッドを呼ぶ。{ok:false} はここで例外化して1箇所に集約 */
export async function call(name, ...args) {
  const api = await getApi();
  const res = await api[name](...args);
  if (!res || res.ok !== true) {
    throw new Error(res?.error ?? `API ${name} が不正な応答を返しました`);
  }
  return res;
}
