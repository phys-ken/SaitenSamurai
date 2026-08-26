import { call } from './bridge.js';

// ---------------------------------------------------------------
// ログ
// ---------------------------------------------------------------
function log(message) {
  const el = document.getElementById('log');
  el.textContent += (el.textContent ? '\n' : '') + message;
  el.scrollTop = el.scrollHeight;
}

// ---------------------------------------------------------------
// state → 画面（UIは state の再描画だけを行う）
// ---------------------------------------------------------------
function setRow(rowId, summaryId, path, summaryHtmlBuilder) {
  const value = document.querySelector(`#${rowId} .ds-value`);
  value.textContent = path ?? '';
  value.classList.toggle('selected', Boolean(path));
  const summary = document.getElementById(summaryId);
  const built = path ? summaryHtmlBuilder() : null;
  if (built) {
    summary.textContent = built.text;
    summary.className = `ds-summary ${built.kind ?? ''}`;
    summary.hidden = false;
  } else {
    summary.hidden = true;
  }
}

export function render(state) {
  setRow('row-image-folder', 'summary-image-folder', state.image_folder,
    () => ({ text: `画像 ${state.image_count} 枚`, kind: '' }));

  setRow('row-coord-file', 'summary-coord-file', state.coord_file, () => {
    const s = state.coord_summary;
    if (!s) return null;
    if (s.warning) return { text: `⚠ ${s.warning}`, kind: 'warn' };
    return { text: `解答欄 ${s.answer_rows} 行 / 1行あたり ${s.marks_per_row} マーク`, kind: '' };
  });

  setRow('row-answer-key', 'summary-answer-key', state.answer_key, () => {
    const k = state.key_summary;
    if (!k) return null;
    if (!k.ok) return { text: `❌ ${k.errors.join('\n')}`, kind: 'err' };
    if (k.warnings.length) {
      return { text: `⚠ ${k.warnings.join('\n')}\n✓ ${k.stats_line}`, kind: 'warn' };
    }
    return { text: `✓ ${k.stats_line}`, kind: '' };
  });

  document.getElementById('skip-input').value = state.skip_questions;
}

// ---------------------------------------------------------------
// 操作
// ---------------------------------------------------------------
const ACTION_LOG_LABEL = {
  select_image_folder: '画像フォルダ',
  select_coord_file: '座標ファイル',
  select_answer_key: '正答データ',
};

async function runAction(name) {
  try {
    const res = await call(name);
    if (res.cancelled) return;
    render(res.state);
    log(`✓ ${ACTION_LOG_LABEL[name]}を選択しました`);
  } catch (e) {
    log(`❌ ${e.message}`);
    alert(e.message);
  }
}

function wireEvents() {
  document.querySelectorAll('[data-action]').forEach((btn) => {
    btn.addEventListener('click', () => runAction(btn.dataset.action));
  });
  document.getElementById('skip-input').addEventListener('change', async (ev) => {
    try {
      const res = await call('set_skip_questions', ev.target.value);
      render(res.state);
    } catch (e) {
      log(`❌ ${e.message}`);
      const res = await call('get_state');
      render(res.state);  // 不正値は元に戻す
    }
  });
}

// ---------------------------------------------------------------
// 起動
// ---------------------------------------------------------------
async function init() {
  const status = document.getElementById('bridge-status');
  try {
    await call('ping');
    const info = await call('get_app_info');
    document.getElementById('version').textContent =
      `v${info.app_version} (webui ${info.webui_version})`;
    status.textContent = '接続済み';
    status.dataset.state = 'ok';
    const res = await call('get_state');
    render(res.state);
    wireEvents();
  } catch (e) {
    status.textContent = `接続エラー: ${e.message}`;
    status.dataset.state = 'error';
  }
}

init();
