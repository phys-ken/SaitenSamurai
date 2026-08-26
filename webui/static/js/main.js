import { call } from './bridge.js';

async function init() {
  const status = document.getElementById('bridge-status');
  try {
    await call('ping');
    const info = await call('get_app_info');
    document.getElementById('version').textContent =
      `v${info.app_version} (webui ${info.webui_version})`;
    status.textContent = '接続済み';
    status.dataset.state = 'ok';
  } catch (e) {
    status.textContent = `接続エラー: ${e.message}`;
    status.dataset.state = 'error';
  }
}

init();
