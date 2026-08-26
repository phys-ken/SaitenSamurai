/**
 * info.js — ℹ ポップオーバー（補足説明）と「用語とファイルの地図」。
 *
 * 方針（ux-refresh-plan.md）:
 * - 画面に常時出すのは役割語だけ。実名・由来・消してよいか等はここに退避する
 * - ホバー/クリック/キーボードフォーカスのどれでも開ける（タッチ・a11y対応）
 * - 各解説の末尾から「用語とファイルの地図」へ飛べる
 */

// 解説文の台帳。data-info="キー" で参照する
const INFO_TEXTS = {
  'image-folder': {
    title: '答案画像フォルダ',
    body: 'スキャンした答案画像（jpg/png）が入ったフォルダを選びます。' +
      '処理結果はこの中に作られる「_saiten_grading_results」フォルダに入ります。' +
      '元の答案画像が書き換えられることはありません。',
  },
  'coord-file': {
    title: 'マーク位置の定義（座標ファイル）',
    body: 'マークシート上のどこにマーク欄があるかを定義した Excel です' +
      '（mark_areas.xlsx など）。Mark2 などのテンプレート作成ツールで作ります。' +
      '答案の様式を変えない限り、同じファイルを使い回せます。',
  },
  'answer-key': {
    title: '正答・配点（answer_key.xlsx）',
    body: '各問題の正答・配点・観点を入力した Excel です。読み取りを実行すると' +
      '雛形が結果フォルダに自動生成されるので、Excel で開いて入力してください。' +
      '選択すると内容を自動チェックし、問題があればここに表示します。',
  },
  'omr-result': {
    title: '読み取り結果',
    body: 'マークを読み取った結果の Excel（Mark2-Result-～.xlsx）です。' +
      '「答案を読み取る」を実行すると自動で作られ、ここに設定されます。' +
      '通常は手で選び直す必要はありません。過去の読み取り結果を使いたいときだけ' +
      '「ファイル選択」で指定します。',
  },
  'skip': {
    title: 'ID欄の列数（Skip）',
    body: 'マークシートの先頭にある、学年・クラス・出席番号などの' +
      '「採点しない列」の数です。Mark2 標準テンプレートでは 4 です。' +
      'ここが合っていないと問題番号がずれて採点されます。',
  },
  'omr-mode': {
    title: '読み取り方式',
    body: '「クラスタリング」は答案全体の濃淡から自動で判定する推奨方式です。' +
      '「しきい値」は従来式で、濃さ・面積の基準を手で調整できます。' +
      '判定に迷いが多い答案では、読み取り後に「マークを確認する」で直せます。',
  },
  'checker': {
    title: 'マークの確認',
    body: '読み取りに自信がないマーク（無記入・二重マークなど）を一覧で確認し、' +
      'その場で訂正できます。訂正は「訂正をxlsxに反映」を押すまで読み取り結果には' +
      '書き込まれません（訂正の下書きは自動保存されます）。',
  },
  'scoring': {
    title: '採点済み答案の生成',
    body: '読み取り結果と正答・配点から、○×・得点・合計点を描き込んだ答案画像を' +
      '「02_Graded_Detail」フォルダに出力します。描き込む内容は' +
      '「出力の設定」で変えられます。何度でもやり直せます。',
  },
  'summary': {
    title: '集計',
    body: '成績一覧 Excel、試験全体の統計、分析レポートを' +
      '「03_Final_Report」フォルダに出力します。こちらも何度でもやり直せます。',
  },
  'output-settings': {
    title: '出力の設定',
    body: '採点済み答案と集計に「何をどう描くか・載せるか」の設定です。' +
      'すべて既定のままで使えます。変えた設定は自動保存され、' +
      '次の生成・集計から反映されます。',
  },
  'total-position': {
    title: '合計点の位置',
    body: '答案画像のどこに合計点を書くかをドラッグで指定します。' +
      '未指定なら余白に自動配置されます。設定は' +
      'total_display_config.json に保存されます。',
  },
  'name-trim': {
    title: '集計シートの氏名画像',
    body: '答案の氏名欄を切り出して、成績一覧 Excel の各行に貼り付けます。' +
      '誰の答案か、Excel 上で照合しやすくなります。位置指定が必要です。',
  },
  'include-desc': {
    title: '記述を分析に含める',
    body: 'オフにすると、統計・分析レポートはマーク設問だけで計算されます。' +
      '記述の採点が終わっていない段階で集計したいときに使います。',
  },
  'session': {
    title: '中断と再開',
    body: '選んだファイルや設定は、操作のたびに結果フォルダ内の ' +
      'session_state.json へ自動保存されます。アプリを閉じても、' +
      '同じ答案画像フォルダを選び直せば「前回の続きから再開しますか？」と' +
      '確認されます。モード選択画面の「前回の採点を再開する」からも開けます。',
  },
  'desc-config': {
    title: '記述問題の設定',
    body: '答案上の記述欄をドラッグで囲んで、問題名・配点・観点を登録します。' +
      '設定は descriptive_config.json、採点結果は descriptive_scores.json に' +
      '自動保存されます。',
  },
};

let popover = null;
let currentKey = null;
let hideTimer = null;

function ensurePopover() {
  if (popover) return popover;
  popover = document.createElement('div');
  popover.className = 'info-popover';
  popover.setAttribute('role', 'tooltip');
  popover.hidden = true;
  popover.addEventListener('mouseenter', () => clearTimeout(hideTimer));
  popover.addEventListener('mouseleave', scheduleHide);
  document.body.appendChild(popover);
  return popover;
}

function showFor(btn) {
  const key = btn.dataset.info;
  const item = INFO_TEXTS[key];
  if (!item) return;
  clearTimeout(hideTimer);
  const pop = ensurePopover();
  currentKey = key;
  pop.replaceChildren();
  const h = document.createElement('h4');
  h.textContent = item.title;
  const p = document.createElement('p');
  p.textContent = item.body;
  const link = document.createElement('a');
  link.href = '#';
  link.textContent = '→ 用語とファイルの地図';
  link.addEventListener('click', (e) => {
    e.preventDefault();
    hideNow();
    document.getElementById('map-overlay').hidden = false;
  });
  pop.append(h, p, link);
  pop.hidden = false;
  const r = btn.getBoundingClientRect();
  const pw = Math.min(340, window.innerWidth - 24);
  pop.style.maxWidth = `${pw}px`;
  let left = r.left;
  if (left + pw > window.innerWidth - 12) left = window.innerWidth - pw - 12;
  pop.style.left = `${Math.max(12, left)}px`;
  const top = r.bottom + 6;
  pop.style.top = `${top}px`;
}

function scheduleHide() {
  clearTimeout(hideTimer);
  hideTimer = setTimeout(hideNow, 250);
}

function hideNow() {
  if (popover) popover.hidden = true;
  currentKey = null;
}

export function wireInfo() {
  document.querySelectorAll('.info-btn[data-info]').forEach((btn) => {
    btn.setAttribute('aria-label', '補足説明');
    btn.addEventListener('mouseenter', () => showFor(btn));
    btn.addEventListener('mouseleave', scheduleHide);
    btn.addEventListener('focus', () => showFor(btn));
    btn.addEventListener('blur', scheduleHide);
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      if (currentKey === btn.dataset.info && !popover.hidden) hideNow();
      else showFor(btn);
    });
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') hideNow();
  });
  // 地図ビューの開閉
  document.getElementById('btn-open-map')?.addEventListener('click', () => {
    document.getElementById('map-overlay').hidden = false;
  });
  document.getElementById('btn-map-close')?.addEventListener('click', () => {
    document.getElementById('map-overlay').hidden = true;
  });
  document.getElementById('map-overlay')?.addEventListener('click', (e) => {
    if (e.target.id === 'map-overlay') e.target.hidden = true;
  });
}
