# 横断精読による改善バックログ（2026-08-26）

## 実行計画（grill-me 確定）

運用確認の結果:
- **フォルダ切替は完全クリア**（切替＝新しい試験。設定含め初期化し、復元提案だけ残す。
  正答の自動検出と Mark2-Result の自動拾い直しは新フォルダに対して働く）
- **打鍵ごと保存は維持**（成績最優先）。性能は保存以外（IPC削減・走査キャッシュ）で稼ぐ
- **S+A 修正完了で即 v5.0.0-beta.2**（実機確認は beta.2 で行う）
- **docs は新UI前提に書き換え、従来UIは差分タブ**（pymdownx.tabbed）

フェーズ:
1. **S群**（下表12件）— ✅ 完了（2026-08-26）
2. **A群**（A1〜A10）— ✅ 完了
3. **v5.0.0-beta.2 タグ** — ✅ 公開済み（S/A/B 修正を同梱）
4. **B群** — ✅ 完了（vlist共通化・aria系の一部は D 群へ残置）
5. **docs 全面改稿** — ✅ 完了（新UI前提＋従来UIタブ、strict ビルド警告ゼロ）

---

3本の監査（フロントエンド／ブリッジ・main_src差分／ドキュメント整合）＋手動検証の統合。
機能追加は含まない。**S はベータに載っている実バグで、修正推奨順に並べてある。**

## S. 重大バグ（誤採点・データ喪失に直結）

| # | 内容 | 場所 | 概要と直し方 |
|---|---|---|---|
| S1 | 復元候補のセッションを上書き | bridge.py select_image_folder | `session_found` を返す直前に自動保存が走り、復元対象ファイルをほぼ空の state で上書き（実データは .bak に一世代のみ）。「復元しますか？」に はい と答えても前回の内容は戻らない。→ session_found 検出時は自動保存をスキップ |
| S2 | フォルダ切替で記述の取り違え | descriptive.py `_desc_crops` | `_load_descriptive_state` がクロップキャッシュを捨てないため、フォルダAの切り出し・ファイル名がフォルダBでも配られ、**Aのファイル名でBのスコアに保存**される。→ 再読込時に `_desc_crops = None` |
| S3 | フォルダ切替でチェッカーが旧フォルダへ書く | checker.py | `_checker` が残留し、新フォルダの画像×旧エントリ、訂正CSVは旧フォルダへ。→ フォルダ変更時に `close_mark_checker()` |
| S4 | フォルダ切替で読み取り結果が残留 | bridge.py | 旧フォルダの Mark2-Result のまま採点・集計の検証を通過し、**別試験の読み取り結果で採点**できてしまう。→ フォルダ変更時に `omr_result=None` |
| S5 | 手書き: 答案送り連打の競合 | descriptive.js gotoSheet | await 順序の競合で古い画像＋古い筆跡が新しい見出しの下に載り、**筆跡が別答案に保存**される。→ 世代トークン。あわせて `img.onerror` 未設定（失敗時にビューが固まる） |
| S6 | 手書き: 消しゴムドラッグで例外→undo喪失 | handwriting.js redraw | 消しゴム中の `drawing` に points が無く TypeError。空振り判定が誤って `undoStack.pop()` し、**ドラッグ中に消した筆跡は Ctrl+Z で戻らない**。→ `if (drawing && !drawing.eraser)` |
| S7 | 2桁バッファが答案・問題を跨いで確定 | descriptive.js | 「1」入力→500ms 以内に答案送り/タブ切替すると**次の対象に1点が入る**。→ 遷移時に `clearTimeout`＋バッファ破棄 |
| S8 | 領域再指定後も旧クロップを表示 | descriptive.js cropLoader | JS 側 LRU に `clear()` 呼び出しが無く、**旧領域の切り出しを見て採点**しうる。→ 領域変更・削除・再切り出し時に clear（vlist の clear は pending/queue も破棄するよう補強） |
| S9 | 焼き込み中の筆跡保存でジョブ失敗 | handwriting.py `_bake` | 反復中の dict をJSスレッドが変更→ RuntimeError で「処理に失敗しました」（生成自体は完了済み）。→ スナップショットを取って反復 |
| S10 | region 検証前に config を破壊 | descriptive.py update/add | 不正 region で assert 前に `q["region"]` が壊れ、次の保存でディスクへ永続化。幅・高さ正値検証も欠落。→ 検証後に代入 |
| S11 | チェッカー既定ソートが表示と不一致 | checker.js openChecker | `view` を丸ごと差し替えて `sort` が消え、**実際は画像名順なのにセレクトは白さ順**。→ 代入をプロパティ単位に |
| S12 | ウィンドウリサイズで枠・筆跡がずれる | descriptive.js/region-picker.js | 一枚採点の領域枠・手書きキャンバス・ピッカー矩形が旧寸法のまま。→ ResizeObserver で layoutSheet 再実行 |

## A. 高優先（品質・性能・導線）

| # | 内容 | 概要 |
|---|---|---|
| A1 | 手書きの高DPI・タッチ対応 | canvas が devicePixelRatio 非対応で **Windows 125/150% 表示で筆跡がぼける**。`.sheet-stage/.hw-canvas` に touch-action が無くペン/タッチでストロークが途切れる |
| A2 | 採点キー1打の裏で IPC 3往復＋フォルダ走査 | set_score→get_state→render→get_progress(出力フォルダ iterdir)。400人連続採点で効く。応答 state の再利用・ステッパー更新の契機限定・打鍵ごとの fsync/全走査のデバウンス |
| A3 | チェッカー訂正1件ごとに可視カード全再生成 | 入力中の input からフォーカスが外れる。`grid.refreshItem(i)` を追加 |
| A4 | ジョブ排他の穴 | 実行中でもデータソース変更・PDF・自動調整・xlsx反映が通る（ワーカーは実行時の state を読む）。job 開始時に引数をスナップショット＋実行中ガード＋Lock |
| A5 | 訂正CSVがアトミック保存でない | 1打鍵ごとに全件直書き。異常終了で目視訂正が全ロスト。atomic_json_save と同じ手順を CSV にも |
| A6 | モード選択画面にステッパーが出る | updateStepper が無条件 `hidden=false` |
| A7 | ダウンロード導線: β が latest に出ない | docs の「最新版」ボタンからは v5.0.0-beta.1 に永久に辿り着けない。βへの直リンク節と「データは両UIで完全共通」を download/faq に明記。リリース本文の手順が従来UIを指す点も |
| A8 | アプリ内の旧用語残存 | モードカード説明「座標ファイル/OMR認識」（最初に見る画面）、ログ台帳・bridge エラー文の「座標ファイル/OMR結果/Skip/認識」約20箇所 |
| A9 | 「答案に書き込む」がマークのみ×既存記述設定で開けない | 過去に記述採点したフォルダだと内部APIエラー文の alert。コメント専用経路では questions を空に |
| A10 | 一枚採点を先に開くと開発者向けエラー | 「先に start_descriptive_scoring を実行してください」がそのまま alert に。list_descriptive_targets をクロップ非依存に |

## B. 中優先

- モーダル作法の統一（rs/map/picker/lightbox で Esc・背景クリック・フォーカスがバラバラ。ライトボックス背後にショートカットが貫通し**裏の答案に得点が入る**）
- エラー処理の4流派（alert+log／logのみ／握り潰し／未処理拒否）を reportError 1本に統一。未 catch の async ハンドラ約10箇所
- チェッカーを閉じてもメイン state が同期されない（ステッパー・ボタン活性が古いまま）
- 復元失敗時にモードだけ書き換わる（session.py の検証順）／get_checker_entries の引数検証／open_original_image のパストラバーサル／set_handwriting・set_descriptive_score のファイル名実在チェックとサイズ上限
- `_temp/` の切り出し一時フォルダが開くたび無限に増える（数百MB級）＋記述ビューを開くたびの全再切り出し
- innerHTML へのファイル名/問題名流し込み（表示崩れ）→ textContent 化
- find_japanese_font が DejaVu を返すようになり、CTT グラフ/PDF の「日本語フォント無し」警告経路が死んだ（分離した API に）
- 平均表示が前問のまま／削除済み問題の currentQid 残留／得点順ソート中の自動送り先／コメント専用時の「全問採点済み N枚」誤表示
- 仮想グリッド系の checker/descriptive 重複（片方だけ直る構造）を vlist に共通化
- アプリ内文言ゆれ（一枚ずつ採点する/一枚採点、記述問題の設定/記述問題設定、補正画像の説明3通り 等）とヒント行のキー説明抜け（m/b/Ctrl+Z）
- credits/THIRDPARTYLICENSES に pywebview 未掲載

## C. ドキュメント全面改稿（別フェーズ推奨・方針決定が先）

docs は files-map/changelog 以外、**実質すべて旧 tk UI 前提**（スクショ36枚も全部 tk）。
- 方針決定: 各ページを「新UI/従来UI」タブ化（pymdownx.tabbed 導入済み）か、新UI ページ新設か
- quickstart/usage 3本/features/faq/index/README の用語・手順・ボタン名の全面追随（監査で約80箇所を特定済み — 詳細はセッション記録）
- 新UI 限定機能（手書き・進行バー・キーボード採点・出力の設定・地図）の記載追加、tk 限定機能（キャリブレーター画面・パス修復ダイアログ・薄い/濃いタブ 等）の「従来UIのみ」明示
- 新UI スクショ 7枚以上の撮り直し、FAQ に WebView2・新旧の選び方・手書きの安全性を追加
- files-map に handwriting.json / 00_Processing_Clean を追記、00_Processing の説明を実体（枠描画済み）に統一

## D. 低優先（整理・磨き）

- CSS: 死に規則（.checker-grid/.checker-pager/.placeholder）、`.crop-img { flex:1 }` が JS の height 指定を殺している、白色トークンの混在、`shu-stamp` アニメが**未配線**（署名要素の意図が実装されていない — 付けるか消すか）
- vlist 画像キューのキャンセル不能／一枚採点の得点表の直列取得／get_sheet_image の毎回 base64
- stepper の li をキーボード操作可能に（button化＋aria-current）、ログの aria-live 全文再読み上げ、progress のアクセシブル名
- `_push` の U+2028 エスケープ、単一PDFの中断不可、smoke 失敗ハンドラ、xdg-open のゾンビ
- render-settings の数値が文字列で保存される、`isOpen()` 未使用 export ほか
