# -*- mode: python ; coding: utf-8 -*-
"""
採点侍 (SaitenSamurai) — PyInstaller spec ファイル

ビルド方法:
  build_exe.bat を実行するか、以下のコマンドを実行:
  pyinstaller saitensamurai.spec
"""

import os
import re
import sys

block_cipher = None

# main_src をインポートパスに追加
MAIN_SRC = os.path.join(os.path.dirname(os.path.abspath(SPEC)), 'main_src')

# saitensamurai.py からバージョンを読み取り、exe 名に付与
_version_file = os.path.join(MAIN_SRC, 'saitensamurai.py')
with open(_version_file, 'r', encoding='utf-8') as _f:
    _m = re.search(r'バージョン:\s*(\S+)', _f.read())
    _version = _m.group(1) if _m else 'unknown'
EXE_NAME = f'SaitenSamurai_v{_version}'

a = Analysis(
    ['main_src/saitensamurai.py'],
    pathex=[MAIN_SRC],
    binaries=[],
    datas=[
        ('resources/icon.ico', 'resources'),
        ('resources/samurai.png', 'resources'),
    ],
    hiddenimports=[
        # エントリポイント (遅延インポートで使われる)
        'saitensamurai',
        # main_src 内の全モジュール
        'constants',
        'omr_engine',
        'threshold_calibrator',
        'scoring_engine',
        'image_renderer',
        'summary_generator',
        'ctt_analyzer',
        'mark_checker',
        'gui_components',
        'main_gui',
        'descriptive_scorer',
        'descriptive_gui',
        'descriptive_renderer',
        'name_trimmer',
        'r_export',
        # オプショナル依存 (exe には含める)
        'fitz',
        'matplotlib',
        'matplotlib.backends.backend_agg',
        'matplotlib.backends.backend_tkagg',
        'reportlab',
        'reportlab.lib',
        'reportlab.platypus',
        'reportlab.graphics',
        # scikit-learn (KMeans, StandardScaler, PCA のみ使用)
        'sklearn',
        'sklearn.cluster',
        'sklearn.cluster._kmeans',
        'sklearn.preprocessing',
        'sklearn.preprocessing._data',
        'sklearn.decomposition',
        'sklearn.decomposition._pca',
        'sklearn.utils',
        'sklearn.utils._param_validation',
        'sklearn.metrics',
        'sklearn.metrics.pairwise',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 不要なモジュールを除外して軽量化
        # 注意: 標準ライブラリでも pandas/numpy/matplotlib が依存するものは
        #       除外しないこと (secrets, email, html, http 等)

        # --- 開発ツール ---
        'pytest',
        'IPython',
        'jupyter',
        'notebook',
        'sphinx',
        'docutils',
        'pip',
        'wheel',
        'test',
        # 'pydoc',  # sklearn が内部的に依存 (inspect→pydoc) — 除外不可
        'lib2to3',
        'tkinter.test',
        'idlelib',

        # --- matplotlib 不要バックエンド ---
        'matplotlib.backends.backend_gtk3',
        'matplotlib.backends.backend_gtk3agg',
        'matplotlib.backends.backend_gtk3cairo',
        'matplotlib.backends.backend_gtk4',
        'matplotlib.backends.backend_gtk4agg',
        'matplotlib.backends.backend_gtk4cairo',
        'matplotlib.backends.backend_qt',
        'matplotlib.backends.backend_qt5',
        'matplotlib.backends.backend_qt5agg',
        'matplotlib.backends.backend_qt5cairo',
        'matplotlib.backends.backend_qtagg',
        'matplotlib.backends.backend_qtcairo',
        'matplotlib.backends.backend_wx',
        'matplotlib.backends.backend_wxagg',
        'matplotlib.backends.backend_wxcairo',
        'matplotlib.backends.backend_webagg',
        'matplotlib.backends.backend_webagg_core',
        'matplotlib.backends.backend_nbagg',
        'matplotlib.backends.backend_macosx',
        'matplotlib.backends.backend_cairo',
        'matplotlib.backends.backend_pgf',
        'matplotlib.tests',

        # --- numpy/pandas テスト ---
        'numpy.tests',
        'pandas.tests',

        # --- Pillow 不要プラグイン ---
        'PIL.AvifImagePlugin',
        'PIL.WebPImagePlugin',

        # --- 不要な標準ライブラリ ---
        'sqlite3',
        'xmlrpc',
        'ftplib',
        'imaplib',
        'poplib',
        'smtplib',
        'nntplib',
        'socketserver',
        'pdb',
        'pickletools',

        # --- ネットワーク/タイムゾーン (オフラインアプリでは不要) ---
        'certifi',
        'urllib3',
        'requests',

        # --- GUI ツールキット (tkinter のみ使用) ---
        'PyQt5',
        'PyQt6',
        'PySide2',
        'PySide6',
        'wx',
        'gi',

        # --- sklearn 不要サブモジュール (KMeans/StandardScaler/PCA 以外) ---
        'sklearn.linear_model',
        'sklearn.ensemble',
        'sklearn.tree',
        'sklearn.svm',
        'sklearn.neighbors',
        'sklearn.model_selection',
        'sklearn.datasets',
        'sklearn.feature_selection',
        'sklearn.feature_extraction',
        'sklearn.manifold',
        'sklearn.inspection',
        'sklearn.impute',
        'sklearn.gaussian_process',
        'sklearn.compose',
        'sklearn.covariance',
        'sklearn.neural_network',
        'sklearn.mixture',
        'sklearn.cross_decomposition',
        'sklearn.semi_supervised',
        'sklearn.frozen',
        'sklearn.tests',
        'sklearn.experimental',
        'sklearn._loss',
        # sklearn 不要内部モジュール (KMeans Lloyd/Elkan 以外)
        'sklearn.cluster._hdbscan',
        'sklearn.cluster._dbscan_inner',
        'sklearn.cluster._hierarchical_fast',
        'sklearn.cluster._spectral',
        'sklearn.cluster._agglomerative',
        'sklearn.cluster._bisect_k_means',
        'sklearn.cluster._optics',
        'sklearn.cluster._meanshift',
        'sklearn.cluster._affinity_propagation',
        'sklearn.cluster._birch',
        'sklearn.cluster._feature_agglomeration',
        'sklearn.cluster._k_means_minibatch',
        'sklearn.decomposition._cdnmf_fast',
        'sklearn.decomposition._online_lda_fast',
        'sklearn.preprocessing._target_encoder_fast',
        'sklearn.preprocessing._csr_polynomial_expansion',
        'sklearn.preprocessing._polynomial',

    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# --- 不要データファイルの除外 ---
# haarcascade XML（顔認識等、本アプリ不使用）: ~7MB 非圧縮
# matplotlib サンプルデータ: ~0.5MB
# matplotlib 不要フォント: ~6MB（DejaVuSans のみ残す）
# matplotlib images/plot_directive: ~0.15MB
# sklearn/scipy テストデータ・不要データ
a.datas = [
    d for d in a.datas
    if not d[0].startswith('cv2/data/haarcascade')
    and not d[0].startswith('matplotlib/mpl-data/sample_data')
    and not d[0].startswith('matplotlib/mpl-data/images')
    and not d[0].startswith('matplotlib/mpl-data/plot_directive')
    and not d[0].startswith('matplotlib/mpl-data/fonts/afm')
    and not d[0].startswith('matplotlib/mpl-data/fonts/pdfcorefonts')
    and not (
        d[0].startswith('matplotlib/mpl-data/fonts/ttf')
        and not d[0].split('/')[-1].startswith('DejaVuSans')
    )
    and not d[0].startswith('sklearn/datasets/')
    and 'tests' not in d[0].split('/')
]

# --- 不要バイナリの除外 ---
# Pillow AVIF/WebP プラグイン DLL（マークシート処理で不使用）
# opencv ffmpeg DLL（headless版では不要）
_excluded_binary_prefixes = (
    '_avif', '_webp', 'opencv_videoio_ffmpeg',
)
a.binaries = [
    b for b in a.binaries
    if not any(b[0].startswith(p) for p in _excluded_binary_prefixes)
]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name=EXE_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI アプリなのでコンソール非表示
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='resources/icon.ico',
)
