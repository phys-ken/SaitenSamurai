"""
threshold_calibrator.py - 閾値キャリブレーションモジュール

OMRマーク認識の閾値(color_threshold, area_threshold)を自動推定する。
全画像から二値化パラメータを分析し、最適な閾値を提案する。
"""

from pathlib import Path
import logging
import cv2
import numpy as np

logger = logging.getLogger(__name__)

from constants import RESULTS_FOLDER
from omr_engine import (
    parse_excel_coordinates,
    detect_corner_markers,
    apply_perspective_transform,
)


def collect_mark_fill_ratios(image, coordinates, color_threshold):
    """
    1枚の補正済み画像の全マーク領域について、黒画素率(fill_ratio)を計測する。
    recognize_marks() と同じ二値化ロジックだが、判定(0/1)ではなく生の比率を返す。
    中間ファイルは生成しない。

    Args:
        image: 補正済み画像 (Gray or BGR, 595x842)
        coordinates: マーク領域のリスト
        color_threshold: 画素値の閾値 (0.0-1.0)

    Returns:
        list[dict]: 各要素は {
            'question_no': int,
            'choice': int,
            'fill_ratio': float (0.0-1.0),
            'x': int, 'y': int, 'w': int, 'h': int
        }
    """
    if len(image.shape) == 3:
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    else:
        gray = image

    pixel_threshold = int((1.0 - color_threshold) * 255)
    _, binary = cv2.threshold(gray, pixel_threshold, 255, cv2.THRESH_BINARY_INV)

    results = []
    for coord in coordinates:
        x, y, w, h = coord['x'], coord['y'], coord['width'], coord['height']
        if w <= 0 or h <= 0:
            continue
        # 画像境界チェック
        img_h, img_w = binary.shape[:2]
        if y < 0 or x < 0 or y + h > img_h or x + w > img_w:
            continue

        roi = binary[y:y+h, x:x+w]
        total_pixels = w * h
        if total_pixels == 0:
            continue
        marked_pixels = cv2.countNonZero(roi)
        ratio = marked_pixels / total_pixels

        results.append({
            'question_no': coord['question_no'],
            'choice': coord['choice'],
            'fill_ratio': ratio,
            'x': x, 'y': y, 'w': w, 'h': h
        })
    return results


def estimate_color_threshold_from_pixels(images_gray_list, coordinates):
    """
    複数画像のマーク領域内の全画素値からOtsu法で最適なcolor_thresholdを推定する。
    cv2のOtsuは画像単位だが、ここでは全マーク領域を統合した1つのヒストグラムに対して
    クラス間分散最大化を適用する。

    Args:
        images_gray_list: グレースケール画像のリスト (各595x842)
        coordinates: マーク領域のリスト

    Returns:
        dict: {
            'recommended_color_threshold': float (0.0-1.0),
            'otsu_pixel_value': int (0-255),
            'histogram': np.array (256,)
        }
    """
    # 全マーク領域の画素値を統合ヒストグラムとして収集
    histogram = np.zeros(256, dtype=np.int64)
    for gray in images_gray_list:
        img_h, img_w = gray.shape[:2]
        for coord in coordinates:
            x, y, w, h = coord['x'], coord['y'], coord['width'], coord['height']
            if w <= 0 or h <= 0 or y < 0 or x < 0 or y + h > img_h or x + w > img_w:
                continue
            roi = gray[y:y+h, x:x+w]
            hist = cv2.calcHist([roi], [0], None, [256], [0, 256])
            histogram += hist.flatten().astype(np.int64)

    total = histogram.sum()
    if total == 0:
        return {'recommended_color_threshold': 0.1, 'otsu_pixel_value': 229, 'histogram': histogram}

    # Otsu法: クラス間分散最大化
    best_threshold = 0
    best_variance = 0.0
    sum_total = np.dot(np.arange(256), histogram)
    sum_bg = 0.0
    weight_bg = 0

    for t in range(256):
        weight_bg += histogram[t]
        if weight_bg == 0:
            continue
        weight_fg = total - weight_bg
        if weight_fg == 0:
            break

        sum_bg += t * histogram[t]
        mean_bg = sum_bg / weight_bg
        mean_fg = (sum_total - sum_bg) / weight_fg

        variance = weight_bg * weight_fg * (mean_bg - mean_fg) ** 2
        if variance > best_variance:
            best_variance = variance
            best_threshold = t

    # pixel_threshold → color_threshold に変換
    # color_threshold = 1.0 - (pixel_threshold / 255)
    recommended = 1.0 - (best_threshold / 255.0)
    # 妥当な範囲にクランプ (0.03-0.35)
    recommended = max(0.03, min(0.35, recommended))

    return {
        'recommended_color_threshold': round(recommended, 3),
        'otsu_pixel_value': best_threshold,
        'histogram': histogram
    }


def kmeans_2class(data, max_iter=50):
    """
    2クラスのK-Means (numpy のみで実装)。scikit-learn不要。

    Args:
        data: 1次元のnp.array (全マークのfill_ratio)
        max_iter: 最大反復回数

    Returns:
        dict: {
            'center_low': float,   # 未マーク群の中心
            'center_high': float,  # マーク群の中心
            'boundary': float,     # 2中心の中間値
            'labels': np.array,    # 各データの所属ラベル (0=low, 1=high)
        }
    """
    if len(data) == 0:
        return {'center_low': 0.0, 'center_high': 1.0, 'boundary': 0.5, 'labels': np.array([])}

    data = np.asarray(data, dtype=np.float64)

    # 初期中心: 最小値と最大値
    c_low = data.min()
    c_high = data.max()
    if c_low == c_high:
        return {'center_low': c_low, 'center_high': c_high, 'boundary': c_low, 'labels': np.zeros(len(data), dtype=int)}

    labels = np.zeros(len(data), dtype=int)
    for _ in range(max_iter):
        # 割り当て
        dist_low = np.abs(data - c_low)
        dist_high = np.abs(data - c_high)
        new_labels = (dist_high < dist_low).astype(int)

        # 収束判定
        if np.array_equal(labels, new_labels):
            break
        labels = new_labels

        # 中心更新
        mask_low = labels == 0
        mask_high = labels == 1
        if mask_low.any():
            c_low = data[mask_low].mean()
        if mask_high.any():
            c_high = data[mask_high].mean()

    # c_low < c_high を保証
    if c_low > c_high:
        c_low, c_high = c_high, c_low
        labels = 1 - labels

    boundary = (c_low + c_high) / 2.0

    return {
        'center_low': float(c_low),
        'center_high': float(c_high),
        'boundary': float(boundary),
        'labels': labels
    }


def analyze_fill_ratio_distribution(all_ratios):
    """
    全マーク箇所のfill_ratioリストを分析し、推奨area_thresholdと分類結果を返す。

    Args:
        all_ratios: list[dict] — collect_mark_fill_ratios() の戻り値を全画像分結合したもの

    Returns:
        dict: {
            'recommended_area_threshold': float,
            'cluster_unmarked_mean': float,
            'cluster_marked_mean': float,
            'total_count': int,
            'marked_count': int,
            'unmarked_count': int,
            'classified': list[dict] — 各要素に 'is_marked' (bool) キーを追加
            'borderline_marked': list[dict] — マーク有と判定された中で fill_ratio が低い順 (境界付近)
            'borderline_unmarked': list[dict] — マーク無と判定された中で fill_ratio が高い順 (境界付近)
            'stable_marked': list[dict] — 確実にマーク有 (fill_ratio が高い)
            'stable_unmarked': list[dict] — 確実にマーク無 (fill_ratio が低い)
        }
    """
    if not all_ratios:
        return {
            'recommended_area_threshold': 0.4,
            'cluster_unmarked_mean': 0.0,
            'cluster_marked_mean': 1.0,
            'total_count': 0, 'marked_count': 0, 'unmarked_count': 0,
            'classified': [], 'borderline_marked': [], 'borderline_unmarked': [],
            'stable_marked': [], 'stable_unmarked': []
        }

    ratios_array = np.array([r['fill_ratio'] for r in all_ratios])
    km = kmeans_2class(ratios_array)

    recommended = km['boundary']
    # 妥当な範囲にクランプ (0.05-0.80)
    recommended = max(0.05, min(0.80, recommended))

    # 各要素にクラスタリング結果を付与
    classified = []
    for i, entry in enumerate(all_ratios):
        item = dict(entry)
        item['is_marked'] = bool(km['labels'][i] == 1) if len(km['labels']) > i else False
        classified.append(item)

    # マーク有/無に分離
    marked = [c for c in classified if c['is_marked']]
    unmarked = [c for c in classified if not c['is_marked']]

    # 境界付近: マーク有の中で fill_ratio が低い順 (最も薄いマーク)
    borderline_marked = sorted(marked, key=lambda x: x['fill_ratio'])[:10]
    # 境界付近: マーク無の中で fill_ratio が高い順 (最も濃いノーマーク)
    borderline_unmarked = sorted(unmarked, key=lambda x: x['fill_ratio'], reverse=True)[:10]
    # 安定: マーク有の中で fill_ratio が高い順 (確実な黒塗り)
    stable_marked = sorted(marked, key=lambda x: x['fill_ratio'], reverse=True)[:10]
    # 安定: マーク無の中で fill_ratio が低い順 (確実な白紙)
    stable_unmarked = sorted(unmarked, key=lambda x: x['fill_ratio'])[:10]

    return {
        'recommended_area_threshold': round(recommended, 3),
        'cluster_unmarked_mean': km['center_low'],
        'cluster_marked_mean': km['center_high'],
        'total_count': len(classified),
        'marked_count': len(marked),
        'unmarked_count': len(unmarked),
        'classified': classified,
        'borderline_marked': borderline_marked,
        'borderline_unmarked': borderline_unmarked,
        'stable_marked': stable_marked,
        'stable_unmarked': stable_unmarked
    }


def run_threshold_calibration(image_folder, coord_excel_path, skip_questions=0):
    """
    閾値キャリブレーションのオーケストレーション関数。
    画像フォルダ内のサンプル画像を使って最適な color_threshold と area_threshold を推定する。
    ファイル出力は一切行わない（全てメモリ上で完結）。

    Args:
        image_folder: 画像フォルダのパス
        coord_excel_path: 座標定義Excelファイルのパス
        skip_questions: スキップする問題数

    Returns:
        dict: {
            'recommended_color_threshold': float,
            'recommended_area_threshold': float,
            'color_analysis': dict (estimate_color_threshold_from_pixels の戻り値),
            'area_analysis': dict (analyze_fill_ratio_distribution の戻り値),
            'corrected_images': list[(image_name, gray_image)],
            'coordinates': list[dict],
            'image_count': int,
            'error_images': list[str]
        }
    """
    image_folder = Path(image_folder)
    coord_excel_path = Path(coord_excel_path)

    # 1. 座標ファイルをパース (テンプレートCSVの保存は parse_excel_coordinates 内で行われるが
    #    gitignoreされているため問題なし)
    coordinates, question_groups = parse_excel_coordinates(coord_excel_path, skip_questions)
    if not coordinates:
        raise ValueError("座標データが空です。座標ファイルを確認してください。")

    # 2. 画像読み込み → コーナー検出 → 射影変換
    image_files = sorted(image_folder.glob('*.jpg')) + sorted(image_folder.glob('*.png'))
    image_files = [f for f in image_files if RESULTS_FOLDER not in str(f)]

    if not image_files:
        raise ValueError(f"{image_folder} に画像ファイルが見つかりません")

    corrected_images = []  # [(filename, gray_image), ...]
    error_images = []

    for img_path in image_files:
        try:
            with open(str(img_path), 'rb') as f:
                img_bytes = f.read()
            image = cv2.imdecode(np.frombuffer(img_bytes, np.uint8), cv2.IMREAD_COLOR)
            if image is None:
                raise ValueError("画像を読み込めません")
            markers = detect_corner_markers(image, debug=False)
            corrected, _ = apply_perspective_transform(image, markers)
            gray = cv2.cvtColor(corrected, cv2.COLOR_BGR2GRAY)
            corrected_images.append((img_path.name, gray))
        except Exception as e:
            error_images.append(f"{img_path.name}: {e}")

    if not corrected_images:
        raise ValueError("処理可能な画像がありません")

    # 3. color_threshold を推定
    gray_list = [g for _, g in corrected_images]
    color_analysis = estimate_color_threshold_from_pixels(gray_list, coordinates)
    recommended_color = color_analysis['recommended_color_threshold']

    # 4. 推定された color_threshold で全画像の fill_ratio を収集
    all_ratios = []
    for img_name, gray in corrected_images:
        ratios = collect_mark_fill_ratios(gray, coordinates, recommended_color)
        for r in ratios:
            r['image_name'] = img_name
        all_ratios.extend(ratios)

    # 5. area_threshold を推定
    area_analysis = analyze_fill_ratio_distribution(all_ratios)

    return {
        'recommended_color_threshold': recommended_color,
        'recommended_area_threshold': area_analysis['recommended_area_threshold'],
        'color_analysis': color_analysis,
        'area_analysis': area_analysis,
        'corrected_images': corrected_images,
        'coordinates': coordinates,
        'image_count': len(corrected_images),
        'error_images': error_images
    }


def reclassify_with_threshold(all_ratios, area_threshold):
    """
    既に収集済みの fill_ratio リストに対して、新しい area_threshold で再分類する。
    スライダー操作時のリアルタイム更新に使用。再計算コスト最小。

    Args:
        all_ratios: list[dict] — fill_ratio を含む全マーク情報
        area_threshold: float — 新しい面積閾値

    Returns:
        analyze_fill_ratio_distribution() と同じ構造の dict
        (ただし K-Means ではなく、単純に area_threshold で切り分ける)
    """
    if not all_ratios:
        return analyze_fill_ratio_distribution([])

    classified = []
    for entry in all_ratios:
        item = dict(entry)
        item['is_marked'] = entry['fill_ratio'] > area_threshold
        classified.append(item)

    marked = [c for c in classified if c['is_marked']]
    unmarked = [c for c in classified if not c['is_marked']]

    borderline_marked = sorted(marked, key=lambda x: x['fill_ratio'])[:10]
    borderline_unmarked = sorted(unmarked, key=lambda x: x['fill_ratio'], reverse=True)[:10]
    stable_marked = sorted(marked, key=lambda x: x['fill_ratio'], reverse=True)[:10]
    stable_unmarked = sorted(unmarked, key=lambda x: x['fill_ratio'])[:10]

    return {
        'recommended_area_threshold': area_threshold,
        'cluster_unmarked_mean': np.mean([c['fill_ratio'] for c in unmarked]) if unmarked else 0.0,
        'cluster_marked_mean': np.mean([c['fill_ratio'] for c in marked]) if marked else 0.0,
        'total_count': len(classified),
        'marked_count': len(marked),
        'unmarked_count': len(unmarked),
        'classified': classified,
        'borderline_marked': borderline_marked,
        'borderline_unmarked': borderline_unmarked,
        'stable_marked': stable_marked,
        'stable_unmarked': stable_unmarked
    }


def recollect_and_reclassify(corrected_images, coordinates, color_threshold, area_threshold):
    """
    color_threshold が変更されたときに、全画像のfill_ratioを再収集し、
    area_threshold で再分類する。color_thresholdが変わると二値化が変わるため
    fill_ratio 自体の再計算が必要。

    Args:
        corrected_images: list[(filename, gray_image)]
        coordinates: マーク座標リスト
        color_threshold: 新しい色閾値
        area_threshold: 新しい面積閾値

    Returns:
        tuple: (all_ratios, analysis)
            all_ratios: list[dict] — 新しいfill_ratio リスト
            analysis: reclassify_with_threshold() の戻り値
    """
    all_ratios = []
    for img_name, gray in corrected_images:
        ratios = collect_mark_fill_ratios(gray, coordinates, color_threshold)
        for r in ratios:
            r['image_name'] = img_name
        all_ratios.extend(ratios)

    analysis = reclassify_with_threshold(all_ratios, area_threshold)
    return all_ratios, analysis
