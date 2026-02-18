#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
test_kmeans_omr.py — K-means OMR エンジン単体テスト (v4.5)

テスト対象:
  - extract_mark_features(): 特徴量抽出
  - recognize_marks_kmeans(): K-means クラスタリング認識
  - recognize_marks_kmeans() フォールバック動作
  - generate_kmeans_report(): HTML レポート生成
  - process_box_drawer() の omr_mode パラメータ後方互換
  - constants.py の K-means 定数
"""

import sys
from pathlib import Path

import cv2
import numpy as np
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "main_src"))

from constants import (
    APP_VERSION,
    OMR_MODE_THRESHOLD,
    OMR_MODE_KMEANS,
    KMEANS_N_CLUSTERS,
    KMEANS_MIN_SAMPLES,
    KMEANS_FEATURES,
)
from omr_engine import (
    extract_mark_features,
    recognize_marks,
    recognize_marks_kmeans,
    generate_kmeans_report,
)


# ============================================================
# ヘルパー
# ============================================================

def _make_gray_image(w=595, h=842):
    """白いグレースケール画像を作成する"""
    return np.full((h, w), 255, dtype=np.uint8)


def _make_coordinates(n_questions=10, n_choices=5, start_x=100, start_y=50,
                      cell_w=20, cell_h=12, gap_y=20):
    """テスト用座標リストを生成する"""
    coords = []
    y = start_y
    for q in range(1, n_questions + 1):
        for c in range(n_choices):
            coords.append({
                'question_no': q,
                'question': f'Q{q}',
                'choice': c,
                'raw_choice': c,
                'x': start_x + c * (cell_w + 2),
                'y': y,
                'width': cell_w,
                'height': cell_h,
            })
        y += cell_h + gap_y
    return coords


def _fill_mark(gray, coord, intensity=30):
    """座標の領域を暗い色で塗りつぶす（マーク済みに見せる）
    
    1ピクセルの白い枠を残すことで Otsu 閾値が正しく機能するようにする。
    """
    x, y, w, h = coord['x'], coord['y'], coord['width'], coord['height']
    # 1px 白枠を残して内側だけ塗りつぶし（Otsu が分散を検出できるように）
    x1, y1 = x + 1, y + 1
    x2, y2 = x + w - 1, y + h - 1
    if x2 > x1 and y2 > y1:
        gray[y1:y2, x1:x2] = intensity


# ============================================================
# テスト: 定数
# ============================================================

class TestConstants:
    def test_app_version(self):
        assert APP_VERSION == "4.5.0"

    def test_omr_modes(self):
        assert OMR_MODE_THRESHOLD == "threshold"
        assert OMR_MODE_KMEANS == "kmeans"

    def test_kmeans_params(self):
        assert KMEANS_N_CLUSTERS == 2
        assert KMEANS_MIN_SAMPLES == 50
        assert len(KMEANS_FEATURES) == 4
        assert 'filled_ratio' in KMEANS_FEATURES


# ============================================================
# テスト: 特徴量抽出
# ============================================================

class TestExtractMarkFeatures:
    def test_basic_extraction(self):
        """白画像 → filled_ratio ≈ 0, mean_inv_brightness ≈ 0"""
        gray = _make_gray_image()
        coords = _make_coordinates(n_questions=2, n_choices=3)
        features, meta = extract_mark_features(gray, coords)

        assert features.shape == (6, 4)  # 2問 × 3選択肢 = 6
        assert len(meta) == 6
        # 白画像なので filled_ratio は 0 に近い
        assert np.all(features[:, 0] < 0.1)

    def test_marked_region(self):
        """塗りつぶし領域 → filled_ratio > 0.5"""
        gray = _make_gray_image()
        coords = _make_coordinates(n_questions=1, n_choices=3)
        # 1番目の選択肢を塗りつぶし
        _fill_mark(gray, coords[0])
        features, meta = extract_mark_features(gray, coords)

        # 塗った領域は filled_ratio が高い
        assert features[0, 0] > 0.5
        # 塗っていない領域は低い
        assert features[1, 0] < 0.1
        assert features[2, 0] < 0.1

    def test_meta_has_question_info(self):
        """メタデータに question_no と choice が含まれる"""
        gray = _make_gray_image()
        coords = _make_coordinates(n_questions=3, n_choices=2)
        _, meta = extract_mark_features(gray, coords)

        assert all('question_no' in m for m in meta)
        assert all('choice' in m for m in meta)
        assert meta[0]['question_no'] == 1
        assert meta[0]['choice'] == 0

    def test_zero_area_region(self):
        """幅 or 高さ 0 の領域 → ゼロベクトル"""
        gray = _make_gray_image()
        coords = [{'question_no': 1, 'choice': 0, 'x': 10, 'y': 10, 'width': 0, 'height': 10}]
        features, _ = extract_mark_features(gray, coords)
        assert np.all(features[0] == 0.0)


# ============================================================
# テスト: K-means 認識
# ============================================================

class TestRecognizeMarksKmeans:
    def _make_test_image_with_marks(self, n_questions=20, n_choices=5, marked_choices=None):
        """テスト用のマーク済み画像と座標を生成する

        Args:
            marked_choices: {question_no: [choice_idx, ...]} 
                           Noneの場合はランダムにマーク
        """
        gray = _make_gray_image()
        coords = _make_coordinates(n_questions=n_questions, n_choices=n_choices)

        if marked_choices is None:
            rng = np.random.RandomState(42)
            marked_choices = {}
            for q in range(1, n_questions + 1):
                c = rng.randint(0, n_choices)
                marked_choices[q] = [c]

        for q_no, choices in marked_choices.items():
            for c in choices:
                target = next(
                    (coord for coord in coords
                     if coord['question_no'] == q_no and coord['choice'] == c),
                    None,
                )
                if target:
                    _fill_mark(gray, target, intensity=20)

        return gray, coords, marked_choices

    def test_basic_kmeans_recognition(self):
        """基本的なK-means認識が動作する"""
        gray, coords, expected = self._make_test_image_with_marks(n_questions=20)
        results, kmeans_info = recognize_marks_kmeans(gray, coords, min_samples=10)

        assert isinstance(results, dict)
        assert kmeans_info is not None
        assert 'features' in kmeans_info
        assert 'labels' in kmeans_info
        assert 'cutoff' in kmeans_info

    def test_kmeans_accuracy(self):
        """K-meansが正しいマークを検出できる"""
        marked_choices = {
            1: [0], 2: [1], 3: [2], 4: [3], 5: [4],
            6: [0], 7: [1], 8: [2], 9: [3], 10: [4],
            11: [0], 12: [1], 13: [2], 14: [3], 15: [4],
            16: [0], 17: [1], 18: [2], 19: [3], 20: [4],
        }
        gray, coords, _ = self._make_test_image_with_marks(
            n_questions=20, n_choices=5, marked_choices=marked_choices
        )
        results, info = recognize_marks_kmeans(gray, coords, min_samples=10)

        # 各問題で正しいマークが検出されている
        correct = 0
        for q_no, expected_choices in marked_choices.items():
            if q_no in results and set(results[q_no]) == set(expected_choices):
                correct += 1
        # 少なくとも80%以上正解
        assert correct >= 16, f"正解率が低すぎます: {correct}/20"

    def test_kmeans_info_structure(self):
        """kmeans_info にレポート生成用の情報が含まれる"""
        gray, coords, _ = self._make_test_image_with_marks(n_questions=20)
        _, info = recognize_marks_kmeans(gray, coords, min_samples=10)

        assert 'n_marked' in info
        assert 'n_empty' in info
        assert 'cluster_means' in info
        assert 'marked_cluster' in info
        assert info['n_marked'] + info['n_empty'] == len(coords)

    def test_fallback_to_threshold(self):
        """サンプル数 < min_samples → 閾値方式にフォールバック"""
        gray = _make_gray_image()
        coords = _make_coordinates(n_questions=2, n_choices=3)  # 6要素 < 50
        _fill_mark(gray, coords[0])

        results, kmeans_info = recognize_marks_kmeans(gray, coords, min_samples=50)

        # フォールバック時は kmeans_info = None
        assert kmeans_info is None
        # 結果は閾値方式で返される
        assert isinstance(results, dict)

    def test_bgr_image_handled(self):
        """BGR画像が入力された場合も正しく動作する"""
        gray = _make_gray_image()
        coords = _make_coordinates(n_questions=20, n_choices=5)
        _fill_mark(gray, coords[0])
        # BGRに変換
        bgr = cv2.cvtColor(gray, cv2.COLOR_GRAY2BGR)
        results, info = recognize_marks_kmeans(bgr, coords, min_samples=10)
        assert isinstance(results, dict)
        assert info is not None

    def test_double_mark_detection(self):
        """1設問に複数マーク → results にリストで格納"""
        marked_choices = {1: [0, 1]}  # Q1で2つマーク
        # 残りの問題も正常にマーク
        for q in range(2, 21):
            marked_choices[q] = [0]
        gray, coords, _ = self._make_test_image_with_marks(
            n_questions=20, n_choices=5, marked_choices=marked_choices,
        )
        results, _ = recognize_marks_kmeans(gray, coords, min_samples=10)
        # Q1で2つのマークが検出される
        if 1 in results:
            assert len(results[1]) >= 2, "ダブルマークが検出されていません"


# ============================================================
# テスト: HTML レポート生成
# ============================================================

class TestGenerateKmeansReport:
    def test_report_creation(self, tmp_path):
        """HTMLレポートファイルが正しく生成される"""
        # テストデータ作成
        features = np.random.rand(100, 4).astype(np.float64)
        labels = np.array([0] * 90 + [1] * 10)
        meta = [{'question_no': i // 5 + 1, 'choice': i % 5} for i in range(100)]

        infos = [{
            'filename': 'test_001.jpg',
            'info': {
                'features': features,
                'labels': labels,
                'meta': meta,
                'cluster_means': [0.1, 0.8],
                'marked_cluster': 1,
                'cutoff': 0.45,
                'n_marked': 10,
                'n_empty': 90,
                'scaler_mean': [0.5, 0.5, 0.5, 0.5],
                'scaler_scale': [0.1, 0.1, 0.1, 0.1],
            },
        }]

        output_path = tmp_path / "test_report.html"
        generate_kmeans_report(output_path, infos)

        assert output_path.exists()
        content = output_path.read_text(encoding='utf-8')
        assert 'K-means OMR Report' in content
        assert 'chart.js' in content.lower()
        assert '100' in content  # 総領域数

    def test_report_with_multiple_images(self, tmp_path):
        """複数画像のレポートが正しく集約される"""
        infos = []
        for i in range(3):
            features = np.random.rand(50, 4).astype(np.float64)
            labels = np.array([0] * 45 + [1] * 5)
            meta = [{'question_no': j // 5 + 1, 'choice': j % 5} for j in range(50)]
            infos.append({
                'filename': f'test_{i:03d}.jpg',
                'info': {
                    'features': features,
                    'labels': labels,
                    'meta': meta,
                    'cluster_means': [0.1, 0.8],
                    'marked_cluster': 1,
                    'cutoff': 0.45,
                    'n_marked': 5,
                    'n_empty': 45,
                    'scaler_mean': [0.5] * 4,
                    'scaler_scale': [0.1] * 4,
                },
            })

        output_path = tmp_path / "multi_report.html"
        generate_kmeans_report(output_path, infos)

        content = output_path.read_text(encoding='utf-8')
        assert '150' in content  # 50 × 3 = 150 total
        assert '処理画像数: 3' in content


# ============================================================
# テスト: 従来方式との互換性
# ============================================================

class TestBackwardCompatibility:
    def test_threshold_mode_unchanged(self):
        """閾値方式の recognize_marks() が変更なしで動作する"""
        gray = _make_gray_image()
        coords = _make_coordinates(n_questions=5, n_choices=5)
        _fill_mark(gray, coords[0])  # Q1-choice0
        _fill_mark(gray, coords[6])  # Q2-choice1

        results = recognize_marks(gray, coords, color_threshold=0.1, area_threshold=0.4)
        assert 1 in results
        assert 0 in results[1]

    def test_both_modes_produce_same_format(self):
        """K-meansと閾値方式の出力フォーマットが同じ"""
        gray, coords, _ = TestRecognizeMarksKmeans()._make_test_image_with_marks(n_questions=20)

        results_threshold = recognize_marks(gray, coords)
        results_kmeans, _ = recognize_marks_kmeans(gray, coords, min_samples=10)

        # 両方ともdict型で、値はリスト
        assert isinstance(results_threshold, dict)
        assert isinstance(results_kmeans, dict)
        for q_no, choices in results_threshold.items():
            assert isinstance(choices, list)
        for q_no, choices in results_kmeans.items():
            assert isinstance(choices, list)
