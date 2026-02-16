#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
StudentAnswerSheetViewer のユニットテスト

テスト対象:
1. クラスのインスタンス化
2. fill_ratio のキャッシュ動作
3. 描画コードがエラーなく実行されること
4. 閾値同期メソッド
5. ThresholdCalibratorGUI のクリックハンドラ

GUIテストのため、tkinter のヘッドレステスト手法を使用。
"""

import sys
import os
import unittest
import numpy as np

import tkinter as tk


class TestStudentAnswerSheetViewerUnit(unittest.TestCase):
    """StudentAnswerSheetViewer のユニットテスト"""

    @classmethod
    def setUpClass(cls):
        """tkinter のルートウィンドウを取得（conftest の共有ルートを使用）"""
        from conftest import get_shared_tk_root
        cls.root = get_shared_tk_root()

    @classmethod
    def tearDownClass(cls):
        """共有ルートは conftest が管理するため何もしない"""
        pass

    def _create_dummy_data(self):
        """テスト用のダミーデータを生成"""
        # 595×842 のグレースケール画像（白黒のパターン）
        gray = np.ones((842, 595), dtype=np.uint8) * 230  # 白に近い背景

        # いくつかのマーク領域に黒い塗りつぶしを描画（マーク有を模擬）
        # Q1-C0: 確実にマーク有
        gray[100:115, 100:115] = 30  # 濃い黒

        # Q1-C1: マーク無
        # (白のまま)

        # Q2-C0: 薄いマーク
        gray[150:165, 100:115] = 150

        coordinates = [
            {'question_no': 1, 'question': 'Q1', 'choice': 0, 'x': 100, 'y': 100, 'width': 15, 'height': 15},
            {'question_no': 1, 'question': 'Q1', 'choice': 1, 'x': 130, 'y': 100, 'width': 15, 'height': 15},
            {'question_no': 2, 'question': 'Q2', 'choice': 0, 'x': 100, 'y': 150, 'width': 15, 'height': 15},
            {'question_no': 2, 'question': 'Q2', 'choice': 1, 'x': 130, 'y': 150, 'width': 15, 'height': 15},
        ]

        return gray, coordinates

    def test_01_import(self):
        """saitensamurai モジュールがインポートできること"""
        import saitensamurai
        self.assertTrue(hasattr(saitensamurai, 'StudentAnswerSheetViewer'))
        self.assertTrue(hasattr(saitensamurai, 'ThresholdCalibratorGUI'))
        self.assertTrue(hasattr(saitensamurai, 'collect_mark_fill_ratios'))

    def test_02_instantiation(self):
        """StudentAnswerSheetViewer がインスタンス化できること"""
        from saitensamurai import StudentAnswerSheetViewer
        gray, coordinates = self._create_dummy_data()

        viewer = StudentAnswerSheetViewer(
            parent_window=self.root,
            image_name="test_image.jpg",
            gray_image=gray,
            coordinates=coordinates,
            color_threshold=0.1,
            area_threshold=0.4,
            calibrator_gui=None
        )

        # ウィンドウが作成されていること
        self.assertIsNotNone(viewer.window)
        self.assertEqual(viewer.image_name, "test_image.jpg")
        self.assertAlmostEqual(viewer.color_var.get(), 0.1)
        self.assertAlmostEqual(viewer.area_var.get(), 0.4)

        viewer._on_close()

    def test_03_fill_ratio_caching(self):
        """fill_ratio のキャッシュが正しく動作すること"""
        from saitensamurai import StudentAnswerSheetViewer
        gray, coordinates = self._create_dummy_data()

        viewer = StudentAnswerSheetViewer(
            parent_window=self.root,
            image_name="test_image.jpg",
            gray_image=gray,
            coordinates=coordinates,
            color_threshold=0.1,
            area_threshold=0.4,
            calibrator_gui=None
        )

        # 初回描画を手動実行
        viewer._update_display()

        # キャッシュが作成されていること
        self.assertIsNotNone(viewer._cached_fill_ratios)
        self.assertAlmostEqual(viewer._cached_color_threshold, 0.1)

        # 同じ color_threshold で再度実行 → キャッシュが再利用されること
        cached_id = id(viewer._cached_fill_ratios)
        viewer._update_display()
        self.assertEqual(id(viewer._cached_fill_ratios), cached_id)

        # color_threshold を変更 → キャッシュが更新されること
        viewer.color_var.set(0.15)
        viewer._update_display()
        self.assertAlmostEqual(viewer._cached_color_threshold, 0.15)

        viewer._on_close()

    def test_04_fill_ratio_accuracy(self):
        """fill_ratio の計算が正しいこと"""
        from saitensamurai import collect_mark_fill_ratios
        gray, coordinates = self._create_dummy_data()

        ratios = collect_mark_fill_ratios(gray, coordinates, color_threshold=0.1)

        # Q1-C0 は確実にマーク有（塗りつぶし率が高い）
        q1_c0 = [r for r in ratios if r['question_no'] == 1 and r['choice'] == 0]
        self.assertEqual(len(q1_c0), 1)
        self.assertGreater(q1_c0[0]['fill_ratio'], 0.5)

        # Q1-C1 は確実にマーク無（白のまま）
        q1_c1 = [r for r in ratios if r['question_no'] == 1 and r['choice'] == 1]
        self.assertEqual(len(q1_c1), 1)
        self.assertLess(q1_c1[0]['fill_ratio'], 0.1)

    def test_05_display_update_no_crash(self):
        """_update_display が例外なく完了すること"""
        from saitensamurai import StudentAnswerSheetViewer
        gray, coordinates = self._create_dummy_data()

        viewer = StudentAnswerSheetViewer(
            parent_window=self.root,
            image_name="test_image.jpg",
            gray_image=gray,
            coordinates=coordinates,
            color_threshold=0.1,
            area_threshold=0.4,
        )

        # 表示更新が例外なく完了すること
        viewer._update_display()

        # 表示画像が設定されていること
        self.assertIsNotNone(viewer._display_image_ref)

        viewer._on_close()

    def test_06_different_thresholds(self):
        """異なる閾値で描画しても正常動作すること"""
        from saitensamurai import StudentAnswerSheetViewer
        gray, coordinates = self._create_dummy_data()

        viewer = StudentAnswerSheetViewer(
            parent_window=self.root,
            image_name="test_image.jpg",
            gray_image=gray,
            coordinates=coordinates,
            color_threshold=0.05,
            area_threshold=0.2,
        )

        viewer._update_display()
        self.assertIsNotNone(viewer._display_image_ref)

        # 閾値を変更して再描画
        viewer.color_var.set(0.3)
        viewer.area_var.set(0.6)
        viewer._update_display()
        self.assertIsNotNone(viewer._display_image_ref)

        viewer._on_close()

    def test_07_sync_methods_without_calibrator(self):
        """calibrator_gui=None の場合に同期メソッドがエラーにならないこと"""
        from saitensamurai import StudentAnswerSheetViewer
        gray, coordinates = self._create_dummy_data()

        viewer = StudentAnswerSheetViewer(
            parent_window=self.root,
            image_name="test_image.jpg",
            gray_image=gray,
            coordinates=coordinates,
            calibrator_gui=None
        )

        # 例外が起きないこと
        viewer._sync_to_calibrator()
        viewer._sync_from_calibrator()

        viewer._on_close()

    def test_08_sync_with_mock_calibrator(self):
        """calibrator_gui が設定されている場合の閾値同期"""
        from saitensamurai import StudentAnswerSheetViewer
        gray, coordinates = self._create_dummy_data()

        # モック calibrator_gui（DoubleVar を持つだけ）
        class MockCalibrator:
            def __init__(self):
                self.color_var = tk.DoubleVar(value=0.12)
                self.area_var = tk.DoubleVar(value=0.45)
                self.recollect_called = False

            def force_recollect_all(self):
                self.recollect_called = True

        mock_cal = MockCalibrator()

        viewer = StudentAnswerSheetViewer(
            parent_window=self.root,
            image_name="test_image.jpg",
            gray_image=gray,
            coordinates=coordinates,
            color_threshold=0.2,
            area_threshold=0.5,
            calibrator_gui=mock_cal
        )

        # viewer → calibrator 同期
        viewer._sync_to_calibrator()
        self.assertAlmostEqual(mock_cal.color_var.get(), 0.2)
        self.assertAlmostEqual(mock_cal.area_var.get(), 0.5)
        # force_recollect_all が呼ばれたことを確認
        self.assertTrue(mock_cal.recollect_called)

        # calibrator → viewer 同期
        mock_cal.color_var.set(0.15)
        mock_cal.area_var.set(0.35)
        viewer._sync_from_calibrator()
        self.assertAlmostEqual(viewer.color_var.get(), 0.15)
        self.assertAlmostEqual(viewer.area_var.get(), 0.35)

        viewer._on_close()

    def test_09_empty_coordinates(self):
        """座標が空の場合にエラーにならないこと"""
        from saitensamurai import StudentAnswerSheetViewer
        gray = np.ones((842, 595), dtype=np.uint8) * 230

        viewer = StudentAnswerSheetViewer(
            parent_window=self.root,
            image_name="empty_test.jpg",
            gray_image=gray,
            coordinates=[],
        )

        viewer._update_display()
        self.assertIsNotNone(viewer._display_image_ref)

        viewer._on_close()

    def test_10_edge_threshold_values(self):
        """閾値の極端な値でエラーにならないこと"""
        from saitensamurai import StudentAnswerSheetViewer
        gray, coordinates = self._create_dummy_data()

        # 最小値
        viewer = StudentAnswerSheetViewer(
            parent_window=self.root,
            image_name="edge_test.jpg",
            gray_image=gray,
            coordinates=coordinates,
            color_threshold=0.03,
            area_threshold=0.05,
        )
        viewer._update_display()
        viewer._on_close()

        # 最大値
        viewer = StudentAnswerSheetViewer(
            parent_window=self.root,
            image_name="edge_test.jpg",
            gray_image=gray,
            coordinates=coordinates,
            color_threshold=0.35,
            area_threshold=0.80,
        )
        viewer._update_display()
        viewer._on_close()

    def test_11_multiple_viewers(self):
        """複数のビューアを同時に開けること"""
        from saitensamurai import StudentAnswerSheetViewer
        gray, coordinates = self._create_dummy_data()

        viewers = []
        for i in range(3):
            v = StudentAnswerSheetViewer(
                parent_window=self.root,
                image_name=f"multi_test_{i}.jpg",
                gray_image=gray.copy(),
                coordinates=coordinates,
            )
            v._update_display()
            viewers.append(v)

        # 全てのビューアが存在すること
        self.assertEqual(len(viewers), 3)

        for v in viewers:
            v._on_close()

    def test_12_highlight_question(self):
        """highlight_question_no が指定された場合にエラーなく描画されること"""
        from saitensamurai import StudentAnswerSheetViewer
        gray, coordinates = self._create_dummy_data()

        viewer = StudentAnswerSheetViewer(
            parent_window=self.root,
            image_name="highlight_test.jpg",
            gray_image=gray,
            coordinates=coordinates,
            highlight_question_no=1,
        )
        viewer._update_display()
        self.assertIsNotNone(viewer._display_image_ref)
        self.assertEqual(viewer.highlight_question_no, 1)
        viewer._on_close()

    def test_13_highlight_nonexistent_question(self):
        """存在しない question_no を highlight しても安全なこと"""
        from saitensamurai import StudentAnswerSheetViewer
        gray, coordinates = self._create_dummy_data()

        viewer = StudentAnswerSheetViewer(
            parent_window=self.root,
            image_name="highlight_test.jpg",
            gray_image=gray,
            coordinates=coordinates,
            highlight_question_no=999,
        )
        viewer._update_display()
        self.assertIsNotNone(viewer._display_image_ref)
        viewer._on_close()


class TestCollectMarkFillRatios(unittest.TestCase):
    """collect_mark_fill_ratios 関数のテスト"""

    def test_basic_functionality(self):
        """基本的なfill_ratio計算"""
        from saitensamurai import collect_mark_fill_ratios

        # 全て白の画像
        white_img = np.ones((842, 595), dtype=np.uint8) * 255
        coords = [
            {'question_no': 1, 'question': 'Q1', 'choice': 0,
             'x': 100, 'y': 100, 'width': 20, 'height': 20},
        ]

        ratios = collect_mark_fill_ratios(white_img, coords, 0.1)
        self.assertEqual(len(ratios), 1)
        self.assertAlmostEqual(ratios[0]['fill_ratio'], 0.0, places=2)

    def test_fully_marked(self):
        """完全に塗りつぶされたマーク"""
        from saitensamurai import collect_mark_fill_ratios

        img = np.ones((842, 595), dtype=np.uint8) * 255
        # 完全に黒く塗りつぶし
        img[100:120, 100:120] = 0

        coords = [
            {'question_no': 1, 'question': 'Q1', 'choice': 0,
             'x': 100, 'y': 100, 'width': 20, 'height': 20},
        ]

        ratios = collect_mark_fill_ratios(img, coords, 0.1)
        self.assertGreater(ratios[0]['fill_ratio'], 0.9)

    def test_boundary_coordinates(self):
        """画像端の座標でもエラーにならないこと"""
        from saitensamurai import collect_mark_fill_ratios

        img = np.ones((842, 595), dtype=np.uint8) * 255
        coords = [
            {'question_no': 1, 'question': 'Q1', 'choice': 0,
             'x': 580, 'y': 830, 'width': 15, 'height': 12},
        ]

        ratios = collect_mark_fill_ratios(img, coords, 0.1)
        self.assertEqual(len(ratios), 1)


if __name__ == '__main__':
    unittest.main(verbosity=2)
