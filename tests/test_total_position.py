"""
test_total_position.py — ⑤ 合計得点デフォルト表示位置のテスト

下部マーカー間へのデフォルト配置ロジックをテストする。
"""
import unittest
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'main_src'))


class TestCalculateMarkerDefaultRegion(unittest.TestCase):
    """_calculate_marker_default_region() の単体テスト"""

    def setUp(self):
        from descriptive_scorer import _calculate_marker_default_region
        self.calc = _calculate_marker_default_region

    def test_returns_tuple(self):
        """戻り値は (x, y, w, h) のタプル"""
        result = self.calc(595, 842, 40)
        self.assertIsInstance(result, tuple)
        self.assertEqual(len(result), 4)

    def test_position_within_image_scale1(self):
        """scale=1.0 (595×842) で画像内に収まる"""
        x, y, w, h = self.calc(595, 842, 40)
        self.assertGreaterEqual(x, 0)
        self.assertGreaterEqual(y, 0)
        self.assertLess(x + w, 595)
        self.assertLess(y + h, 842)

    def test_position_within_image_scale2(self):
        """scale=2.0 (1190×1684) で画像内に収まる"""
        x, y, w, h = self.calc(1190, 1684, 80)
        self.assertGreaterEqual(x, 0)
        self.assertGreaterEqual(y, 0)
        self.assertLess(x + w, 1190)
        self.assertLess(y + h, 1684)

    def test_x_between_markers(self):
        """ボックスが左右マーカーの間に収まる"""
        from constants import MARKER_X_FRAC_LEFT, MARKER_X_FRAC_RIGHT
        from descriptive_scorer import _MARKER_HALF_SIZE_FRAC
        w, h = 595, 842
        x, y, bw, bh = self.calc(w, h, 40)
        left_inner = (MARKER_X_FRAC_LEFT + _MARKER_HALF_SIZE_FRAC) * w
        right_inner = (MARKER_X_FRAC_RIGHT - _MARKER_HALF_SIZE_FRAC) * w
        self.assertGreaterEqual(x, left_inner - 1)  # 小数点の丸め許容
        self.assertLessEqual(x + bw, right_inner + 1)

    def test_y_centered_on_marker(self):
        """Y中心がマーカーの中心付近にある"""
        from constants import MARKER_Y_FRAC_BOTTOM
        w, h = 595, 842
        box_h = 40
        x, y, bw, bh = self.calc(w, h, box_h)
        marker_cy = MARKER_Y_FRAC_BOTTOM * h
        box_cy = y + bh / 2
        self.assertAlmostEqual(box_cy, marker_cy, delta=2)

    def test_proportional_scaling(self):
        """画像スケールが変わっても比率が維持される"""
        x1, y1, w1, h1 = self.calc(595, 842, 40)
        x2, y2, w2, h2 = self.calc(1190, 1684, 80)
        # 比率はほぼ同じ（整数丸めの誤差許容）
        self.assertAlmostEqual(x1 / 595, x2 / 1190, delta=0.01)
        self.assertAlmostEqual(y1 / 842, y2 / 1684, delta=0.01)
        self.assertAlmostEqual(w1 / 595, w2 / 1190, delta=0.01)

    def test_width_fills_marker_span(self):
        """ボックス幅がマーカー間の利用可能幅いっぱいになる"""
        from constants import MARKER_X_FRAC_LEFT, MARKER_X_FRAC_RIGHT
        from descriptive_scorer import _MARKER_HALF_SIZE_FRAC
        x, y, w, h = self.calc(595, 842, 40)
        left_inner = (MARKER_X_FRAC_LEFT + _MARKER_HALF_SIZE_FRAC + 0.005) * 595
        right_inner = (MARKER_X_FRAC_RIGHT - _MARKER_HALF_SIZE_FRAC - 0.005) * 595
        expected_w = int(right_inner - left_inner)
        self.assertEqual(w, expected_w)

    def test_bottom_region_not_too_high(self):
        """Y位置が画像下部（80%以降）にある"""
        x, y, w, h = self.calc(595, 842, 40)
        self.assertGreater(y, 842 * 0.80)

    def test_left_aligned(self):
        """ボックスは左寄せ（マーカー内側すぐ）"""
        from constants import MARKER_X_FRAC_LEFT
        from descriptive_scorer import _MARKER_HALF_SIZE_FRAC
        x, y, w, h = self.calc(595, 842, 40)
        expected_left = int((MARKER_X_FRAC_LEFT + _MARKER_HALF_SIZE_FRAC + 0.005) * 595)
        self.assertAlmostEqual(x, expected_left, delta=2)


class TestSelectTotalPositionSignature(unittest.TestCase):
    """select_total_position のシグネチャが正しいことを確認"""

    def test_has_use_marker_default_param(self):
        """use_marker_default パラメータが存在する"""
        import inspect
        from descriptive_scorer import select_total_position
        sig = inspect.signature(select_total_position)
        self.assertIn('use_marker_default', sig.parameters)
        # デフォルトは True
        self.assertEqual(
            sig.parameters['use_marker_default'].default, True
        )


class TestMarkerConstants(unittest.TestCase):
    """マーカー定数の値が仕様通りであることを確認

    定数は constants.py に一元化され、omr_engine / mark_checker /
    descriptive_scorer がすべて同じ定数を参照する(値の食い違いは構造上
    起こらない)。ここでは定数自体が仕様値からズレていないことを固定する。
    """

    def test_marker_fractions_spec_values(self):
        """マーカー比率が仕様値(Mark2様式)と一致"""
        from constants import (
            MARKER_X_FRAC_LEFT, MARKER_X_FRAC_RIGHT,
            MARKER_Y_FRAC_TOP, MARKER_Y_FRAC_BOTTOM,
        )
        self.assertAlmostEqual(MARKER_X_FRAC_LEFT, 0.155, places=3)
        self.assertAlmostEqual(MARKER_X_FRAC_RIGHT, 0.845, places=3)
        self.assertAlmostEqual(MARKER_Y_FRAC_TOP, 0.04, places=3)
        self.assertAlmostEqual(MARKER_Y_FRAC_BOTTOM, 0.96, places=3)


if __name__ == "__main__":
    unittest.main()
