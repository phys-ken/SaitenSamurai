"""
fit_image_to_display 関数のユニットテスト

マークチェック画像が表示領域に収まるようリサイズされることを検証する。
"""
import pytest
from PIL import Image

from saitensamurai import fit_image_to_display, MAX_DISPLAY_WIDTH, MAX_DISPLAY_HEIGHT


class TestFitImageToDisplay:
    """fit_image_to_display のテスト"""

    def test_image_within_bounds_unchanged(self):
        """表示領域内の画像はリサイズされない"""
        img = Image.new("RGB", (800, 200))
        result = fit_image_to_display(img)
        assert result.size == (800, 200)

    def test_wide_image_scaled_down(self):
        """幅が超過する画像は縮小される（10選択肢分を想定）"""
        # 実測値: 1655px 幅の画像（選択肢10個分のマーク領域）
        img = Image.new("RGB", (1655, 130))
        result = fit_image_to_display(img)
        assert result.width <= MAX_DISPLAY_WIDTH
        assert result.height <= MAX_DISPLAY_HEIGHT

    def test_aspect_ratio_preserved(self):
        """アスペクト比が維持される"""
        img = Image.new("RGB", (2000, 100))
        result = fit_image_to_display(img)
        original_ratio = 2000 / 100
        result_ratio = result.width / result.height
        assert abs(original_ratio - result_ratio) < 1.0  # 整数丸め分の許容

    def test_tall_image_scaled_by_height(self):
        """高さが制約になる場合も正しくリサイズされる"""
        img = Image.new("RGB", (500, 800))
        result = fit_image_to_display(img)
        assert result.height <= MAX_DISPLAY_HEIGHT
        assert result.width <= MAX_DISPLAY_WIDTH

    def test_exact_boundary_no_resize(self):
        """ちょうど境界サイズの画像はリサイズされない"""
        img = Image.new("RGB", (MAX_DISPLAY_WIDTH, MAX_DISPLAY_HEIGHT))
        result = fit_image_to_display(img)
        assert result.size == (MAX_DISPLAY_WIDTH, MAX_DISPLAY_HEIGHT)

    def test_custom_max_dimensions(self):
        """カスタム最大サイズが指定できる"""
        img = Image.new("RGB", (600, 300))
        result = fit_image_to_display(img, max_width=400, max_height=200)
        assert result.width <= 400
        assert result.height <= 200

    def test_very_small_image_not_enlarged(self):
        """小さい画像は拡大されない"""
        img = Image.new("RGB", (100, 50))
        result = fit_image_to_display(img)
        assert result.size == (100, 50)

    def test_realistic_mark_area_dimensions(self):
        """実際のマーク領域サイズ（10選択肢・3446x4871画像）で全選択肢が表示される
        
        bbox w=176 (base) → 1019px (actual) → 1324px (expand) → 1655px (scale)
        bbox h=14 (base) → 81px (actual) → 105px (expand) → 131px (scale)
        """
        img = Image.new("RGB", (1655, 131))
        result = fit_image_to_display(img)
        # 幅が制約: ratio = 1100/1655 ≈ 0.665
        assert result.width == MAX_DISPLAY_WIDTH
        expected_height = max(int(131 * (MAX_DISPLAY_WIDTH / 1655)), 1)
        assert result.height == expected_height

    def test_minimum_size_one_pixel(self):
        """極端に大きい画像でも最小1ピクセルが保証される"""
        img = Image.new("RGB", (100000, 1))
        result = fit_image_to_display(img)
        assert result.width >= 1
        assert result.height >= 1
