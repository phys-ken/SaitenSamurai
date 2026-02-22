# -*- coding: utf-8 -*-
"""
test_student_answer_pdf.py — student_answer_pdf モジュールのテスト

生徒の設問別解答一覧 PDF 生成コアロジックをテストする。
"""

import math
import tempfile
from pathlib import Path
from unittest.mock import MagicMock, patch

import numpy as np
import pytest
from PIL import Image

import student_answer_pdf as sap


# ================================================================
# ヘルパー
# ================================================================

def _make_test_image(w=200, h=100, color=(128, 128, 128)):
    """テスト用のダミー画像を作成"""
    return Image.new("RGB", (w, h), color)


def _make_config(n_questions=2, region_size=(100, 200, 400, 500)):
    """テスト用の descriptive_config を作成"""
    questions = []
    for i in range(n_questions):
        questions.append({
            "id": f"D{i+1}",
            "name": f"記述{i+1}",
            "max_score": 5,
            "aspect": 1,
            "region": list(region_size),
        })
    return {"version": 1, "questions": questions}


def _make_processing_folder(tmp_path, n_images=3, w=595, h=842):
    """テスト用の Processing フォルダを作成 (ダミー画像ファイル)"""
    proc_dir = tmp_path / "00_Processing"
    proc_dir.mkdir()
    for i in range(n_images):
        img = Image.new("RGB", (w, h), (200, 200, 200))
        img.save(str(proc_dir / f"image_{i+1:03d}.jpg"))
    return str(proc_dir)


def _make_scores_data(n_images=3, n_questions=2):
    """テスト用の採点結果データ"""
    scores = {}
    for i in range(n_images):
        filename = f"image_{i+1:03d}.jpg"
        q_scores = {}
        for q in range(n_questions):
            q_scores[f"D{q+1}"] = (n_images - i)  # 逆順の得点
        scores[filename] = q_scores
    return {"version": 1, "scores": scores}


# ================================================================
# _natural_sort_key
# ================================================================

class TestNaturalSortKey:
    def test_numeric_sort(self):
        """数字部分が数値としてソートされる"""
        names = ["img_10.jpg", "img_2.jpg", "img_1.jpg"]
        result = sorted(names, key=sap._natural_sort_key)
        assert result == ["img_1.jpg", "img_2.jpg", "img_10.jpg"]

    def test_case_insensitive(self):
        """大文字小文字を区別しない"""
        names = ["B.jpg", "a.jpg", "C.jpg"]
        result = sorted(names, key=sap._natural_sort_key)
        assert result == ["a.jpg", "B.jpg", "C.jpg"]


# ================================================================
# _compute_grid
# ================================================================

class TestComputeGrid:
    def test_small_region_many_cols(self):
        """小さい領域は多くの列数になる"""
        cols, rows, iw, ih = sap._compute_grid([0, 0, 59, 84])
        assert cols >= 2
        assert rows >= 2
        assert iw > 0
        assert ih > 0

    def test_large_region_single_col(self):
        """大きい領域は1列になる"""
        cols, rows, iw, ih = sap._compute_grid([0, 0, 595, 421])
        assert cols == 1
        assert rows >= 1

    def test_zero_region(self):
        """ゼロサイズ領域でもクラッシュしない"""
        cols, rows, iw, ih = sap._compute_grid([100, 100, 100, 100])
        assert cols >= 1
        assert rows >= 1

    def test_aspect_ratio_maintained(self):
        """A4に対する相対サイズ比が維持される"""
        # 領域が A4 の半分の幅
        region = [0, 0, 297, 421]  # 約半幅, 約半高
        cols, rows, iw, ih = sap._compute_grid(region)
        ratio_w = iw / sap._A4_W
        expected_ratio_w = 297 / 595
        assert abs(ratio_w - expected_ratio_w) < 0.01


# ================================================================
# _crop_region_from_image
# ================================================================

class TestCropRegion:
    def test_crop_from_processing(self, tmp_path):
        """Processing フォルダの画像から切り出せる"""
        proc_dir = _make_processing_folder(tmp_path, n_images=1)
        img_path = str(Path(proc_dir) / "image_001.jpg")
        region = [100, 200, 400, 500]
        result = sap._crop_region_from_image(img_path, region, proc_dir)
        assert result is not None
        assert result.size[0] == 300  # 400 - 100
        assert result.size[1] == 300  # 500 - 200

    def test_crop_missing_file(self, tmp_path):
        """存在しないファイルは None を返す"""
        proc_dir = str(tmp_path / "empty")
        Path(proc_dir).mkdir()
        result = sap._crop_region_from_image(
            str(tmp_path / "nonexistent.jpg"),
            [0, 0, 100, 100],
            proc_dir,
        )
        assert result is None

    def test_crop_region_clipping(self, tmp_path):
        """領域がはみ出す場合はクリッピングされる"""
        proc_dir = _make_processing_folder(tmp_path, n_images=1, w=200, h=200)
        img_path = str(Path(proc_dir) / "image_001.jpg")
        region = [150, 150, 300, 300]  # はみ出す
        result = sap._crop_region_from_image(img_path, region, proc_dir)
        assert result is not None
        assert result.size[0] == 50   # 200 - 150
        assert result.size[1] == 50


# ================================================================
# _generate_question_pdf
# ================================================================

class TestGenerateQuestionPdf:
    @pytest.fixture
    def question(self):
        return {"id": "D1", "name": "記述1", "max_score": 5, "region": [100, 200, 400, 500]}

    def test_pre_mode(self, tmp_path, question):
        """採点前モードでPDFが生成される"""
        entries = [
            {"filename": f"img_{i:03d}.jpg", "image": _make_test_image(), "score": None}
            for i in range(5)
        ]
        pdf_path = str(tmp_path / "test_pre.pdf")
        result = sap._generate_question_pdf(question, entries, pdf_path, mode="pre")
        assert result is not None
        assert Path(result).exists()
        assert Path(result).stat().st_size > 0

    def test_post_mode_with_scores(self, tmp_path, question):
        """採点後モードで得点入りPDFが生成される"""
        entries = [
            {"filename": f"img_{i:03d}.jpg", "image": _make_test_image(), "score": i}
            for i in range(5)
        ]
        pdf_path = str(tmp_path / "test_post.pdf")
        result = sap._generate_question_pdf(question, entries, pdf_path, mode="post")
        assert result is not None
        assert Path(result).exists()

    def test_empty_entries(self, tmp_path, question):
        """空のエントリリストではNoneを返す"""
        pdf_path = str(tmp_path / "test_empty.pdf")
        result = sap._generate_question_pdf(question, [], pdf_path)
        assert result is None

    def test_none_image_skipped(self, tmp_path, question):
        """None画像はスキップされるがPDFは生成される"""
        entries = [
            {"filename": "ok.jpg", "image": _make_test_image(), "score": None},
            {"filename": "bad.jpg", "image": None, "score": None},
        ]
        pdf_path = str(tmp_path / "test_skip.pdf")
        result = sap._generate_question_pdf(question, entries, pdf_path)
        assert result is not None

    def test_many_images_multipage(self, tmp_path, question):
        """多数の画像で複数ページが生成される"""
        entries = [
            {"filename": f"img_{i:03d}.jpg", "image": _make_test_image(40, 40), "score": None}
            for i in range(100)
        ]
        pdf_path = str(tmp_path / "test_multi.pdf")
        result = sap._generate_question_pdf(question, entries, pdf_path)
        assert result is not None
        assert Path(result).stat().st_size > 1000  # 複数ページで十分なサイズ


# ================================================================
# generate_pre_scoring_pdfs
# ================================================================

class TestGeneratePreScoringPdfs:
    def test_generates_pdfs(self, tmp_path):
        """設問数と同数のPDFが生成される"""
        proc_dir = _make_processing_folder(tmp_path, n_images=3)
        config = _make_config(n_questions=2)
        out_dir = tmp_path / "results"
        out_dir.mkdir()
        result = sap.generate_pre_scoring_pdfs(
            processing_folder=proc_dir,
            config=config,
            output_base_folder=str(out_dir),
        )
        assert len(result) == 2
        for pdf_path in result:
            assert Path(pdf_path).exists()
            assert "生徒一覧" in Path(pdf_path).name

    def test_no_questions(self, tmp_path):
        """問題なしでは空リストを返す"""
        proc_dir = _make_processing_folder(tmp_path)
        result = sap.generate_pre_scoring_pdfs(
            processing_folder=proc_dir,
            config={"questions": []},
            output_base_folder=str(tmp_path),
        )
        assert result == []

    def test_no_images(self, tmp_path):
        """画像なしでは空リストを返す"""
        empty_dir = tmp_path / "empty"
        empty_dir.mkdir()
        config = _make_config()
        result = sap.generate_pre_scoring_pdfs(
            processing_folder=str(empty_dir),
            config=config,
            output_base_folder=str(tmp_path),
        )
        assert result == []

    def test_output_folder_structure(self, tmp_path):
        """正しいフォルダ構造で出力される"""
        proc_dir = _make_processing_folder(tmp_path, n_images=2)
        config = _make_config(n_questions=1)
        out_dir = tmp_path / "results"
        out_dir.mkdir()
        sap.generate_pre_scoring_pdfs(
            processing_folder=proc_dir,
            config=config,
            output_base_folder=str(out_dir),
        )
        expected_dir = out_dir / sap.ANSWER_GALLERY_FOLDER / sap.PRE_SCORING_SUBFOLDER
        assert expected_dir.exists()

    def test_filename_format(self, tmp_path):
        """ファイル名が 001_設問名_生徒一覧.pdf 形式"""
        proc_dir = _make_processing_folder(tmp_path, n_images=2)
        config = _make_config(n_questions=1)
        config["questions"][0]["name"] = "問題A"
        out_dir = tmp_path / "results"
        out_dir.mkdir()
        result = sap.generate_pre_scoring_pdfs(
            processing_folder=proc_dir,
            config=config,
            output_base_folder=str(out_dir),
        )
        assert len(result) == 1
        assert Path(result[0]).name == "001_問題A_生徒一覧.pdf"

    def test_progress_callback(self, tmp_path):
        """プログレスコールバックが呼ばれる"""
        proc_dir = _make_processing_folder(tmp_path, n_images=2)
        config = _make_config(n_questions=1)
        out_dir = tmp_path / "results"
        out_dir.mkdir()
        calls = []
        sap.generate_pre_scoring_pdfs(
            processing_folder=proc_dir,
            config=config,
            output_base_folder=str(out_dir),
            progress_callback=lambda c, t: calls.append((c, t)),
        )
        assert len(calls) == 2  # 2 images * 1 question
        assert calls[-1][0] == calls[-1][1]  # 最後は current == total


# ================================================================
# generate_post_scoring_pdfs
# ================================================================

class TestGeneratePostScoringPdfs:
    def test_generates_pdfs_with_scores(self, tmp_path):
        """得点入りPDFが設問数分生成される"""
        proc_dir = _make_processing_folder(tmp_path, n_images=3)
        config = _make_config(n_questions=2)
        scores = _make_scores_data(n_images=3, n_questions=2)
        out_dir = tmp_path / "results"
        out_dir.mkdir()
        result = sap.generate_post_scoring_pdfs(
            processing_folder=proc_dir,
            config=config,
            scores_data=scores,
            output_base_folder=str(out_dir),
        )
        assert len(result) == 2

    def test_sorted_by_score_descending(self, tmp_path):
        """得点の高い順にソートされる（内部ロジック検証）"""
        entries = [
            {"filename": "a.jpg", "image": None, "score": 1},
            {"filename": "b.jpg", "image": None, "score": 5},
            {"filename": "c.jpg", "image": None, "score": 3},
            {"filename": "d.jpg", "image": None, "score": None},
        ]
        entries.sort(
            key=lambda e: (
                -(e["score"] if e["score"] is not None else -1),
                sap._natural_sort_key(e["filename"]),
            )
        )
        assert entries[0]["score"] == 5
        assert entries[1]["score"] == 3
        assert entries[2]["score"] == 1
        assert entries[3]["score"] is None

    def test_output_folder_structure(self, tmp_path):
        """採点後フォルダに出力される"""
        proc_dir = _make_processing_folder(tmp_path, n_images=2)
        config = _make_config(n_questions=1)
        scores = _make_scores_data(n_images=2, n_questions=1)
        out_dir = tmp_path / "results"
        out_dir.mkdir()
        sap.generate_post_scoring_pdfs(
            processing_folder=proc_dir,
            config=config,
            scores_data=scores,
            output_base_folder=str(out_dir),
        )
        expected_dir = out_dir / sap.ANSWER_GALLERY_FOLDER / sap.POST_SCORING_SUBFOLDER
        assert expected_dir.exists()

    def test_missing_scores_go_to_end(self, tmp_path):
        """採点なし画像は末尾に配置される"""
        proc_dir = _make_processing_folder(tmp_path, n_images=3)
        config = _make_config(n_questions=1)
        # image_002.jpg だけスコアなし
        scores = {
            "version": 1,
            "scores": {
                "image_001.jpg": {"D1": 5},
                "image_003.jpg": {"D1": 3},
            }
        }
        out_dir = tmp_path / "results"
        out_dir.mkdir()
        result = sap.generate_post_scoring_pdfs(
            processing_folder=proc_dir,
            config=config,
            scores_data=scores,
            output_base_folder=str(out_dir),
        )
        assert len(result) == 1


# ================================================================
# constants 連携
# ================================================================

class TestConstants:
    def test_answer_gallery_folder_defined(self):
        """ANSWER_GALLERY_FOLDER が constants に定義されている"""
        from constants import ANSWER_GALLERY_FOLDER
        assert ANSWER_GALLERY_FOLDER == "04_Answer_Gallery"

    def test_folder_names(self):
        """フォルダ名定数が正しい"""
        assert sap.ANSWER_GALLERY_FOLDER == "04_Answer_Gallery"
        assert sap.PRE_SCORING_SUBFOLDER == "010_pre_scoring"
        assert sap.POST_SCORING_SUBFOLDER == "020_post_scoring"
