"""mask.py — 氏名欄などの黒塗りと、LLMに見せる写し（_llm_view）づくり。

約束:
- 黒塗りは LLM に見せる写しにだけ掛ける。元画像（scans/）と、採点侍が
  読み取りに使う画像には一切掛けない（マーカー検出・マーク判定と干渉させない）
- 塗りは不透明な黒。半透明は文字が透けるので使わない
- 範囲は「傾きを補正した用紙」に対する割合（0〜1）で持つ。
  マーカーのある用紙は全員の位置が揃うので、1か所決めれば全員に当たる
- 先生が全員分の塗り漏れを確かめて承認するまで、写しは作らない
"""
import hashlib
import shutil

import cv2

from omr_engine import (apply_perspective_transform, find_sheet_markers,
                        imread_unicode)
from saiten_llm.workspace import now_iso

ALIGNED_SCALE = 2.0      # 傾き補正後の大きさ 595x842 → 1190x1684
VIEW_JPEG_QUALITY = 85


def aligned_page(image):
    """マーカーがあれば傾きを補正した用紙を、無ければ元のまま返す。

    Returns: (画像, 補正したか)
    """
    markers = find_sheet_markers(image)
    if markers is None:
        return image, False
    corrected, _ = apply_perspective_transform(image, markers,
                                               output_scale=ALIGNED_SCALE)
    return corrected, True


def load_aligned(path):
    img = imread_unicode(str(path))
    if img is None:
        raise ValueError(f"画像を読めません: {path.name}")
    return aligned_page(img)


def normalize_rect(rect):
    """[x1,y1,x2,y2]（0〜1）に正規化。妥当でなければ ValueError"""
    try:
        x1, y1, x2, y2 = (min(1.0, max(0.0, float(v))) for v in rect)
    except (TypeError, ValueError):
        raise ValueError(f"範囲の形式が正しくありません: {rect}")
    x1, x2 = sorted((x1, x2))
    y1, y2 = sorted((y1, y2))
    if x2 - x1 < 0.002 or y2 - y1 < 0.002:
        raise ValueError(f"範囲が小さすぎます: {rect}")
    return [round(x1, 4), round(y1, 4), round(x2, 4), round(y2, 4)]


def rects_for(spec, filename):
    """そのページに掛ける範囲（全員共通＋そのページだけの追加）"""
    return list(spec.get("regions", [])) + list(spec.get("extra", {}).get(filename, []))


def apply_mask(image, rects):
    """範囲を不透明な黒で塗った写しを返す（元の配列は変えない）"""
    out = image.copy()
    h, w = out.shape[:2]
    for x1, y1, x2, y2 in rects:
        cv2.rectangle(out, (int(x1 * w), int(y1 * h)),
                      (int(round(x2 * w)), int(round(y2 * h))), (0, 0, 0), -1)
    return out


def encode_jpeg(image, max_width=None):
    if max_width and image.shape[1] > max_width:
        ratio = max_width / image.shape[1]
        image = cv2.resize(image, (max_width, int(image.shape[0] * ratio)),
                           interpolation=cv2.INTER_AREA)
    ok, buf = cv2.imencode(".jpg", image, [cv2.IMWRITE_JPEG_QUALITY, VIEW_JPEG_QUALITY])
    if not ok:
        raise ValueError("画像を書き出せません")
    return buf.tobytes()


def validate_spec(spec, filenames):
    """サーバーから受け取った範囲を検査して、保存できる形にする"""
    regions = [normalize_rect(r) for r in spec.get("regions", [])]
    extra = {}
    for name, rects in (spec.get("extra") or {}).items():
        if name not in filenames:
            raise ValueError(f"答案が見つかりません: {name}")
        if rects:
            extra[name] = [normalize_rect(r) for r in rects]
    return {"regions": regions, "extra": extra}


def pages_fingerprint(project):
    """答案の顔ぶれ（ファイル名・大きさ・更新日時）。承認の後で答案が変わったら気づくため"""
    h = hashlib.sha256()
    for p in project.scan_files():
        st = p.stat()
        h.update(f"{p.name}\t{st.st_size}\t{st.st_mtime_ns}\n".encode("utf-8"))
    return h.hexdigest()


def save_spec(project, spec, approved=False):
    record = dict(spec)
    record["saved_at"] = now_iso()
    if approved:
        record["approved_at"] = record["saved_at"]
        record["approved_pages"] = len(project.scan_files())
        record["approved_fingerprint"] = pages_fingerprint(project)
    project.data["mask"] = record
    project.save()


def mask_status(project):
    """黒塗りの状態: none（未設定）/ draft（未承認）/ approved / stale（承認後に答案が変わった）"""
    spec = project.data.get("mask")
    if not spec:
        return "none"
    if not spec.get("approved_at"):
        return "draft"
    if spec.get("approved_fingerprint") != pages_fingerprint(project):
        return "stale"
    return "approved"


def build_llm_view(project):
    """承認済みの黒塗りを全ページに掛け、_llm_view/pages/ に書き出す。

    Returns: 書き出した枚数
    """
    status = mask_status(project)
    if status != "approved":
        raise PermissionError(
            {"none": "黒塗りの範囲がまだ決まっていません。先に mask-server を実行してください",
             "draft": "黒塗りの範囲を先生がまだ承認していません",
             "stale": "承認の後で答案が変わりました（枚数や向きなど）。もう一度黒塗りを確かめてください",
             }[status])
    spec = project.data["mask"]
    pages_dir = project.llm_view / "pages"
    if pages_dir.exists():
        shutil.rmtree(pages_dir)
    pages_dir.mkdir(parents=True)
    count = 0
    aligned_flags = {}
    for path in project.scan_files():
        image, aligned = load_aligned(path)
        aligned_flags[path.name] = aligned
        masked = apply_mask(image, rects_for(spec, path.name))
        (pages_dir / (path.stem + ".jpg")).write_bytes(encode_jpeg(masked))
        count += 1
    (project.llm_view / "README.txt").write_text(
        "このフォルダは、LLM に見せてよい黒塗り済みの写しです。\n"
        "pages/ の画像は、傾きを補正してから氏名欄などを黒く塗ってあります。\n"
        "元の答案（scans/）は LLM に見せません。\n", encoding="utf-8")
    project.data["llm_view"] = {"built_at": now_iso(), "pages": count,
                                "unaligned": sorted(n for n, a in aligned_flags.items() if not a)}
    project.save()
    return count
