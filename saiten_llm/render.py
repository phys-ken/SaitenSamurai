"""render.py — 返却答案（個票）の見た目の設定と、黒塗り済みの写しでの確かめ。

  render-settings  採点侍の描画の設定（観点を出すか・得点を出すか など）
  total-region     合計点を書く枠（「得点」欄）を探す・設定する
  render-preview   黒塗り済みの写しの上に、いまの設定で個票を描いてみる

LLM は個票の実物（元画像から作る）を開けないので、見た目は render-preview の
画像で確かめる。描くのは採点侍の関数そのもの（render_descriptive_only_image）なので、
確かめた見た目と実物はずれない。
"""
import cv2

from saiten_llm.imageio import read_image, write_image

TRUE_WORDS = {"true", "1", "yes", "on", "はい"}
FALSE_WORDS = {"false", "0", "no", "off", "いいえ"}


# --- 描画の設定 -----------------------------------------------------------

def parse_setting(text):
    """「キー=値」を (キー, 値) にする。値の型は既定値に合わせる。

    Raises:
        ValueError: 形が違う・知らないキー・型が合わない
    """
    from constants import DEFAULT_RENDERING_SETTINGS
    if "=" not in text:
        raise ValueError(f"「キー=値」の形で書いてください: {text}")
    key, value = (s.strip() for s in text.split("=", 1))
    if key not in DEFAULT_RENDERING_SETTINGS:
        raise ValueError(f"知らない設定です: {key}（使える設定は引数なしの render-settings で表示）")
    default = DEFAULT_RENDERING_SETTINGS[key]
    if isinstance(default, bool):
        word = value.lower()
        if word in TRUE_WORDS:
            return key, True
        if word in FALSE_WORDS:
            return key, False
        raise ValueError(f"{key} は true か false で書いてください: {value}")
    try:
        number = float(value)
    except ValueError:
        raise ValueError(f"{key} は数で書いてください: {value}") from None
    if key == "descriptive_opacity" and not 0 <= number <= 1:
        raise ValueError("descriptive_opacity は 0〜1 で書いてください")
    return key, number


# --- 「得点」欄を探す -----------------------------------------------------

def _box_candidates(gray, x0, y_end):
    """右上の範囲にある四角の「内側」の一覧 [(x1, y1, x2, y2), ...]（画像全体の座標）。

    外形でなく内側（穴）を使うのは、黒塗りした氏名欄が隣の枠に接していても、
    枠の内側は別に取れるから
    """
    sub = gray[:y_end, x0:]
    _, bw = cv2.threshold(sub, 140, 255, cv2.THRESH_BINARY_INV)
    contours, hierarchy = cv2.findContours(bw, cv2.RETR_CCOMP, cv2.CHAIN_APPROX_SIMPLE)
    boxes = []
    if hierarchy is None:
        return boxes
    for c, h in zip(contours, hierarchy[0]):
        if h[3] < 0:          # 親が無い＝外形。内側（穴）だけ使う
            continue
        x, y, w, hh = cv2.boundingRect(c)
        boxes.append((x + x0, y, x + x0 + w, y + hh))
    return boxes


def _with_merged_pairs(boxes, gap):
    """縦の線で2つに分かれて見える欄を、1つにつないだものも候補に足す。

    つなぐのは、上下がそろって左右に接する2つ。つながない元の2つも候補に残す
    （隣の氏名欄などとつながって大きすぎる候補になっても、元の欄が残るように）
    """
    out = list(boxes)
    for a in boxes:
        for b in boxes:
            height = min(a[3] - a[1], b[3] - b[1])
            overlap = min(a[3], b[3]) - max(a[1], b[1])
            if 0 <= b[0] - a[2] <= gap and overlap >= 0.8 * height:
                out.append((a[0], min(a[1], b[1]), b[2], max(a[3], b[3])))
    return out


def detect_total_box(image):
    """答案の右上の「得点」欄を探す（黒塗り済みの写しで）。

    条件: 右上の範囲にあり、右端が用紙の右 10% に入り、大きさが欄らしく、いちばん上にあるもの。
    いちばん大きい四角を選ぶと、右の列の解答欄の上の余白を取り違える（実際にあった）。

    Returns:
        (x1, y1, x2, y2)（写しの画素）。見つからなければ None
    """
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY) if image.ndim == 3 else image
    h, w = gray.shape
    landscape = w > h
    # 縦長の答案は欄が右上の隅から離れていることがあるので、探す範囲を広げる
    x0 = int(w * (0.80 if landscape else 0.60))
    y_end = int(h * (0.15 if landscape else 0.20))
    boxes = _with_merged_pairs(_box_candidates(gray, x0, y_end), gap=max(6, w // 150))
    found = []
    row = max(1, h // 100)        # 上端がこの差に収まるものは同じ高さとみなす
    for x1, y1, x2, y2 in boxes:
        bw, bh = x2 - x1, y2 - y1
        if (0.03 * w <= bw <= 0.25 * w and 0.015 * h <= bh <= 0.10 * h
                and x2 >= 0.90 * w and bw > bh):
            # いちばん上 → 右端が右 → 幅が広い（2つに分かれた欄はつないだ方）
            found.append(((y1 // row, -x2, -bw), (x1, y1, x2, y2)))
    if not found:
        return None
    return min(found)[1]


def lower_half(box):
    """欄の下半分。合計点は枠の上端に描かれるので、見出し（「得点」）を避ける"""
    x1, y1, x2, y2 = box
    return (x1, (y1 + y2) // 2, x2, y2)


def draw_box_overlay(image, box, region):
    """見つけた欄（緑）と、合計点を書く枠（赤）を重ねた画像"""
    out = image.copy()
    if box:
        cv2.rectangle(out, box[:2], box[2:], (0, 160, 0), 3)
    if region:
        cv2.rectangle(out, tuple(region[:2]), tuple(region[2:]), (0, 0, 230), 3)
    return out


def default_crop_rect(size):
    """--crop の既定の範囲: 右上の隅（「得点」欄がよくある所）。縦長は左へ広め"""
    w, h = size
    landscape = w > h
    return (int(w * (0.60 if landscape else 0.40)), 0, w, int(h * (0.20 if landscape else 0.18)))


def around(rect, size, pad=150):
    """枠の周りを pad 画素ずつ広げた範囲（答案の外には出さない）"""
    w, h = size
    x1, y1, x2, y2 = rect
    return (max(0, x1 - pad), max(0, y1 - pad), min(w, x2 + pad), min(h, y2 + pad))


GRID_COLOR = (230, 150, 40)        # 目盛りの線（BGR・青）
RULER = (72, 34)                   # 目盛りの数字を書く余白（左, 上）


def grid_crop(image, rect, step=50, target_width=1600, boxes=()):
    """範囲を切り出して拡大し、原寸の画素値の目盛りと格子を重ねる。

    エージェントが、表示の倍率を計算せずに、画像に書かれた数をそのまま読めるようにする。

    Args:
        image: 採点侍の答案と同じ大きさの画像（原寸の画素）
        rect: 切り出す範囲 (x1, y1, x2, y2)（原寸の画素）
        step: 格子の間隔（原寸の画素）
        boxes: 重ねる枠 [((x1, y1, x2, y2), BGR色), ...]（原寸の画素）

    Returns:
        (画像, ticks) — ticks は {"x": [(原寸の値, 画像上の位置), ...], "y": [...]}
    """
    h, w = image.shape[:2]
    x1, y1 = max(0, int(rect[0])), max(0, int(rect[1]))
    x2, y2 = min(w, int(rect[2])), min(h, int(rect[3]))
    if x2 - x1 < 10 or y2 - y1 < 10:
        raise ValueError(f"切り出す範囲が答案（{w}x{h}）の外か、小さすぎます: {rect}")
    scale = max(1.0, min(4.0, target_width / (x2 - x1)))
    sub = cv2.resize(image[y1:y2, x1:x2], (round((x2 - x1) * scale), round((y2 - y1) * scale)),
                     interpolation=cv2.INTER_LINEAR)
    left, top = RULER
    canvas = cv2.copyMakeBorder(sub, top, 0, left, 0, cv2.BORDER_CONSTANT, value=(255, 255, 255))

    def ox(v):
        return left + round((v - x1) * scale)

    def oy(v):
        return top + round((v - y1) * scale)

    ticks = {"x": [], "y": []}
    font = cv2.FONT_HERSHEY_SIMPLEX
    for v in range(-(-x1 // step) * step, x2 + 1, step):
        X = ox(v)
        cv2.line(canvas, (X, top), (X, canvas.shape[0] - 1), GRID_COLOR, 2 if v % 100 == 0 else 1)
        (tw, _), _ = cv2.getTextSize(str(v), font, 0.45, 1)
        cv2.putText(canvas, str(v), (X - tw // 2, top - 10), font, 0.45, (0, 0, 0), 1, cv2.LINE_AA)
        ticks["x"].append((v, X))
    for v in range(-(-y1 // step) * step, y2 + 1, step):
        Y = oy(v)
        cv2.line(canvas, (left, Y), (canvas.shape[1] - 1, Y), GRID_COLOR, 2 if v % 100 == 0 else 1)
        cv2.putText(canvas, str(v), (4, Y + 5), font, 0.45, (0, 0, 0), 1, cv2.LINE_AA)
        ticks["y"].append((v, Y))
    for (bx1, by1, bx2, by2), color in boxes:
        cv2.rectangle(canvas, (ox(bx1), oy(by1)), (ox(bx2), oy(by2)), color, 3)
    # 拡大の倍率は書かない（倍率を掛けて読み違えないように。目盛りの数がそのまま原寸）
    return canvas, ticks


# --- プレビュー -----------------------------------------------------------

def pick_pages(totals, count):
    """合計点の低い・中ぐらい・高い答案を、合計点の順に均等に count 枚選ぶ"""
    ordered = sorted(totals, key=lambda f: (totals[f], f))
    if count >= len(ordered):
        return ordered
    if count == 1:
        return [ordered[len(ordered) // 2]]
    step = (len(ordered) - 1) / (count - 1)
    picked = []
    for i in range(count):
        f = ordered[round(i * step)]
        if f not in picked:
            picked.append(f)
    return picked


def render_preview(view_image, boxed_size, config, scores, rendering_settings):
    """黒塗り済みの写しに、採点侍の描画の関数で個票を描く。

    写しの大きさが採点侍の画像（00_Processing）と違えば（マーカーのある用紙は2倍）、
    採点侍の大きさに合わせてから描く。座標は採点侍のものをそのまま使うため
    """
    from descriptive_scorer import render_descriptive_only_image
    if (view_image.shape[1], view_image.shape[0]) != tuple(boxed_size):
        view_image = cv2.resize(view_image, tuple(boxed_size), interpolation=cv2.INTER_AREA)
    return render_descriptive_only_image(view_image, config, scores, rendering_settings)


def view_page(project, filename):
    from pathlib import Path
    path = project.llm_view / "pages" / (Path(filename).stem + ".jpg")
    image = read_image(path) if path.exists() else None
    if image is None:
        raise FileNotFoundError(f"黒塗り済みの写しがありません: {path.name}")
    return image


def save_jpeg(path, image):
    path.parent.mkdir(parents=True, exist_ok=True)
    write_image(path, image)
    return path
