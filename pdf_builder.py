"""
施工前・施工後 比較報告書 PDF生成モジュール
A4縦 / 1ページに5組(10枚)のペア写真を配置する
"""
from fpdf import FPDF
from PIL import Image, ImageOps
import os
import tempfile

# ---- レイアウト定数 (mm) ----
PAGE_W = 210
PAGE_H = 297
MARGIN = 10
TITLE_H = 15
LABEL_H = 8
PAIRS_PER_PAGE = 5

# ---- 日本語フォント設定 ----
# fpdf2は標準では日本語を描画できないため、UTF-8対応のTTFフォントが必要です。
# fonts/ フォルダに IPAexゴシック等のフォントファイルを配置してください。
FONT_DIR = os.path.join(os.path.dirname(__file__), "fonts")
FONT_PATH = os.path.join(FONT_DIR, "ipaexg.ttf")
FONT_NAME = "JPFont"


class ReportPDF(FPDF):
    def __init__(self, property_name):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.property_name = property_name
        self.set_auto_page_break(auto=False)
        if not os.path.exists(FONT_PATH):
            raise FileNotFoundError(
                f"日本語フォントが見つかりません: {FONT_PATH}\n"
                "fonts/ フォルダに ipaexg.ttf (IPAexゴシック等) を配置してください。"
            )
        self.add_font(FONT_NAME, "", FONT_PATH, uni=True)

    def draw_title(self):
        self.set_xy(0, MARGIN)
        self.set_font(FONT_NAME, "", 16)
        self.cell(PAGE_W, TITLE_H, self.property_name, align="C")

    def draw_page_labels(self, top_y, col_w, gap):
        self.set_font(FONT_NAME, "", 12)
        left_x = MARGIN
        right_x = MARGIN + col_w + gap
        self.set_xy(left_x, top_y)
        self.cell(col_w, LABEL_H, "施工前", align="C")
        self.set_xy(right_x, top_y)
        self.cell(col_w, LABEL_H, "施工後", align="C")


def _load_image(img_path):
    """画像を読み込み、EXIFの向き情報(スマホ撮影で付与される回転情報)を
    ピクセルに反映してから返す。これによりPDF埋め込み時の意図しない回転を防ぐ"""
    im = Image.open(img_path)
    im = ImageOps.exif_transpose(im)  # EXIFの回転情報を実際のピクセルに反映
    if im.mode not in ("RGB", "L"):
        im = im.convert("RGB")
    return im


def _load_image_as_temp_png(img_path, tmp_files):
    """向き補正した画像をPNGとして一時ファイルに保存し、そのパスを返す。
    fpdf2はEXIF情報が残っていると独自にもう一度回転させてしまうため、
    向き補正後に一度PNGとして保存し直すことでEXIF情報自体を取り除く"""
    im = _load_image(img_path)
    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    im.save(tmp.name)
    tmp.close()
    tmp_files.append(tmp.name)
    size = im.size
    im.close()
    return tmp.name, size


def _fit_size(size, max_w, max_h):
    """画像の縦横比を保ったまま、指定した枠(max_w x max_h)に収まるサイズを計算する"""
    w, h = size
    ratio = min(max_w / w, max_h / h)
    return w * ratio, h * ratio


def build_pdf(property_name, pairs, output_path):
    """
    property_name: str  物件名
    pairs: [(施工前画像パス, 施工後画像パス), ...]  最大枚数の制限なし(自動改ページ)
    output_path: 出力するPDFのパス
    """
    pdf = ReportPDF(property_name)

    usable_w = PAGE_W - 2 * MARGIN
    gap = 20  # mm 矢印を描くための中央スペース
    col_w = (usable_w - gap) / 2
    row_h = (PAGE_H - 2 * MARGIN - TITLE_H - LABEL_H) / PAIRS_PER_PAGE
    pad = 2  # mm 画像とセル枠の間の余白

    tmp_files = []  # 最後にまとめて削除する一時PNGファイル

    try:
        content_top = None
        for i, (before, after) in enumerate(pairs):
            pos_in_page = i % PAIRS_PER_PAGE
            if pos_in_page == 0:
                pdf.add_page()
                pdf.draw_title()
                top_y = MARGIN + TITLE_H
                pdf.draw_page_labels(top_y, col_w, gap)
                content_top = top_y + LABEL_H

            row_top = content_top + pos_in_page * row_h

            # 施工前(EXIFの向きを補正し、PNGとして保存し直した画像を配置)
            before_path, before_size = _load_image_as_temp_png(before, tmp_files)
            bw, bh = _fit_size(before_size, col_w - 2 * pad, row_h - 2 * pad)
            bx = MARGIN + (col_w - bw) / 2
            by = row_top + (row_h - bh) / 2
            pdf.image(before_path, x=bx, y=by, w=bw, h=bh)

            # 施工後(同上)
            after_path, after_size = _load_image_as_temp_png(after, tmp_files)
            aw, ah = _fit_size(after_size, col_w - 2 * pad, row_h - 2 * pad)
            ax = MARGIN + col_w + gap + (col_w - aw) / 2
            ay = row_top + (row_h - ah) / 2
            pdf.image(after_path, x=ax, y=ay, w=aw, h=ah)

            # 中央の矢印(→)
            arrow_y = row_top + row_h / 2
            arrow_x1 = MARGIN + col_w + 3
            arrow_x2 = MARGIN + col_w + gap - 3
            pdf.set_line_width(0.8)
            pdf.line(arrow_x1, arrow_y, arrow_x2, arrow_y)
            pdf.line(arrow_x2, arrow_y, arrow_x2 - 3, arrow_y - 2)
            pdf.line(arrow_x2, arrow_y, arrow_x2 - 3, arrow_y + 2)

        pdf.output(output_path)
    finally:
        for f in tmp_files:
            try:
                os.remove(f)
            except OSError:
                pass

    return output_path