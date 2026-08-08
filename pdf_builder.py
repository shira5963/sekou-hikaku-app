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

# ---- 記入欄(コメント欄)の設定 ----
COMMENT_BOX_H = 11  # mm 記入欄の高さ
COMMENT_GAP = 1.5  # mm 写真と記入欄の間隔
COMMENT_FONT_SIZE = 10.5  # 記入欄ラベルのフォントサイズ

# ---- 画像圧縮の設定 ----
# 写真は実際にA4上へ配置されるサイズを基準に、印刷でも見た目がほぼ変わらない
# 解像度(200dpi)まで縮小してからJPEGで保存し、ファイルサイズを大幅に削減する
IMAGE_DPI = 200
JPEG_QUALITY = 85

# ---- 日本語フォント設定 ----
# fpdf2は標準では日本語を描画できないため、UTF-8対応のTTFフォントが必要です。
# fonts/ フォルダに IPAexゴシック等のフォントファイルを配置してください。
FONT_DIR = os.path.join(os.path.dirname(__file__), "fonts")
FONT_PATH = os.path.join(FONT_DIR, "ipaexg.ttf")
FONT_NAME = "JPFont"

# ---- 表紙に記載する会社情報(固定文言) ----
# 社名・住所・電話番号などを変更したい場合はここを書き換えてください
COMPANY_NAME = "株式会社インクコーポレーション"
COMPANY_ADDRESS = "住所：東京都葛飾区立石8-39-6"
COMPANY_TEL_FAX = "TEL：03-3697-9889　FAX：03-3697-9868"
LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")


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


def _fit_size(size, max_w, max_h):
    """画像の縦横比を保ったまま、指定した枠(max_w x max_h)に収まるサイズを計算する"""
    w, h = size
    ratio = min(max_w / w, max_h / h)
    return w * ratio, h * ratio


def _prepare_image_for_pdf(img_path, max_w_mm, max_h_mm, tmp_files,
                            dpi=IMAGE_DPI, quality=JPEG_QUALITY):
    """画像を読み込み、向き補正(EXIF)をした上で、実際にPDF上へ配置される
    サイズ(mm)を基準に、印刷でも十分きれいに見える解像度(dpi)まで縮小し、
    JPEGとして一時ファイルに保存する。

    スマホ写真をそのまま埋め込むと1枚数MBになりPDFが非常に重くなるため、
    「表示サイズ×dpi」で必要十分なピクセル数まで落とすことでファイルサイズを
    大幅に削減しつつ、印刷時の見た目はほぼ変わらないようにしている。

    戻り値: (一時ファイルパス, (表示幅mm, 表示高さmm))
    """
    im = Image.open(img_path)
    im = ImageOps.exif_transpose(im)  # EXIFの回転情報を実際のピクセルに反映
    if im.mode not in ("RGB", "L"):
        im = im.convert("RGB")

    disp_w, disp_h = _fit_size(im.size, max_w_mm, max_h_mm)

    # 表示サイズ(mm)を、指定dpiで必要なピクセル数に変換
    target_px_w = max(1, round(disp_w / 25.4 * dpi))
    target_px_h = max(1, round(disp_h / 25.4 * dpi))

    # 元画像がそれより大きい場合のみ縮小する(小さい画像を無理に拡大しない)
    if im.size[0] > target_px_w or im.size[1] > target_px_h:
        im = im.resize((target_px_w, target_px_h), Image.LANCZOS)

    tmp = tempfile.NamedTemporaryFile(suffix=".jpg", delete=False)
    im.save(tmp.name, format="JPEG", quality=quality, optimize=True)
    tmp.close()
    tmp_files.append(tmp.name)
    im.close()

    return tmp.name, (disp_w, disp_h)


def _draw_comment_box(pdf, x, y, w, h):
    """写真下の記入欄(枠線のみの空欄)を描画する"""
    pdf.set_draw_color(180, 180, 180)
    pdf.set_line_width(0.2)
    pdf.rect(x, y, w, h)
    pdf.set_font(FONT_NAME, "", COMMENT_FONT_SIZE)
    pdf.set_text_color(150, 150, 150)
    pdf.set_xy(x + 2, y + 1)
    pdf.cell(w - 4, 5, "記入欄：", align="L")
    # 色を元に戻しておく(以降の描画に影響しないように)
    pdf.set_draw_color(0, 0, 0)
    pdf.set_text_color(0, 0, 0)


def draw_cover_page(pdf, property_name, cover_photo_path, contact_name, tmp_files):
    """表紙ページを1ページ描画する
    property_name: 物件名
    cover_photo_path: 表紙中央に配置する写真のパス(Noneの場合は写真なし)
    contact_name: 担当者名(空文字なら「担当：」の行自体を省略)

    レイアウトの考え方：
    - 写真をA4用紙の上下中央に配置する
    - 「写真報告書」は写真の3行分上、「物件名」はさらにその2行分上
    - ロゴ・会社情報は写真の3行分下
    """
    pdf.add_page()

    LINE_MM = 8  # 「1行分」の目安の高さ

    # ---- 写真サイズを計算(現在(60%)の120% = 元サイズの72%) ----
    max_w = (PAGE_W - 2 * MARGIN) * 0.72
    max_h = 150 * 0.72  # 150mmは基準となる写真エリアの高さ
    if cover_photo_path:
        photo_path, (pw, ph) = _prepare_image_for_pdf(cover_photo_path, max_w, max_h, tmp_files)
    else:
        photo_path = None
        pw, ph = 0, max_h * 0.7  # 写真がない場合も位置決めの基準として仮の高さを使う

    # ---- 写真をA4用紙の上下中央に配置 ----
    photo_top = PAGE_H / 2 - ph / 2
    photo_bottom = PAGE_H / 2 + ph / 2
    if cover_photo_path:
        photo_x = (PAGE_W - pw) / 2
        pdf.image(photo_path, x=photo_x, y=photo_top, w=pw, h=ph)

    # ---- 「写真報告書」(写真の3行分上) ----
    subtitle_h = 14
    subtitle_top = photo_top - 3 * LINE_MM - subtitle_h
    pdf.set_xy(0, subtitle_top)
    pdf.set_font(FONT_NAME, "", 22)
    pdf.cell(PAGE_W, subtitle_h, "写真報告書", align="C")

    # ---- 「物件名」(写真報告書のさらに2行分上) ----
    title_h = 12
    title_top = subtitle_top - 2 * LINE_MM - title_h
    pdf.set_xy(0, title_top)
    pdf.set_font(FONT_NAME, "", 20)
    title_text = f"物件名　{property_name}"
    pdf.cell(PAGE_W, title_h, title_text, align="C")
    # 物件名の下に下線を引く(テンプレートに合わせた装飾)
    text_w = pdf.get_string_width(title_text)
    line_y = title_top + title_h + 1
    pdf.set_line_width(0.3)
    pdf.line((PAGE_W - text_w) / 2, line_y, (PAGE_W + text_w) / 2, line_y)

    # ---- ロゴ(写真の3行分下) ----
    logo_y = photo_bottom + 3 * LINE_MM
    if os.path.exists(LOGO_PATH):
        with Image.open(LOGO_PATH) as logo_im:
            logo_size = logo_im.size
        logo_w, logo_h = _fit_size(logo_size, 110, 20)
        logo_x = (PAGE_W - logo_w) / 2
        pdf.image(LOGO_PATH, x=logo_x, y=logo_y, w=logo_w, h=logo_h)
        logo_y += logo_h + 4

    # ---- 会社情報 ----
    pdf.set_xy(0, logo_y)
    pdf.set_font(FONT_NAME, "", 13)
    pdf.cell(PAGE_W, 7, COMPANY_NAME, align="C")

    next_y = logo_y + 7
    if contact_name.strip():
        pdf.set_xy(0, next_y)
        pdf.set_font(FONT_NAME, "", 11)
        pdf.cell(PAGE_W, 6, f"担当：{contact_name.strip()}", align="C")
        next_y += 6

    pdf.set_xy(0, next_y)
    pdf.set_font(FONT_NAME, "", 11)
    pdf.cell(PAGE_W, 6, COMPANY_ADDRESS, align="C")

    pdf.set_xy(0, next_y + 6)
    pdf.cell(PAGE_W, 6, COMPANY_TEL_FAX, align="C")


def build_pdf(property_name, pairs, output_path, cover_photo_path=None,
              contact_name="", include_cover=True):
    """
    property_name: str  物件名
    pairs: [(施工前画像パス, 施工後画像パス), ...]  最大枚数の制限なし(自動改ページ)
    output_path: 出力するPDFのパス
    cover_photo_path: 表紙に載せる写真のパス(Noneなら写真なし)
    contact_name: 担当者名
    include_cover: 表紙ページを作るかどうか
    """
    pdf = ReportPDF(property_name)

    usable_w = PAGE_W - 2 * MARGIN
    gap = 20  # mm 矢印を描くための中央スペース
    col_w = (usable_w - gap) / 2
    row_h = (PAGE_H - 2 * MARGIN - TITLE_H - LABEL_H) / PAIRS_PER_PAGE
    pad = 2  # mm 画像とセル枠の間の余白

    # 各行の高さのうち、下部を記入欄に充て、残りを写真エリアとする
    photo_zone_h = row_h - COMMENT_GAP - COMMENT_BOX_H

    tmp_files = []  # 最後にまとめて削除する一時ファイル

    try:
        if include_cover:
            draw_cover_page(pdf, property_name, cover_photo_path, contact_name, tmp_files)

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

            # 施工前(向き補正・圧縮した画像を写真エリア内に配置)
            before_path, (bw, bh) = _prepare_image_for_pdf(
                before, col_w - 2 * pad, photo_zone_h - 2 * pad, tmp_files
            )
            bx = MARGIN + (col_w - bw) / 2
            by = row_top + (photo_zone_h - bh) / 2
            pdf.image(before_path, x=bx, y=by, w=bw, h=bh)

            # 施工後(同上)
            after_path, (aw, ah) = _prepare_image_for_pdf(
                after, col_w - 2 * pad, photo_zone_h - 2 * pad, tmp_files
            )
            ax = MARGIN + col_w + gap + (col_w - aw) / 2
            ay = row_top + (photo_zone_h - ah) / 2
            pdf.image(after_path, x=ax, y=ay, w=aw, h=ah)

            # 中央の矢印(→) : 写真エリアの高さを基準に中央に配置
            arrow_y = row_top + photo_zone_h / 2
            arrow_x1 = MARGIN + col_w + 3
            arrow_x2 = MARGIN + col_w + gap - 3
            pdf.set_line_width(0.8)
            pdf.line(arrow_x1, arrow_y, arrow_x2, arrow_y)
            pdf.line(arrow_x2, arrow_y, arrow_x2 - 3, arrow_y - 2)
            pdf.line(arrow_x2, arrow_y, arrow_x2 - 3, arrow_y + 2)

            # 記入欄(施工前・施工後それぞれの写真の下)
            box_y = row_top + photo_zone_h + COMMENT_GAP
            _draw_comment_box(pdf, MARGIN, box_y, col_w, COMMENT_BOX_H)
            _draw_comment_box(pdf, MARGIN + col_w + gap, box_y, col_w, COMMENT_BOX_H)

        pdf.output(output_path)
    finally:
        for f in tmp_files:
            try:
                os.remove(f)
            except OSError:
                pass

    return output_path
