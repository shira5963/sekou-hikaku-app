"""
施工前・施工後 比較報告書 Word(.docx)生成モジュール
pdf_builder.py と同じレイアウト仕様(A4縦/1ページ5組)をWordの表で再現する
"""
import os
import tempfile

from docx import Document
from docx.shared import Mm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image, ImageOps

# ---- レイアウト定数 (mm) : pdf_builder.py と揃えている ----
PAGE_W_MM = 210
PAGE_H_MM = 297
MARGIN_MM = 10
TITLE_H_MM = 15
LABEL_H_MM = 8
GAP_MM = 20  # 中央(矢印)列の幅
PAIRS_PER_PAGE = 5

# ---- 表紙に記載する会社情報(固定文言) ----
COMPANY_NAME = "株式会社インクコーポレーション"
COMPANY_ADDRESS = "住所：東京都葛飾区立石8-39-6"
COMPANY_TEL_FAX = "TEL：03-3697-9889　FAX：03-3697-9868"
LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")


def _load_image(img_path):
    """画像を読み込み、EXIFの向き情報をピクセルに反映してから返す
    (スマホ写真が回転して埋め込まれるのを防ぐ)"""
    im = Image.open(img_path)
    im = ImageOps.exif_transpose(im)
    if im.mode not in ("RGB", "L"):
        im = im.convert("RGB")
    return im


def _fit_size(im, max_w, max_h):
    """画像の縦横比を保ったまま、指定した枠(max_w x max_h)に収まるサイズを計算する"""
    w, h = im.size
    ratio = min(max_w / w, max_h / h)
    return w * ratio, h * ratio


def _set_vertical_center(cell):
    """表のセル内の文字・画像を上下中央に配置する"""
    tcPr = cell._tc.get_or_add_tcPr()
    vAlign = OxmlElement("w:vAlign")
    vAlign.set(qn("w:val"), "center")
    tcPr.append(vAlign)


def _set_row_widths(row, widths_mm):
    for cell, w in zip(row.cells, widths_mm):
        cell.width = Mm(w)


def _add_centered_picture(paragraph, img_path, width_mm, height_mm):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.add_run().add_picture(img_path, width=Mm(width_mm), height=Mm(height_mm))


def _add_cover_page(doc, property_name, cover_photo_path, contact_name, tmp_img_files):
    """表紙ページ(1ページ目)を作成する

    レイアウトの考え方(pdf_builder.pyと揃えている)：
    - 写真をA4用紙の上下中央に配置する
    - 「写真報告書」は写真の3行分上、「物件名」はさらにその2行分上
    - ロゴ・会社情報は写真の3行分下
    Wordは絶対座標を指定できないため、段落の space_before/space_after を使って
    行数分の余白を再現し、さらに先頭の物件名の前に計算した余白を入れることで
    ブロック全体がページの上下中央に来るように調整している。
    """
    usable_w = PAGE_W_MM - 2 * MARGIN_MM
    mm_to_pt = 2.83465
    LINE_MM = 8
    line_pt = LINE_MM * mm_to_pt

    # ---- 写真サイズを計算(現在(60%)の120% = 元サイズの72%) ----
    BASE_MAX_H_MM = 150
    max_w = usable_w * 0.72
    max_h = BASE_MAX_H_MM * 0.72
    if cover_photo_path:
        cover_im = _load_image(cover_photo_path)
        cw, ch = _fit_size(cover_im, max_w, max_h)
    else:
        cover_im = None
        cw, ch = 0, max_h * 0.7  # 写真がない場合も位置決めの基準として仮の高さを使う

    # ---- ロゴサイズを事前に計算(全体の高さ見積もりに使う) ----
    if os.path.exists(LOGO_PATH):
        with Image.open(LOGO_PATH) as logo_im:
            lw, lh = _fit_size(logo_im, 110, 20)
    else:
        lw, lh = 0, 0

    # ---- ブロック全体の高さを見積もり、ページ上下中央に来るよう先頭の余白を計算 ----
    title_line_mm = 9
    subtitle_line_mm = 9.5
    company_block_mm = 7 + 5 + 5  # 会社名 + 住所 + TEL/FAX の目安の高さ
    if contact_name.strip():
        company_block_mm += 6  # 担当行がある場合はその分を追加

    total_block_mm = (
        title_line_mm + 2 * LINE_MM + subtitle_line_mm + 3 * LINE_MM
        + ch + 3 * LINE_MM + lh + 4 + company_block_mm
    )
    available_mm = PAGE_H_MM - 2 * MARGIN_MM
    center_offset_mm = max(0, (available_mm - total_block_mm) / 2)

    # ---- 物件名 ----
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(center_offset_mm * mm_to_pt)
    title_run = title_p.add_run(f"物件名　{property_name}")
    title_run.bold = False
    title_run.underline = True
    title_run.font.size = Pt(20)

    # ---- 「写真報告書」(物件名の2行分下) ----
    subtitle_p = doc.add_paragraph()
    subtitle_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle_p.paragraph_format.space_before = Pt(2 * line_pt)
    subtitle_run = subtitle_p.add_run("写真報告書")
    subtitle_run.font.size = Pt(22)

    # ---- 写真(写真報告書の3行分下) ----
    if cover_photo_path:
        tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        cover_im.save(tmp.name)
        tmp.close()
        tmp_img_files.append(tmp.name)
        cover_im.close()

        photo_p = doc.add_paragraph()
        photo_p.paragraph_format.space_before = Pt(3 * line_pt)
        photo_p.paragraph_format.space_after = Pt(3 * line_pt)
        _add_centered_picture(photo_p, tmp.name, cw, ch)
    else:
        # 写真を入れない場合も、レイアウトの間隔だけは揃えておく
        spacer_p = doc.add_paragraph()
        spacer_p.paragraph_format.space_before = Pt(3 * line_pt)
        spacer_p.paragraph_format.space_after = Pt(3 * line_pt + ch * mm_to_pt)

    # ---- ロゴ ----
    if os.path.exists(LOGO_PATH):
        logo_p = doc.add_paragraph()
        logo_p.paragraph_format.space_after = Pt(4)
        _add_centered_picture(logo_p, LOGO_PATH, lw, lh)

    # ---- 会社情報 ----
    name_p = doc.add_paragraph()
    name_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    name_run = name_p.add_run(COMPANY_NAME)
    name_run.font.size = Pt(13)

    if contact_name.strip():
        contact_p = doc.add_paragraph()
        contact_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        contact_run = contact_p.add_run(f"担当：{contact_name.strip()}")
        contact_run.font.size = Pt(11)

    addr_p = doc.add_paragraph()
    addr_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    addr_run = addr_p.add_run(COMPANY_ADDRESS)
    addr_run.font.size = Pt(11)

    tel_p = doc.add_paragraph()
    tel_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tel_run = tel_p.add_run(COMPANY_TEL_FAX)
    tel_run.font.size = Pt(11)


def build_docx(property_name, pairs, output_path, cover_photo_path=None,
               contact_name="", include_cover=True):
    """
    property_name: str  物件名
    pairs: [(施工前画像パス, 施工後画像パス), ...]
    output_path: 出力する.docxのパス
    cover_photo_path: 表紙に載せる写真のパス(Noneなら写真なし)
    contact_name: 担当者名
    include_cover: 表紙ページを作るかどうか
    """
    doc = Document()

    section = doc.sections[0]
    section.page_width = Mm(PAGE_W_MM)
    section.page_height = Mm(PAGE_H_MM)
    section.top_margin = Mm(MARGIN_MM)
    section.bottom_margin = Mm(MARGIN_MM)
    section.left_margin = Mm(MARGIN_MM)
    section.right_margin = Mm(MARGIN_MM)

    usable_w = PAGE_W_MM - 2 * MARGIN_MM
    col_w = (usable_w - GAP_MM) / 2
    row_h = (PAGE_H_MM - 2 * MARGIN_MM - TITLE_H_MM - LABEL_H_MM) / PAIRS_PER_PAGE
    pad = 2
    widths = [col_w, GAP_MM, col_w]

    tmp_img_files = []  # 最後にまとめて削除する一時ファイル

    try:
        if include_cover:
            _add_cover_page(doc, property_name, cover_photo_path, contact_name, tmp_img_files)

        total = len(pairs)
        for page_start in range(0, total, PAIRS_PER_PAGE):
            # ---- 物件名(タイトル) ----
            title_p = doc.add_paragraph()
            if page_start > 0 or include_cover:
                # 表の直後に doc.add_page_break() で改ページ用の段落を追加すると
                # 表がページをほぼ埋めている状態と組み合わさり、Wordが
                # 自動改行と手動改行の両方を行って空白ページが生まれることがある。
                # そのため、次のタイトル段落自体に「ページ区切りしてから開始」の
                # 属性を持たせることで、余分な段落を増やさずに改ページする。
                title_p.paragraph_format.page_break_before = True
            title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            title_run = title_p.add_run(property_name)
            title_run.bold = True
            title_run.font.size = Pt(18)

            # ---- 表の作成(1行目はラベル行) ----
            table = doc.add_table(rows=1, cols=3)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.autofit = False

            header_cells = table.rows[0].cells
            _set_row_widths(table.rows[0], widths)

            hp0 = header_cells[0].paragraphs[0]
            hp0.alignment = WD_ALIGN_PARAGRAPH.CENTER
            hr0 = hp0.add_run("施工前")
            hr0.bold = True
            hr0.font.size = Pt(12)

            hp2 = header_cells[2].paragraphs[0]
            hp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            hr2 = hp2.add_run("施工後")
            hr2.bold = True
            hr2.font.size = Pt(12)

            # ---- 各ペアの行 ----
            chunk = pairs[page_start: page_start + PAIRS_PER_PAGE]
            for before, after in chunk:
                row = table.add_row()
                row.height = Mm(row_h)
                row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
                _set_row_widths(row, widths)
                cells = row.cells

                # 施工前の画像
                before_im = _load_image(before)
                bw, bh = _fit_size(before_im, col_w - 2 * pad, row_h - 2 * pad)
                tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
                before_im.save(tmp.name)
                tmp.close()
                tmp_img_files.append(tmp.name)
                before_im.close()

                p_before = cells[0].paragraphs[0]
                p_before.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_before.add_run().add_picture(tmp.name, width=Mm(bw), height=Mm(bh))
                _set_vertical_center(cells[0])

                # 中央の矢印
                p_arrow = cells[1].paragraphs[0]
                p_arrow.alignment = WD_ALIGN_PARAGRAPH.CENTER
                arrow_run = p_arrow.add_run("→")
                arrow_run.bold = True
                arrow_run.font.size = Pt(28)
                _set_vertical_center(cells[1])

                # 施工後の画像
                after_im = _load_image(after)
                aw, ah = _fit_size(after_im, col_w - 2 * pad, row_h - 2 * pad)
                tmp2 = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
                after_im.save(tmp2.name)
                tmp2.close()
                tmp_img_files.append(tmp2.name)
                after_im.close()

                p_after = cells[2].paragraphs[0]
                p_after.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_after.add_run().add_picture(tmp2.name, width=Mm(aw), height=Mm(ah))
                _set_vertical_center(cells[2])

        doc.save(output_path)
    finally:
        for f in tmp_img_files:
            try:
                os.remove(f)
            except OSError:
                pass

    return output_path
