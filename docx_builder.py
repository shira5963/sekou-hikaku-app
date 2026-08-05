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


def build_docx(property_name, pairs, output_path):
    """
    property_name: str  物件名
    pairs: [(施工前画像パス, 施工後画像パス), ...]
    output_path: 出力する.docxのパス
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
        total = len(pairs)
        for page_start in range(0, total, PAIRS_PER_PAGE):
            # ---- 物件名(タイトル) ----
            title_p = doc.add_paragraph()
            if page_start > 0:
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