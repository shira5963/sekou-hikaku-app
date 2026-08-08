"""
アプリ全体(app.py / pdf_builder.py / docx_builder.py)で共通して使う
レイアウト定数・文字数制限の計算をまとめたモジュール。
ここを1箇所直せば、アプリ画面とPDF/Wordの見た目がズレないようにしている。
"""

# ---- A4ページの基本サイズ (mm) ----
PAGE_W_MM = 210
PAGE_H_MM = 297
MARGIN_MM = 10
USABLE_W_MM = PAGE_W_MM - 2 * MARGIN_MM  # 190mm

# ---- 記入欄(コメント欄)の設定 ----
COMMENT_BOX_H_MM = 11
COMMENT_GAP_MM = 1.5
COMMENT_FONT_PT = 10.5

# 全角文字1文字あたりの目安の幅(mm)。10.5ptの全角文字はおおよそ正方形になるため、
# ポイント→mm換算(×0.3528)に、罫線などとの余裕を見て5%上乗せしている。
_CHAR_WIDTH_MM = COMMENT_FONT_PT * 0.3528 * 1.05


def max_chars_for_width(width_mm, padding_mm=4):
    """指定した幅(mm)のコメント欄に、10.5ptの文字が1行に収まる最大文字数の目安を返す"""
    usable = max(width_mm - padding_mm, 10)
    return max(1, int(usable / _CHAR_WIDTH_MM))


# ---- 施工前・施工後 比較ページのレイアウト ----
GAP_MM = 20  # 施工前後の間、矢印を描くスペース
COL_W_MM = (USABLE_W_MM - GAP_MM) / 2  # 85mm
BEFORE_AFTER_MAX_CHARS = max_chars_for_width(COL_W_MM)

# ---- 施工中画像・その他ページのレイアウト ----
EXTRA_PHOTOS_PER_ROW = 4
EXTRA_GAP_MM = 6  # 写真同士の間隔
EXTRA_COL_W_MM = (USABLE_W_MM - (EXTRA_PHOTOS_PER_ROW - 1) * EXTRA_GAP_MM) / EXTRA_PHOTOS_PER_ROW  # 43mm
EXTRA_MAX_CHARS = max_chars_for_width(EXTRA_COL_W_MM)

EXTRA_ROWS_PER_PAGE = 5  # 縦に5段 × 横4枚 = A4 1枚あたり最大20枚
EXTRA_PHOTOS_PER_PAGE = EXTRA_PHOTOS_PER_ROW * EXTRA_ROWS_PER_PAGE
