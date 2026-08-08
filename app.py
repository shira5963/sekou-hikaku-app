"""
施工前&施工後 比較報告書 作成アプリ
Streamlitで動くWeb画面。写真をアップロードすると自動でペアリングし、
表紙付きのA4 PDF/Word比較報告書を生成する。
"""
import streamlit as st
import os
import re
import base64
import tempfile
from pdf_builder import build_pdf
from docx_builder import build_docx
from usage_logger import log_usage, read_log
from layout_utils import (
    BEFORE_AFTER_MAX_CHARS,
    EXTRA_MAX_CHARS,
    EXTRA_PHOTOS_PER_ROW,
    EXTRA_PHOTOS_PER_PAGE,
)

APP_DIR = os.path.dirname(__file__)
LOGO_PATH = os.path.join(APP_DIR, "assets", "logo.png")

st.set_page_config(
    page_title="施工前&施工後比較報告書",
    page_icon="🏗️",
    layout="wide",
)

# ----------------------------------------------------------------------
# デザイン(カスタムCSS)
# 会社ロゴの深緑色をブランドカラーとして採用し、カード型のレイアウトに変更
# ----------------------------------------------------------------------
BRAND_DARK = "#003222"
BRAND_MAIN = "#0B4A34"
BRAND_LIGHT = "#EAF1EC"
BRAND_ACCENT = "#C9A24B"

st.markdown(
    f"""<style>
.stApp {{
    background-color: #F7F8F6;
    font-family: 'Yu Gothic', 'Hiragino Sans', 'Meiryo', sans-serif;
}}
/* 上部ヘッダーバナー */
.app-header {{
    background: linear-gradient(135deg, {BRAND_DARK} 0%, {BRAND_MAIN} 100%);
    padding: 28px 36px;
    border-radius: 16px;
    margin-bottom: 28px;
    display: flex;
    align-items: center;
    flex-wrap: wrap;
    gap: 20px;
    box-shadow: 0 6px 18px rgba(0,50,34,0.18);
}}
.app-header img {{
    height: auto;
    width: 100%;
    max-width: 180px;
    filter: brightness(0) invert(1);
    flex-shrink: 0;
}}
.app-header .title-block {{
    flex: 1 1 240px;
    min-width: 0;
}}
.app-header .title-block h1 {{
    color: #FFFFFF;
    font-size: 22px;
    margin: 0;
    font-weight: 700;
    letter-spacing: 0.02em;
    line-height: 1.4;
}}
.app-header .title-block p {{
    color: {BRAND_LIGHT};
    margin: 4px 0 0 0;
    font-size: 13px;
    line-height: 1.5;
}}
/* スマホ幅では、ロゴとタイトルを縦に並べて中央揃えにする */
@media (max-width: 640px) {{
    .app-header {{
        flex-direction: column;
        align-items: center;
        text-align: center;
        padding: 22px 18px;
        gap: 14px;
    }}
    .app-header img {{
        max-width: 140px;
    }}
    .app-header .title-block h1 {{
        font-size: 18px;
    }}
    .app-header .title-block p {{
        font-size: 12px;
    }}
}}
/* ステップ見出し */
.step-label {{
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background-color: {BRAND_DARK};
    color: white;
    padding: 4px 14px;
    border-radius: 999px;
    font-size: 13px;
    font-weight: 600;
    margin-bottom: 10px;
}}
/* カード(枠付きコンテナ)のスタイル */
div[data-testid="stVerticalBlockBorderWrapper"] {{
    border-radius: 14px !important;
    border: 1px solid #E2E7E3 !important;
    box-shadow: 0 2px 10px rgba(0,0,0,0.04);
    background-color: #FFFFFF;
}}
/* ボタン */
.stButton > button, div[data-testid="stDownloadButton"] > button {{
    border-radius: 8px;
    font-weight: 600;
    border: 1px solid {BRAND_DARK};
}}
.stButton > button[kind="primary"] {{
    background-color: {BRAND_DARK};
    border: 1px solid {BRAND_DARK};
}}
.stButton > button[kind="primary"]:hover {{
    background-color: {BRAND_MAIN};
    border: 1px solid {BRAND_MAIN};
}}
div[data-testid="stDownloadButton"] > button {{
    background-color: {BRAND_ACCENT};
    color: #FFFFFF;
    border: 1px solid {BRAND_ACCENT};
}}
/* テキスト入力・ラジオなどのラベル */
label p {{
    font-weight: 600 !important;
    color: #23352C !important;
}}
</style>""",
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------
# ヘッダーバナー(ロゴ + タイトル)
# ----------------------------------------------------------------------
if os.path.exists(LOGO_PATH):
    with open(LOGO_PATH, "rb") as f:
        logo_b64 = base64.b64encode(f.read()).decode()
    logo_html = f'<img src="data:image/png;base64,{logo_b64}" alt="logo">'
else:
    logo_html = ""

st.markdown(
    f"""<div class="app-header">{logo_html}<div class="title-block">
<h1>施工前&施工後 比較報告書 作成アプリ</h1>
<p>物件写真をアップロードするだけで、表紙付きの比較報告書(PDF / Word)を自動作成します</p>
</div></div>""",
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------
# STEP 1：物件名
# ----------------------------------------------------------------------
st.markdown('<div class="step-label">STEP 1　物件名</div>', unsafe_allow_html=True)
with st.container(border=True):
    property_name = st.text_input(
        "物件名",
        "",
        placeholder="例：◯◯レジデンス101号室",
        label_visibility="collapsed",
    )

st.write("")

# ----------------------------------------------------------------------
# STEP 2：表紙の設定
# ----------------------------------------------------------------------
st.markdown('<div class="step-label">STEP 2　表紙の設定</div>', unsafe_allow_html=True)
with st.container(border=True):
    cover_choice = st.radio(
        "表紙の写真",
        ["表紙あり（写真を追加する）", "表紙なし（写真部分を空白にする）"],
        horizontal=True,
    )
    cover_photo_file = None
    if cover_choice.startswith("表紙あり"):
        cover_photo_file = st.file_uploader(
            "表紙中央に配置する写真をアップロード",
            # type=[...]で拡張子を絞り込むと、Android端末でアップロード時に
            # 「フォト」アプリが選択肢から消えてしまうことがあるため、
            # あえて絞り込まず、アップロード後にPython側でチェックする
            accept_multiple_files=False,
            key="cover_photo",
        )

    contact_name = st.text_input(
        "担当者名（表紙下部に「担当：〇〇」として記載されます／空欄可）",
        "",
        placeholder="例：山田太郎",
    )

st.write("")

# ----------------------------------------------------------------------
# STEP 3：施工前・施工後の写真
# ----------------------------------------------------------------------
st.markdown('<div class="step-label">STEP 3　施工前・施工後の写真</div>', unsafe_allow_html=True)
with st.container(border=True):
    col1, col2 = st.columns(2)
    with col1:
        before_files = st.file_uploader(
            "施工前の写真をアップロード（複数選択可）",
            accept_multiple_files=True,
            key="before",
        )
    with col2:
        after_files = st.file_uploader(
            "施工後の写真をアップロード（複数選択可）",
            accept_multiple_files=True,
            key="after",
        )


ALLOWED_EXT = (".jpg", ".jpeg", ".png")


def filter_image_files(files):
    """拡張子が画像(jpg/jpeg/png)のものだけを残し、それ以外は除外する。
    (アップロード欄側で拡張子を絞り込まなくなったため、ここでチェックする)"""
    if not files:
        return [], []
    valid = [f for f in files if f.name.lower().endswith(ALLOWED_EXT)]
    invalid = [f.name for f in files if not f.name.lower().endswith(ALLOWED_EXT)]
    return valid, invalid


def natural_key(name):
    """ファイル名を自然順(数字は数値として)でソートするためのキー。
    例: 2.jpg が 10.jpg より前に来るようにする"""
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", name)]


st.write("")

cover_photo_file, cover_invalid = filter_image_files(
    [cover_photo_file] if cover_photo_file is not None else []
)
cover_photo_file = cover_photo_file[0] if cover_photo_file else None
if cover_invalid:
    st.error(f"表紙の写真は画像ファイル(jpg/jpeg/png)のみアップロードできます：{cover_invalid[0]}")

before_files, before_invalid = filter_image_files(before_files)
after_files, after_invalid = filter_image_files(after_files)
if before_invalid or after_invalid:
    st.error(
        "画像ファイル(jpg/jpeg/png)以外はアップロードできません。"
        f"該当ファイル：{', '.join(before_invalid + after_invalid)}"
    )

# ----------------------------------------------------------------------
# STEP 4：施工中画像・その他
# ----------------------------------------------------------------------
st.markdown('<div class="step-label">STEP 4　施工中画像・その他</div>', unsafe_allow_html=True)
with st.container(border=True):
    extra_title = st.text_input(
        "このページの題目",
        "施工中画像",
        placeholder="例：施工中画像",
    )
    extra_files_raw = st.file_uploader(
        "施工中の写真やその他の写真をアップロード（複数選択可／なくてもOK）",
        accept_multiple_files=True,
        key="extra_photos",
    )
    extra_files, extra_invalid = filter_image_files(extra_files_raw)
    if extra_invalid:
        st.error(
            "画像ファイル(jpg/jpeg/png)以外はアップロードできません。"
            f"該当ファイル：{', '.join(extra_invalid)}"
        )

    extra_sorted = sorted(extra_files, key=lambda f: natural_key(f.name))

    if extra_sorted:
        extra_sig = tuple(f.name for f in extra_sorted)
        if st.session_state.get("extra_sig") != extra_sig:
            # 前回と異なる写真の組み合わせがアップロードされた場合、
            # 古いコメント入力の状態が残らないよう一度クリアする
            for key in list(st.session_state.keys()):
                if key.startswith("extra_comment_"):
                    del st.session_state[key]
            st.session_state["extra_sig"] = extra_sig

        st.caption(
            f"1ページに横{EXTRA_PHOTOS_PER_ROW}枚 × 縦5段 = 最大{EXTRA_PHOTOS_PER_PAGE}枚まで配置されます"
            f"（{EXTRA_PHOTOS_PER_PAGE}枚を超えると自動で次のページに続きます）。"
            f"各写真の下のコメントは{EXTRA_MAX_CHARS}文字まで入力できます。"
        )
        n_extra = len(extra_sorted)
        for row_start in range(0, n_extra, EXTRA_PHOTOS_PER_ROW):
            row_files = extra_sorted[row_start: row_start + EXTRA_PHOTOS_PER_ROW]
            cols = st.columns(EXTRA_PHOTOS_PER_ROW)
            for col, f in zip(cols, row_files):
                idx = extra_sorted.index(f)
                with col:
                    st.image(f, use_container_width=True)
                    st.text_input(
                        "コメント",
                        key=f"extra_comment_{idx}",
                        max_chars=EXTRA_MAX_CHARS,
                        label_visibility="collapsed",
                        placeholder="コメント",
                    )

st.write("")

if before_files and after_files:
    before_sorted = sorted(before_files, key=lambda f: natural_key(f.name))
    after_sorted = sorted(after_files, key=lambda f: natural_key(f.name))
    after_names = [f.name for f in after_sorted]

    n = len(before_sorted)
    if len(before_sorted) != len(after_sorted):
        st.warning(
            f"施工前({len(before_sorted)}枚)と施工後({len(after_sorted)}枚)の枚数が一致していません。"
            "自動では先頭から順にペアにしていますが、下の「変更」欄で正しい組み合わせに直せます。"
        )

    # ---- ペアリング状態をセッションに保持 ----
    # アップロードされたファイルの組み合わせが変わったら、状態をリセットする
    sig = (tuple(f.name for f in before_sorted), tuple(after_names))
    if st.session_state.get("pairing_sig") != sig:
        # 前回と異なる写真の組み合わせがアップロードされた場合、
        # 古いプルダウンの選択状態が残らないよう一度クリアする
        for key in list(st.session_state.keys()):
            if key.startswith("after_select_"):
                del st.session_state[key]
        st.session_state["pairing_sig"] = sig
        st.session_state["pair_order"] = list(range(n))  # 表示順(施工前のインデックス)
        st.session_state["after_choice"] = {
            i: min(i, len(after_sorted) - 1) for i in range(n)
        }  # 施工前インデックス -> 施工後インデックス

    order = st.session_state["pair_order"]
    after_choice = st.session_state["after_choice"]

    # 重複チェック(同じ施工後写真が複数のペアで使われていないか)
    from collections import Counter
    usage_count = Counter(after_choice.values())

    st.markdown('<div class="step-label">STEP 5　ペアリング結果の確認・修正</div>', unsafe_allow_html=True)
    with st.expander(f"{n}組のペアを確認する（同じ場所の写真になっているか確認してください）", expanded=True):
        st.caption(
            "違う場所の写真がペアになっている場合は「施工後の写真」のプルダウンから正しい写真を選び直せます。"
            "表示順（報告書のページ順）を変えたい場合は ▲▼ ボタンで入れ替えられます。"
            f"各写真の下のコメントは{BEFORE_AFTER_MAX_CHARS}文字まで入力できます。"
        )
        comments_by_idx = {}
        for display_pos, before_idx in enumerate(order):
            b = before_sorted[before_idx]
            current_after_idx = after_choice[before_idx]
            duplicated = usage_count[current_after_idx] > 1

            row_box = st.container(border=True)
            with row_box:
                c1, c2, c3, c4 = st.columns([4, 3, 4, 1])
                with c1:
                    st.image(b, caption=f"{display_pos + 1}. 施工前: {b.name}", use_container_width=True)
                    before_comment = st.text_input(
                        "コメント（施工前）",
                        key=f"before_comment_{before_idx}",
                        max_chars=BEFORE_AFTER_MAX_CHARS,
                        label_visibility="collapsed",
                        placeholder="コメント（施工前）",
                    )
                with c2:
                    st.markdown(
                        f"<div style='text-align:center;margin-top:70px;font-size:26px;color:{BRAND_DARK};'>→</div>",
                        unsafe_allow_html=True,
                    )
                    selected_name = st.selectbox(
                        "施工後の写真",
                        options=after_names,
                        index=current_after_idx,
                        key=f"after_select_{before_idx}",
                    )
                    new_after_idx = after_names.index(selected_name)
                    if new_after_idx != after_choice[before_idx]:
                        after_choice[before_idx] = new_after_idx
                        st.rerun()
                    if duplicated:
                        st.caption("⚠️ 他のペアと同じ施工後写真が選ばれています")
                with c3:
                    a = after_sorted[current_after_idx]
                    st.image(a, caption=f"{display_pos + 1}. 施工後: {a.name}", use_container_width=True)
                    after_comment = st.text_input(
                        "コメント（施工後）",
                        key=f"after_comment_{before_idx}",
                        max_chars=BEFORE_AFTER_MAX_CHARS,
                        label_visibility="collapsed",
                        placeholder="コメント（施工後）",
                    )
                with c4:
                    st.write("")
                    if st.button("▲", key=f"up_{before_idx}", disabled=(display_pos == 0), use_container_width=True):
                        new_order = order.copy()
                        new_order[display_pos - 1], new_order[display_pos] = (
                            new_order[display_pos],
                            new_order[display_pos - 1],
                        )
                        st.session_state["pair_order"] = new_order
                        st.rerun()
                    if st.button("▼", key=f"down_{before_idx}", disabled=(display_pos == len(order) - 1), use_container_width=True):
                        new_order = order.copy()
                        new_order[display_pos + 1], new_order[display_pos] = (
                            new_order[display_pos],
                            new_order[display_pos + 1],
                        )
                        st.session_state["pair_order"] = new_order
                        st.rerun()

                comments_by_idx[before_idx] = (before_comment, after_comment)

    pairs_preview = [
        (before_sorted[idx], after_sorted[after_choice[idx]], *comments_by_idx[idx])
        for idx in order
    ]

    st.write("")
    st.markdown('<div class="step-label">STEP 6　報告書の作成</div>', unsafe_allow_html=True)
    with st.container(border=True):
        btn_col1, btn_col2 = st.columns(2)
        with btn_col1:
            make_pdf = st.button("📄 PDFで作成", type="primary", use_container_width=True)
        with btn_col2:
            make_docx = st.button("📝 Wordで作成", use_container_width=True)

        if make_pdf or make_docx:
            if not property_name.strip():
                st.error("STEP 1で物件名を入力してください。")
            else:
                with tempfile.TemporaryDirectory() as tmpdir:
                    # アップロードされた写真を一時フォルダに保存(PDF/Word共通で使う)
                    pair_paths = []
                    for i, (b, a, b_comment, a_comment) in enumerate(pairs_preview):
                        b_path = os.path.join(tmpdir, f"before_{i}_{b.name}")
                        a_path = os.path.join(tmpdir, f"after_{i}_{a.name}")
                        with open(b_path, "wb") as f:
                            f.write(b.getbuffer())
                        with open(a_path, "wb") as f:
                            f.write(a.getbuffer())
                        pair_paths.append((b_path, a_path, b_comment, a_comment))

                    # 表紙写真の保存(選択されている場合のみ)
                    cover_photo_path = None
                    if cover_photo_file is not None:
                        cover_photo_path = os.path.join(tmpdir, f"cover_{cover_photo_file.name}")
                        with open(cover_photo_path, "wb") as f:
                            f.write(cover_photo_file.getbuffer())

                    # 施工中画像・その他の写真の保存(アップロードされている場合のみ)
                    extra_photo_items = []
                    for i, f in enumerate(extra_sorted):
                        e_path = os.path.join(tmpdir, f"extra_{i}_{f.name}")
                        with open(e_path, "wb") as out:
                            out.write(f.getbuffer())
                        comment = st.session_state.get(f"extra_comment_{i}", "")
                        extra_photo_items.append((e_path, comment))

                    if make_pdf:
                        with st.spinner("PDFを生成しています..."):
                            pdf_path = os.path.join(tmpdir, "report.pdf")
                            build_pdf(
                                property_name,
                                pair_paths,
                                pdf_path,
                                cover_photo_path=cover_photo_path,
                                contact_name=contact_name,
                                include_cover=True,
                                extra_title=extra_title,
                                extra_photos=extra_photo_items,
                            )
                            with open(pdf_path, "rb") as f:
                                pdf_bytes = f.read()
                        log_usage(property_name, contact_name, ["PDF"])
                        st.success("PDF報告書を作成しました。")
                        st.download_button(
                            "⬇ PDFをダウンロード",
                            data=pdf_bytes,
                            file_name=f"{property_name}_比較報告書.pdf",
                            mime="application/pdf",
                        )

                    if make_docx:
                        with st.spinner("Wordファイルを生成しています..."):
                            docx_path = os.path.join(tmpdir, "report.docx")
                            build_docx(
                                property_name,
                                pair_paths,
                                docx_path,
                                cover_photo_path=cover_photo_path,
                                contact_name=contact_name,
                                include_cover=True,
                                extra_title=extra_title,
                                extra_photos=extra_photo_items,
                            )
                            with open(docx_path, "rb") as f:
                                docx_bytes = f.read()
                        log_usage(property_name, contact_name, ["Word"])
                        st.success("Word報告書を作成しました。")
                        st.download_button(
                            "⬇ Wordファイルをダウンロード",
                            data=docx_bytes,
                            file_name=f"{property_name}_比較報告書.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )
else:
    st.info("STEP 3で、施工前・施工後それぞれの写真をアップロードすると、ペアリング結果と作成ボタンが表示されます。")

# ----------------------------------------------------------------------
# 利用ログの確認(管理者向け・パスワード保護)
# ----------------------------------------------------------------------
ADMIN_PASSWORD = "5963"

st.write("")
with st.expander("📋 利用ログを見る（管理者向け）"):
    if not st.session_state.get("log_authenticated", False):
        pw = st.text_input(
            "管理者パスワード（4桁）を入力してください",
            type="password",
            max_chars=4,
            key="log_password_input",
        )
        if pw:
            if pw == ADMIN_PASSWORD:
                st.session_state["log_authenticated"] = True
                st.rerun()
            else:
                st.error("パスワードが違います。")
    else:
        st.caption(
            "報告書を作成するたびに、日時・物件名・担当者名・作成形式・アクセス元IPアドレスが記録されます。"
            "IPアドレスの取得はStreamlitの内部機能を利用しているため、環境によっては「不明」と表示される場合があります。"
        )
        header, rows = read_log()
        if rows:
            import pandas as pd
            log_df = pd.DataFrame(rows, columns=header)
            st.dataframe(log_df, hide_index=True, use_container_width=True)
            with open(os.path.join(os.path.dirname(__file__), "logs", "usage_log.csv"), "rb") as f:
                st.download_button(
                    "⬇ ログをCSVでダウンロード",
                    data=f.read(),
                    file_name="usage_log.csv",
                    mime="text/csv",
                )
        else:
            st.info("まだ利用ログはありません。")

        st.write("")
        if st.button("🔒 ロックする", key="log_lock_button"):
            st.session_state["log_authenticated"] = False
            st.rerun()
