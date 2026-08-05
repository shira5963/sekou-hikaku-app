"""
施工前&施工後 比較報告書 作成アプリ
Streamlitで動くWeb画面。写真をアップロードすると自動でペアリングし、
A4のPDF比較報告書を生成する。
"""
import streamlit as st
import os
import re
import tempfile
from pdf_builder import build_pdf
from docx_builder import build_docx

st.set_page_config(page_title="施工前&施工後比較報告書", layout="wide")
st.title("施工前&施工後 比較報告書 作成アプリ")

property_name = st.text_input("物件名を入力してください", "")

col1, col2 = st.columns(2)
with col1:
    before_files = st.file_uploader(
        "施工前の写真をアップロード（複数選択可）",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key="before",
    )
with col2:
    after_files = st.file_uploader(
        "施工後の写真をアップロード（複数選択可）",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key="after",
    )


def natural_key(name):
    """ファイル名を自然順(数字は数値として)でソートするためのキー。
    例: 2.jpg が 10.jpg より前に来るようにする"""
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", name)]


if before_files and after_files:
    before_sorted = sorted(before_files, key=lambda f: natural_key(f.name))
    after_sorted = sorted(after_files, key=lambda f: natural_key(f.name))

    n = min(len(before_sorted), len(after_sorted))
    if len(before_sorted) != len(after_sorted):
        st.warning(
            f"施工前({len(before_sorted)}枚)と施工後({len(after_sorted)}枚)の枚数が一致していません。"
            f"ファイル名順に先頭から{n}組をペアとして扱います。"
        )

    st.subheader("ペアリング結果の確認")
    st.caption(
        "同じ場所の写真同士がペアになっているか確認してください。"
        "ズレている場合はファイル名の連番を揃えて再アップロードしてください。"
    )

    pairs_preview = list(zip(before_sorted[:n], after_sorted[:n]))
    for i, (b, a) in enumerate(pairs_preview, start=1):
        c1, c2, c3 = st.columns([4, 1, 4])
        with c1:
            st.image(b, caption=f"{i}. 施工前: {b.name}", use_container_width=True)
        with c2:
            st.markdown(
                "<h2 style='text-align:center;margin-top:60px;'>→</h2>",
                unsafe_allow_html=True,
            )
        with c3:
            st.image(a, caption=f"{i}. 施工後: {a.name}", use_container_width=True)

    st.divider()

    st.subheader("報告書の作成")
    btn_col1, btn_col2 = st.columns(2)
    with btn_col1:
        make_pdf = st.button("PDFで作成", type="primary", use_container_width=True)
    with btn_col2:
        make_docx = st.button("Wordで作成", use_container_width=True)

    if make_pdf or make_docx:
        if not property_name.strip():
            st.error("物件名を入力してください。")
        else:
            with tempfile.TemporaryDirectory() as tmpdir:
                # アップロードされた写真を一時フォルダに保存(PDF/Word共通で使う)
                pair_paths = []
                for i, (b, a) in enumerate(pairs_preview):
                    b_path = os.path.join(tmpdir, f"before_{i}_{b.name}")
                    a_path = os.path.join(tmpdir, f"after_{i}_{a.name}")
                    with open(b_path, "wb") as f:
                        f.write(b.getbuffer())
                    with open(a_path, "wb") as f:
                        f.write(a.getbuffer())
                    pair_paths.append((b_path, a_path))

                if make_pdf:
                    with st.spinner("PDFを生成しています..."):
                        pdf_path = os.path.join(tmpdir, "report.pdf")
                        build_pdf(property_name, pair_paths, pdf_path)
                        with open(pdf_path, "rb") as f:
                            pdf_bytes = f.read()
                    st.success("PDF報告書を作成しました。")
                    st.download_button(
                        "PDFをダウンロード",
                        data=pdf_bytes,
                        file_name=f"{property_name}_比較報告書.pdf",
                        mime="application/pdf",
                    )

                if make_docx:
                    with st.spinner("Wordファイルを生成しています..."):
                        docx_path = os.path.join(tmpdir, "report.docx")
                        build_docx(property_name, pair_paths, docx_path)
                        with open(docx_path, "rb") as f:
                            docx_bytes = f.read()
                    st.success("Word報告書を作成しました。")
                    st.download_button(
                        "Wordファイルをダウンロード",
                        data=docx_bytes,
                        file_name=f"{property_name}_比較報告書.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
else:
    st.info("施工前・施工後、それぞれの写真をアップロードしてください。")
