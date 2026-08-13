from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pymupdf
import streamlit as st

from sublist import build_sublist_pdf, parse_packing_list


st.set_page_config(page_title="Packing List → Sublist A5", page_icon="📦", layout="centered")
st.title("Packing List → Sublist A5")
st.caption("Tạo một trang sublist A5 cho mỗi carton, sẵn sàng để in ở 100% / Actual size.")

sample_path = Path(__file__).with_name("1.Packinglist_Total_HK_4pcs.xlsx")
st.download_button(
    "Tải file Excel mẫu",
    sample_path.read_bytes(),
    file_name=sample_path.name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

uploaded = st.file_uploader("Tải packing list Excel", type=["xlsx"])
if uploaded:
    try:
        cartons = parse_packing_list(BytesIO(uploaded.getvalue()))
        pdf_bytes = build_sublist_pdf(cartons)
    except Exception as exc:
        st.error(f"Không thể xử lý file: {exc}")
        st.stop()

    st.success(f"Đã tạo {len(cartons)} carton / {len(cartons)} trang A5.")
    st.download_button(
        "Tải xuống Sublist PDF A5",
        pdf_bytes,
        file_name=f"{uploaded.name.rsplit('.', 1)[0]}_sublist_A5.pdf",
        mime="application/pdf",
        type="primary",
    )

    st.subheader("Xem trước")
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    for page_number, page in enumerate(doc, start=1):
        pix = page.get_pixmap(matrix=pymupdf.Matrix(1.35, 1.35), alpha=False)
        st.image(pix.tobytes("png"), caption=f"Carton {page_number}", use_container_width=True)
