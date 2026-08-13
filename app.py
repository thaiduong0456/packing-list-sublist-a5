from __future__ import annotations

from io import BytesIO
import pymupdf
import streamlit as st
from openpyxl import Workbook

from sublist import build_sublist_pdf, parse_packing_list


st.set_page_config(page_title="Packing List → Sublist A5", page_icon="📦", layout="centered")
st.title("Packing List → Sublist A5")
st.caption("Tạo một trang sublist A5 cho mỗi carton, sẵn sàng để in ở 100% / Actual size.")

def make_sample_workbook() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Packing List"
    sheet.append(["OR No.", "Ref No.", "SKU#", "BarCode/UPC", "Quantity", "Carton#", "Packaging code", "Weight (KG)"])
    sheet.append(["OR SAMPLE", "REF001", "SKU-SAMPLE-01", "4890000000001", 5, "1/1", "PKG-SAMPLE-0001", 2.5])
    sheet.append(["OR SAMPLE", "REF001", "SKU-SAMPLE-02", "4890000000002", 3, None, None, None])
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


st.download_button(
    "Tải file Excel mẫu",
    make_sample_workbook(),
    file_name="packing_list_mau.xlsx",
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
