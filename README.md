# Packing List → Sublist A5

Ứng dụng Streamlit đọc packing list Excel và tạo PDF sublist khổ A5 dọc, mỗi carton một trang.

## Chạy trên máy

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Quy tắc dữ liệu

- Tự tìm sheet và dòng tiêu đề theo tên cột, không phụ thuộc số dòng cố định.
- Dòng có `Carton#` bắt đầu carton mới; dòng trống tiếp theo thuộc carton gần nhất.
- Loại các dòng `TOTAL`, `GRAND TOTAL`, `合计`, `總計`.
- Giữ nguyên thứ tự carton và SKU.
- PDF A5 dọc 148 × 210 mm, in ở `Actual size / 100%`.

