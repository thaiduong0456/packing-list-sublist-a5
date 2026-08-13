# Prompt tạo tool Packing List → Sublist A5

Hãy xây dựng hoàn chỉnh một web tool bằng Python + Streamlit để người dùng tải packing list Excel và tải xuống PDF sublist khổ A5 dọc. Không dừng ở pseudocode.

## Dữ liệu đầu vào

- File `.xlsx`; tự tìm sheet packing list và dòng tiêu đề, không giả định tên sheet hay số dòng cố định.
- Nhận diện các cột theo tên/alias: OR No., Ref No., SKU#, Barcode/UPC hoặc EAN, Quantity/QTY, Carton#, Packaging code/Packing code, Weight/Gross Weight.
- Chuyển mã sang chuỗi để tránh scientific notation và cố giữ số 0 đầu theo định dạng Excel.
- Một dòng có Carton# mở nhóm carton mới. Các dòng Carton# trống phía sau thuộc carton gần nhất.
- Giữ nguyên thứ tự carton và SKU. Bỏ qua TOTAL, GRAND TOTAL, 合计, 總計.

## PDF đầu ra

- PDF vector khổ A5 dọc 148 × 210 mm; mỗi carton một trang; in ở Actual size / 100%.
- Phần đầu trang theo ảnh tham chiếu: Carton #, OR #, Ref #, GW, Packing Code #.
- OR hiển thị dấu gạch dưới `_` thành khoảng trắng.
- Bảng gồm Item No., EAN, QTY; có đường kẻ rõ, header nền xám, tổng QTY cuối bảng.
- Font và chiều cao dòng tự co theo số SKU nhưng không để chồng chữ hoặc tràn trang.

## Giao diện

- Ba luồng chính: tải file Excel mẫu, tải packing list, xem trước từng trang dưới dạng PNG và tải PDF.
- Preview bằng PyMuPDF, không nhúng PDF data URL trong iframe.
- Báo lỗi thân thiện nếu thiếu cột hoặc không có carton hợp lệ.

## Kỹ thuật và kiểm thử

- Dùng openpyxl, ReportLab, Streamlit và PyMuPDF.
- Tách hàm đọc Excel, nhóm carton, tạo PDF và UI để có thể kiểm thử.
- Kiểm thử bằng file thật: số carton bằng số trang, kiểm tra carton nhiều SKU nhất, kích thước trang đúng A5, dữ liệu SKU/EAN/QTY và tổng QTY chính xác.
- Tạo `requirements.txt`, `.streamlit/config.toml`, `.gitignore`, `README.md` và hướng dẫn deploy Streamlit Community Cloud.
