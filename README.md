# KQXS-ver2

Ứng dụng Python dùng để xem và xuất kết quả xổ số miền Nam theo ngày, sử dụng giao diện Tkinter và hỗ trợ export ra file Excel.

## Chương trình cho phép:
- Lấy kết quả (hôm nay hoặc hôm qua)
- Hiển thị theo từng đài (3 tỉnh mỗi ngày)
- Làm mới dữ liệu hoặc bật chế độ tự động cập nhật
- Xuất ra file `example.xlsx` với dữ liệu được ghi đúng vị trí và tự động xoá nội dung cũ trước khi ghi mới

## Các file chính gồm:
- `main.py`: chạy giao diện
- `process_data.py`: lấy và xử lý dữ liệu
- `expo_excel.py`: xuất Excel
- `example.xlsx`: file mẫu để ghi dữ liệu

## Cách chạy chương trình:
Yêu cầu Python 3.9+

## Cài thư viện:
- pip install openpyxl
- pip install pywin32


## Chạy lệnh:

python main.py


## Công nghệ sử dụng:
Python, Tkinter, openpyxl, win32com
