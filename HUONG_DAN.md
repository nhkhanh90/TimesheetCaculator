# HƯỚNG DẪN SỬ DỤNG HỆ THỐNG TÍNH GIỜ CÔNG

## 🚀 Cách sử dụng nhanh

### Ứng dụng Web (Streamlit):
1. Mở terminal/command prompt tại thư mục này
2. Chạy: `streamlit run app.py`
3. Mở trình duyệt tại: http://localhost:8501
4. Upload file Excel và tải xuống kết quả

### Ứng dụng Desktop:
1. Chạy file: `TimesheetCalculator.exe` (nếu đã build)
2. Hoặc chạy: `python desktop_app.py`
3. Chọn file Excel và xuất kết quả

## 📁 Cấu trúc file Excel cần thiết

File Excel phải có **3 cột bắt buộc**:
- **Họ tên**: Tên nhân viên
- **Ngày chốt**: Ngày chấm công (YYYY-MM-DD)
- **Giờ chốt**: Thời gian chấm công (HH:MM:SS)

## ⚙️ Quy tắc tính giờ

- **Giờ vào sớm nhất**: 7h30 (nếu chấm trước 7h30 → tính 7h30)
- **Nghỉ trưa**: Tự động trừ 1 tiếng
- **Giờ thường**: Tối đa 8 tiếng/ngày
- **Giờ OT**: Phần vượt quá 8 tiếng/ngày

## 📊 Kết quả xuất ra

File Excel có 5 sheet:
1. **Chi tiết giờ công & OT** - Dữ liệu từng ngày
2. **Thống kê nhân viên** - Tổng hợp theo nhân viên  
3. **Thống kê theo ngày** - Tổng hợp theo ngày
4. **Ranking OT** - Xếp hạng giờ OT
5. **Dữ liệu gốc** - Dữ liệu chấm công ban đầu

## 🔧 Build file .exe

Để tạo file executable cho người dùng:
```bash
# Cách 1: Dùng script tự động
build_simple.bat

# Cách 2: Thủ công
pip install pyinstaller
pyinstaller --onefile --windowed --name="TimesheetCalculator" desktop_app.py
```

## ❗ Xử lý lỗi thường gặp

1. **Lỗi thiếu cột**: Kiểm tra file Excel có đủ 3 cột bắt buộc
2. **Lỗi format thời gian**: Đảm bảo cột "Giờ chốt" có format HH:MM:SS
3. **Lỗi Arrow/PyArrow**: Đã được fix trong code mới
4. **Lỗi export Excel**: Kiểm tra quyền ghi file và đường dẫn

## 📞 Hỗ trợ

- Kiểm tra file mẫu trong thư mục `data/`
- Xem log chi tiết trong ứng dụng
- Đảm bảo có quyền đọc/ghi file Excel
