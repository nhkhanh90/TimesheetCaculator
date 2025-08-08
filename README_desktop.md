# Hệ Thống Tính Giờ Công - Desktop App

Ứng dụng Windows desktop để tính toán giờ công từ dữ liệu chấm công Excel.

## Tính năng

- ✅ Giao diện thân thiện với Windows
- ✅ Chọn file Excel dễ dàng  
- ✅ Tính giờ làm việc và giờ OT tự động
- ✅ Hiển thị thống kê realtime
- ✅ Xuất báo cáo Excel chi tiết
- ✅ Log xử lý chi tiết
- ✅ Chạy độc lập không cần cài Python

## Cài đặt cho Developer

1. Cài đặt dependencies:
```bash
pip install -r requirements_desktop.txt
```

2. Chạy ứng dụng:
```bash
python desktop_app.py
```

## Build thành file .exe

1. Chạy script build:
```bash
build.bat
```

2. File `TimesheetCalculator.exe` sẽ được tạo

## Sử dụng

### Cho User cuối:
1. Chạy file `TimesheetCalculator.exe`
2. Click "Chọn File" để chọn file Excel
3. Click "🚀 Tính Toán Giờ Công"
4. Click "💾 Xuất Excel" để lưu kết quả

### Cấu trúc file Excel đầu vào:
File Excel cần có các cột:
- **Họ tên**: Tên nhân viên
- **Ngày chốt**: Ngày chấm công (YYYY-MM-DD)
- **Giờ chốt**: Thời gian chấm công (HH:MM:SS)

### Quy tắc tính giờ:
- **Giờ vào sớm nhất**: 7h30 (nếu chấm công trước 7h30 thì tính là 7h30)
- **Nghỉ trưa**: Tự động trừ 1 tiếng
- **Giờ thường**: Tối đa 8 tiếng/ngày
- **Giờ OT**: Phần vượt quá 8 tiếng/ngày

### Kết quả:
File Excel xuất ra bao gồm 4 sheet:
1. **Chi tiết giờ công & OT**: Dữ liệu từng ngày của mỗi nhân viên
2. **Thống kê nhân viên**: Tổng hợp theo nhân viên
3. **Thống kê theo ngày**: Tổng hợp theo ngày
4. **Ranking OT**: Xếp hạng giờ OT

## Cấu trúc thư mục

```
Timesheet_cal/
├── desktop_app.py              # Ứng dụng desktop chính
├── app.py                      # Ứng dụng web Streamlit
├── build.bat                   # Script build .exe
├── requirements_desktop.txt    # Dependencies cho desktop app
├── requirements.txt           # Dependencies cho web app
├── README_desktop.md          # Hướng dẫn desktop app
├── README.md                  # Hướng dẫn web app
├── TimesheetCalculator.exe    # File executable (sau khi build)
└── data/                      # Thư mục chứa file mẫu
    └── file1.xlsx            # File Excel mẫu
```

## Yêu cầu hệ thống

- Windows 7/8/10/11
- RAM: 4GB trở lên
- Dung lượng: 100MB trống

## Hỗ trợ

Nếu có vấn đề, vui lòng kiểm tra:
1. File Excel có đúng format không
2. Các cột bắt buộc có đầy đủ không
3. Dữ liệu thời gian có đúng format HH:MM:SS không
