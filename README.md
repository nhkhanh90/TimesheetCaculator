# Hệ Thống Tính Giờ Công

Ứng dụng web để tính toán giờ công từ dữ liệu chấm công Excel.

## Tính năng

- ✅ Upload file Excel và tự động xử lý
- ✅ Tính giờ làm việc và giờ OT
- ✅ Áp dụng quy tắc giờ vào sớm nhất 7h30
- ✅ Tự động trừ 1 tiếng nghỉ trưa
- ✅ Phân biệt giờ thường (8h) và giờ OT
- ✅ Báo cáo chi tiết theo nhân viên và ngày
- ✅ Xuất file Excel với nhiều sheet thống kê

## Cài đặt

1. Cài đặt Python dependencies:
```bash
pip install -r requirements.txt
```

2. Chạy ứng dụng:
```bash
streamlit run app.py
```

3. Mở trình duyệt tại: http://localhost:8501

## Cấu trúc file Excel đầu vào

File Excel cần có các cột:
- **Họ tên**: Tên nhân viên
- **Ngày chốt**: Ngày chấm công (YYYY-MM-DD)
- **Giờ chốt**: Thời gian chấm công (HH:MM:SS)

## Quy tắc tính giờ

- **Giờ vào sớm nhất**: 7h30 (nếu chấm công trước 7h30 thì tính là 7h30)
- **Nghỉ trưa**: Tự động trừ 1 tiếng
- **Giờ thường**: Tối đa 8 tiếng/ngày
- **Giờ OT**: Phần vượt quá 8 tiếng/ngày

## Kết quả

File Excel xuất ra bao gồm 5 sheet:
1. **Chi tiết giờ công & OT**: Dữ liệu từng ngày của mỗi nhân viên
2. **Thống kê nhân viên**: Tổng hợp theo nhân viên
3. **Thống kê theo ngày**: Tổng hợp theo ngày
4. **Ranking OT**: Xếp hạng giờ OT
5. **Dữ liệu gốc**: Dữ liệu chấm công gốc

## Cấu trúc thư mục

```
Timesheet_cal/
├── app.py              # Ứng dụng chính
├── requirements.txt    # Dependencies
├── README.md          # Hướng dẫn
└── data/              # Thư mục chứa file mẫu
    └── file1.xlsx     # File Excel mẫu
```
