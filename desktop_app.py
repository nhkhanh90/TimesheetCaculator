import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta
import os
import threading
import io

class TimesheetCalculatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Hệ Thống Tính Giờ Công - v1.0")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # Thiết lập icon và style
        self.setup_styles()
        self.create_widgets()
        
        # Biến lưu trữ
        self.input_file_path = None
        self.output_file_path = None
        self.df_result = None
        
    def setup_styles(self):
        """Thiết lập style cho ứng dụng"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Cấu hình màu sắc
        style.configure('Title.TLabel', font=('Arial', 16, 'bold'), foreground='#2c3e50')
        style.configure('Header.TLabel', font=('Arial', 12, 'bold'), foreground='#34495e')
        style.configure('Info.TLabel', font=('Arial', 10), foreground='#7f8c8d')
        style.configure('Success.TLabel', font=('Arial', 10, 'bold'), foreground='#27ae60')
        style.configure('Error.TLabel', font=('Arial', 10, 'bold'), foreground='#e74c3c')
        
    def create_widgets(self):
        """Tạo giao diện chính"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Cấu hình grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="⏰ HỆ THỐNG TÍNH GIỜ CÔNG", style='Title.TLabel')
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Section 1: File Input
        input_frame = ttk.LabelFrame(main_frame, text="📁 Chọn File Excel", padding="15")
        input_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        input_frame.columnconfigure(1, weight=1)
        
        ttk.Label(input_frame, text="File Excel:", style='Header.TLabel').grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        self.file_path_var = tk.StringVar()
        self.file_entry = ttk.Entry(input_frame, textvariable=self.file_path_var, state='readonly')
        self.file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        self.browse_btn = ttk.Button(input_frame, text="Chọn File", command=self.browse_file)
        self.browse_btn.grid(row=0, column=2)
        
        # File info
        self.file_info_var = tk.StringVar()
        self.file_info_label = ttk.Label(input_frame, textvariable=self.file_info_var, style='Info.TLabel')
        self.file_info_label.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(10, 0))
        
        # Section 2: Rules
        rules_frame = ttk.LabelFrame(main_frame, text="⚙️ Quy Tắc Tính Giờ", padding="15")
        rules_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        rules_text = """• Giờ vào sớm nhất: 7h30 (nếu vào trước 7h30 thì tính là 7h30)
• Nghỉ trưa: Tự động trừ 1 tiếng
• Giờ thường: Tối đa 8 tiếng/ngày
• Giờ OT: Phần vượt quá 8 tiếng/ngày"""
        
        ttk.Label(rules_frame, text=rules_text, style='Info.TLabel').grid(row=0, column=0, sticky=tk.W)
        
        # Section 3: Process
        process_frame = ttk.LabelFrame(main_frame, text="🔄 Xử Lý Dữ Liệu", padding="15")
        process_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        process_frame.columnconfigure(0, weight=1)
        
        self.process_btn = ttk.Button(process_frame, text="🚀 Tính Toán Giờ Công", command=self.process_file, state='disabled')
        self.process_btn.grid(row=0, column=0, pady=(0, 10))
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(process_frame, variable=self.progress_var, mode='determinate')
        self.progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Status
        self.status_var = tk.StringVar(value="Sẵn sàng. Hãy chọn file Excel để bắt đầu.")
        self.status_label = ttk.Label(process_frame, textvariable=self.status_var, style='Info.TLabel')
        self.status_label.grid(row=2, column=0, sticky=tk.W)
        
        # Section 4: Results
        results_frame = ttk.LabelFrame(main_frame, text="📊 Kết Quả", padding="15")
        results_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        results_frame.columnconfigure(1, weight=1)
        
        # Statistics
        self.stats_frame = ttk.Frame(results_frame)
        self.stats_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Export section
        ttk.Label(results_frame, text="Xuất file:", style='Header.TLabel').grid(row=1, column=0, sticky=tk.W, padx=(0, 10))
        
        self.output_path_var = tk.StringVar()
        self.output_entry = ttk.Entry(results_frame, textvariable=self.output_path_var, state='readonly')
        self.output_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        self.export_btn = ttk.Button(results_frame, text="💾 Xuất Excel", command=self.export_file, state='disabled')
        self.export_btn.grid(row=1, column=2)
        
        # Section 5: Log
        log_frame = ttk.LabelFrame(main_frame, text="📝 Log", padding="15")
        log_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # Text widget with scrollbar
        text_frame = ttk.Frame(log_frame)
        text_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(text_frame, height=8, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # Clear log button
        ttk.Button(log_frame, text="🗑️ Xóa Log", command=self.clear_log).grid(row=1, column=0, pady=(10, 0))
        
        # Configure row weights
        main_frame.rowconfigure(5, weight=1)
        
    def log_message(self, message):
        """Thêm message vào log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def clear_log(self):
        """Xóa log"""
        self.log_text.delete(1.0, tk.END)
        
    def browse_file(self):
        """Chọn file Excel"""
        file_path = filedialog.askopenfilename(
            title="Chọn file Excel chấm công",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            self.input_file_path = file_path
            self.file_path_var.set(file_path)
            
            # Đọc thông tin file
            try:
                df = pd.read_excel(file_path)
                file_size = os.path.getsize(file_path) / 1024  # KB
                self.file_info_var.set(f"✅ File hợp lệ - {len(df)} dòng, {len(df.columns)} cột, {file_size:.1f} KB")
                
                # Kiểm tra cột cần thiết
                required_columns = ['Họ tên', 'Ngày chốt', 'Giờ chốt']
                missing_columns = [col for col in required_columns if col not in df.columns]
                
                if missing_columns:
                    self.file_info_var.set(f"❌ Thiếu cột: {', '.join(missing_columns)}")
                    self.process_btn.configure(state='disabled')
                    self.log_message(f"❌ File thiếu cột bắt buộc: {', '.join(missing_columns)}")
                    self.log_message(f"Các cột hiện có: {', '.join(df.columns.tolist())}")
                else:
                    self.process_btn.configure(state='normal')
                    self.log_message(f"✅ Đã chọn file: {os.path.basename(file_path)}")
                    
            except Exception as e:
                self.file_info_var.set(f"❌ Lỗi đọc file: {str(e)}")
                self.process_btn.configure(state='disabled')
                self.log_message(f"❌ Lỗi đọc file: {str(e)}")
    
    def preprocess_timesheet_data(self, df):
        """
        Tiền xử lý dữ liệu timesheet với cấu trúc Ngày chốt và Giờ chốt
        Hỗ trợ cả định dạng 24h (HH:MM:SS) và 12h (HH:MM:SS AM/PM)
        """
        df_processed = df.copy()
        
        # Chuẩn hóa tên cột
        df_processed = df_processed.rename(columns={
            'Họ tên': 'employee_name',
            'Ngày chốt': 'date',
            'Giờ chốt': 'time_stamp'
        })
        
        # Chuyển đổi cột ngày
        df_processed['date'] = pd.to_datetime(df_processed['date'])
        
        # Chuyển đổi cột giờ - hỗ trợ nhiều định dạng
        def parse_time_flexible(time_str):
            """
            Parse thời gian linh hoạt cho cả 24h và 12h format
            """
            if pd.isna(time_str):
                return None
                
            time_str = str(time_str).strip()
            
            # Danh sách các format để thử
            time_formats = [
                '%H:%M:%S',           # 14:30:00
                '%I:%M:%S %p',        # 02:30:00 PM
                '%H:%M',              # 14:30
                '%I:%M %p',           # 02:30 PM
                '%I:%M:%S%p',         # 02:30:00PM (không có space)
                '%I:%M%p',            # 02:30PM (không có space)
            ]
            
            for fmt in time_formats:
                try:
                    parsed_time = datetime.strptime(time_str, fmt)
                    return parsed_time.time()
                except ValueError:
                    continue
            
            # Nếu không parse được, thử regex để extract
            import re
            
            # Pattern cho 12-hour format với AM/PM
            pattern_12h = r'(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM|am|pm)'
            match = re.search(pattern_12h, time_str, re.IGNORECASE)
            
            if match:
                hour = int(match.group(1))
                minute = int(match.group(2))
                second = int(match.group(3)) if match.group(3) else 0
                am_pm = match.group(4).upper()
                
                # Chuyển đổi 12h sang 24h
                if am_pm == 'PM' and hour != 12:
                    hour += 12
                elif am_pm == 'AM' and hour == 12:
                    hour = 0
                    
                return time(hour, minute, second)
            
            # Pattern cho 24-hour format
            pattern_24h = r'(\d{1,2}):(\d{2})(?::(\d{2}))?'
            match = re.search(pattern_24h, time_str)
            
            if match:
                hour = int(match.group(1))
                minute = int(match.group(2))
                second = int(match.group(3)) if match.group(3) else 0
                
                if 0 <= hour <= 23 and 0 <= minute <= 59 and 0 <= second <= 59:
                    return time(hour, minute, second)
            
            return None
        
        # Áp dụng parse function
        df_processed['time_stamp'] = df_processed['time_stamp'].apply(parse_time_flexible)
        
        # Loại bỏ những dòng có thời gian null
        df_processed = df_processed.dropna(subset=['time_stamp'])
        
        # Sắp xếp theo nhân viên, ngày và giờ
        df_processed = df_processed.sort_values(['employee_name', 'date', 'time_stamp'])
        
        return df_processed
    
    def calculate_work_hours(self, df):
        """Tính toán giờ làm việc và giờ OT"""
        result_data = []
        total_groups = len(df.groupby(['employee_name', 'date']))
        current_group = 0
        
        # Group theo nhân viên và ngày
        for (employee, date), group in df.groupby(['employee_name', 'date']):
            current_group += 1
            progress = (current_group / total_groups) * 100
            self.progress_var.set(progress)
            self.root.update_idletasks()
            
            # Sắp xếp theo thời gian trong ngày
            group = group.sort_values('time_stamp')
            times = group['time_stamp'].tolist()
            
            if len(times) >= 2:
                # Lấy thời gian đầu tiên là check_in, cuối cùng là check_out
                check_in = times[0]
                check_out = times[-1]
                
                # Áp dụng quy tắc: nếu vào trước 7:30 thì tính là 7:30
                earliest_start = time(7, 30)
                if check_in < earliest_start:
                    adjusted_check_in = earliest_start
                else:
                    adjusted_check_in = check_in
                
                # Tính tổng thời gian
                check_in_dt = datetime.combine(datetime.today(), adjusted_check_in)
                check_out_dt = datetime.combine(datetime.today(), check_out)
                
                # Nếu check_out < check_in (qua ngày), thêm 1 ngày
                if check_out_dt < check_in_dt:
                    check_out_dt += timedelta(days=1)
                
                total_time = check_out_dt - check_in_dt
                total_hours = total_time.total_seconds() / 3600
                
                # Trừ 1 tiếng nghỉ trưa
                work_time = max(0, total_hours - 1)
                
                # Tính giờ thường và OT
                # Giờ thường tối đa 8 tiếng/ngày
                regular_hours = min(work_time, 8)
                overtime_hours = max(0, work_time - 8)
                
                result_data.append({
                    'employee_name': employee,
                    'date': date,
                    'check_in_original': str(check_in),
                    'check_out_original': str(check_out),
                    'check_in_adjusted': str(adjusted_check_in),
                    'total_time_with_lunch': round(total_hours, 2),
                    'work_hours': round(work_time, 2),
                    'regular_hours': round(regular_hours, 2),
                    'overtime_hours': round(overtime_hours, 2),
                    'number_of_punches': len(times)
                })
            else:
                # Không đủ dữ liệu chấm công
                result_data.append({
                    'employee_name': employee,
                    'date': date,
                    'check_in_original': str(times[0]) if times else 'N/A',
                    'check_out_original': 'N/A',
                    'check_in_adjusted': 'N/A',
                    'total_time_with_lunch': 0,
                    'work_hours': 0,
                    'regular_hours': 0,
                    'overtime_hours': 0,
                    'number_of_punches': len(times)
                })
        
        return pd.DataFrame(result_data)
    
    def update_statistics(self, df_result):
        """Cập nhật thống kê"""
        # Xóa stats cũ
        for widget in self.stats_frame.winfo_children():
            widget.destroy()
        
        # Tính toán thống kê
        total_employees = df_result['employee_name'].nunique()
        total_work_days = len(df_result)
        total_work_hours = df_result['work_hours'].sum()
        total_ot_hours = df_result['overtime_hours'].sum()
        
        # Hiển thị thống kê
        stats = [
            ("👥 Nhân viên:", total_employees),
            ("📅 Ngày làm việc:", total_work_days),
            ("⏰ Tổng giờ LV:", f"{total_work_hours:.1f}h"),
            ("🔥 Tổng giờ OT:", f"{total_ot_hours:.1f}h")
        ]
        
        for i, (label, value) in enumerate(stats):
            ttk.Label(self.stats_frame, text=label, style='Header.TLabel').grid(row=0, column=i*2, sticky=tk.W, padx=(0, 5))
            ttk.Label(self.stats_frame, text=str(value), style='Success.TLabel').grid(row=0, column=i*2+1, sticky=tk.W, padx=(0, 20))
    
    def process_file(self):
        """Xử lý file Excel"""
        if not self.input_file_path:
            messagebox.showerror("Lỗi", "Vui lòng chọn file Excel trước!")
            return
        
        def process_thread():
            try:
                self.process_btn.configure(state='disabled')
                self.progress_var.set(0)
                self.status_var.set("Đang đọc file Excel...")
                self.log_message("🔄 Bắt đầu xử lý file...")
                
                # Đọc file Excel
                df = pd.read_excel(self.input_file_path)
                self.log_message(f"✅ Đã đọc {len(df)} dòng dữ liệu")
                
                # Tiền xử lý
                self.status_var.set("Đang tiền xử lý dữ liệu...")
                df_processed = self.preprocess_timesheet_data(df)
                self.log_message("✅ Hoàn thành tiền xử lý dữ liệu")
                
                # Tính toán giờ công
                self.status_var.set("Đang tính toán giờ công...")
                self.df_result = self.calculate_work_hours(df_processed)
                self.log_message(f"✅ Đã tính toán {len(self.df_result)} ngày làm việc")
                
                # Cập nhật thống kê
                self.update_statistics(self.df_result)
                
                # Kiểm tra dữ liệu thiếu
                incomplete_data = self.df_result[self.df_result['number_of_punches'] < 2]
                if len(incomplete_data) > 0:
                    self.log_message(f"⚠️ Phát hiện {len(incomplete_data)} ngày có dữ liệu không đầy đủ")
                
                self.progress_var.set(100)
                self.status_var.set("✅ Hoàn thành! Sẵn sàng xuất file Excel.")
                self.export_btn.configure(state='normal')
                self.log_message("🎉 Xử lý hoàn tất! Bạn có thể xuất file Excel.")
                
            except Exception as e:
                self.log_message(f"❌ Lỗi xử lý: {str(e)}")
                self.status_var.set(f"❌ Lỗi: {str(e)}")
                messagebox.showerror("Lỗi xử lý", f"Đã xảy ra lỗi:\n{str(e)}")
            
            finally:
                self.process_btn.configure(state='normal')
        
        # Chạy trong thread riêng
        threading.Thread(target=process_thread, daemon=True).start()
    
    def export_file(self):
        """Xuất file Excel"""
        if self.df_result is None:
            messagebox.showerror("Lỗi", "Chưa có dữ liệu để xuất!")
            return
        
        # Chọn vị trí lưu file
        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"bao_cao_gio_cong_{current_time}.xlsx"
        
        file_path = filedialog.asksaveasfilename(
            title="Lưu báo cáo Excel",
            defaultextension=".xlsx",
            initialname=default_filename,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                self.status_var.set("Đang xuất file Excel...")
                self.log_message(f"📤 Đang xuất file: {os.path.basename(file_path)}")
                
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    # Sheet 1: Dữ liệu chi tiết
                    self.df_result.to_excel(writer, sheet_name='Chi tiết giờ công & OT', index=False)
                    
                    # Sheet 2: Thống kê theo nhân viên
                    employee_stats = self.df_result.groupby('employee_name').agg({
                        'work_hours': ['sum', 'mean', 'count'],
                        'regular_hours': ['sum', 'mean'],
                        'overtime_hours': ['sum', 'mean']
                    }).round(2)
                    employee_stats.columns = [
                        'Tổng giờ LV', 'TB giờ LV/ngày', 'Số ngày làm',
                        'Tổng giờ thường', 'TB giờ thường/ngày',
                        'Tổng giờ OT', 'TB giờ OT/ngày'
                    ]
                    employee_stats.to_excel(writer, sheet_name='Thống kê nhân viên')
                    
                    # Sheet 3: Thống kê theo ngày
                    daily_stats = self.df_result.groupby('date').agg({
                        'work_hours': ['sum', 'mean'],
                        'regular_hours': 'sum',
                        'overtime_hours': 'sum',
                        'employee_name': 'count'
                    }).round(2)
                    daily_stats.columns = ['Tổng giờ LV', 'TB giờ/người', 'Tổng giờ thường', 'Tổng giờ OT', 'Số người']
                    daily_stats.to_excel(writer, sheet_name='Thống kê theo ngày')
                    
                    # Sheet 4: Ranking OT
                    ot_ranking = self.df_result.groupby('employee_name').agg({
                        'overtime_hours': ['sum', 'mean', 'count']
                    }).round(2)
                    ot_ranking.columns = ['Tổng giờ OT', 'TB giờ OT/ngày', 'Số ngày có OT']
                    ot_ranking = ot_ranking.sort_values('Tổng giờ OT', ascending=False)
                    ot_ranking.to_excel(writer, sheet_name='Ranking OT')
                
                self.output_path_var.set(file_path)
                self.status_var.set("✅ Đã xuất file thành công!")
                self.log_message(f"✅ Đã xuất file thành công: {os.path.basename(file_path)}")
                
                # Hỏi có muốn mở file không
                if messagebox.askyesno("Thành công", "Đã xuất file thành công!\nBạn có muốn mở file này không?"):
                    os.startfile(file_path)
                
            except Exception as e:
                self.log_message(f"❌ Lỗi xuất file: {str(e)}")
                messagebox.showerror("Lỗi xuất file", f"Đã xảy ra lỗi khi xuất file:\n{str(e)}")

def main():
    root = tk.Tk()
    app = TimesheetCalculatorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
