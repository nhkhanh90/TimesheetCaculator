import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta
import io
import warnings
warnings.filterwarnings('ignore')

# Cấu hình trang
st.set_page_config(
    page_title="Hệ Thống Tính Giờ Công",
    page_icon="⏰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS tùy chỉnh
st.markdown("""
<style>
    .main-header {
        text-align: center;
        color: #1e3a8a;
        font-size: 2.5rem;
        font-weight: bold;
        margin-bottom: 2rem;
    }
    .sub-header {
        color: #3b82f6;
        font-size: 1.2rem;
        margin-bottom: 1rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin: 0.5rem 0;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .warning-box {
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        color: #856404;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def preprocess_timesheet_data(df):
    """
    Tiền xử lý dữ liệu timesheet với cấu trúc Ngày chốt và Giờ chốt
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
    
    # Chuyển đổi cột giờ - sử dụng string để tránh lỗi Arrow
    try:
        # Thử chuyển đổi thành datetime rồi lấy time
        df_processed['time_stamp'] = pd.to_datetime(df_processed['time_stamp'], format='%H:%M:%S').dt.time
        # Chuyển thành string để tương thích với Arrow
        df_processed['time_stamp_str'] = df_processed['time_stamp'].astype(str)
    except:
        # Nếu lỗi, giữ nguyên dạng string
        df_processed['time_stamp_str'] = df_processed['time_stamp'].astype(str)
        # Parse lại thành time object
        df_processed['time_stamp'] = pd.to_datetime(df_processed['time_stamp_str'], format='%H:%M:%S').dt.time
    
    # Sắp xếp theo nhân viên, ngày và giờ
    df_processed = df_processed.sort_values(['employee_name', 'date', 'time_stamp'])
    
    return df_processed

def calculate_work_hours(df):
    """
    Tính toán giờ làm việc và giờ OT từ dữ liệu chấm công
    """
    result_data = []
    
    # Group theo nhân viên và ngày
    for (employee, date), group in df.groupby(['employee_name', 'date']):
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
                'check_in_original': str(check_in),  # Chuyển thành string
                'check_out_original': str(check_out),  # Chuyển thành string
                'check_in_adjusted': str(adjusted_check_in),  # Chuyển thành string
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

def create_excel_report(df_final, df_raw):
    """
    Tạo file Excel báo cáo với nhiều sheet
    """
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: Dữ liệu chi tiết
        df_final.to_excel(writer, sheet_name='Chi tiết giờ công & OT', index=False)
        
        # Sheet 2: Thống kê theo nhân viên
        employee_stats = df_final.groupby('employee_name').agg({
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
        daily_stats = df_final.groupby('date').agg({
            'work_hours': ['sum', 'mean'],
            'regular_hours': 'sum',
            'overtime_hours': 'sum',
            'employee_name': 'count'
        }).round(2)
        daily_stats.columns = ['Tổng giờ LV', 'TB giờ/người', 'Tổng giờ thường', 'Tổng giờ OT', 'Số người']
        daily_stats.to_excel(writer, sheet_name='Thống kê theo ngày')
        
        # Sheet 4: Ranking OT
        ot_ranking = df_final.groupby('employee_name').agg({
            'overtime_hours': ['sum', 'mean', 'count']
        }).round(2)
        ot_ranking.columns = ['Tổng giờ OT', 'TB giờ OT/ngày', 'Số ngày có OT']
        ot_ranking = ot_ranking.sort_values('Tổng giờ OT', ascending=False)
        ot_ranking.to_excel(writer, sheet_name='Ranking OT')
        
        # Sheet 5: Dữ liệu gốc
        df_raw.to_excel(writer, sheet_name='Dữ liệu gốc', index=False)
    
    output.seek(0)
    return output

# Giao diện chính
def main():
    # Header
    st.markdown('<h1 class="main-header">⏰ HỆ THỐNG TÍNH GIỜ CÔNG</h1>', unsafe_allow_html=True)
    
    # Sidebar - Hướng dẫn
    with st.sidebar:
        st.markdown("### 📋 Hướng dẫn sử dụng")
        st.markdown("""
        1. **Upload file Excel** có cấu trúc:
           - Cột "Họ tên": Tên nhân viên
           - Cột "Ngày chốt": Ngày chấm công
           - Cột "Giờ chốt": Thời gian chấm công
        
        2. **Quy tắc tính giờ**:
           - Giờ vào sớm nhất: 7h30
           - Nghỉ trưa: 1 tiếng
           - Giờ thường: Tối đa 8h/ngày
           - Giờ OT: Vượt quá 8h/ngày
        
        3. **Kết quả** sẽ bao gồm:
           - Chi tiết giờ công từng ngày
           - Thống kê theo nhân viên
           - Thống kê theo ngày
           - Ranking OT
        """)
        
        st.markdown("### ⚙️ Cài đặt")
        show_raw_data = st.checkbox("Hiển thị dữ liệu gốc", value=False)
        show_detailed_stats = st.checkbox("Hiển thị thống kê chi tiết", value=True)
    
    # Upload file
    st.markdown('<p class="sub-header">📁 Upload File Excel</p>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "Chọn file Excel chấm công",
        type=['xlsx', 'xls'],
        help="File Excel cần có các cột: Họ tên, Ngày chốt, Giờ chốt"
    )
    
    if uploaded_file is not None:
        try:
            # Đọc file Excel
            with st.spinner('Đang đọc file Excel...'):
                df = pd.read_excel(uploaded_file)
            
            # Hiển thị thông tin file
            st.markdown('<div class="success-box">✅ Đã đọc file thành công!</div>', unsafe_allow_html=True)
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Số dòng", len(df))
            with col2:
                st.metric("Số cột", len(df.columns))
            with col3:
                st.metric("Kích thước", f"{uploaded_file.size / 1024:.1f} KB")
            
            # Kiểm tra cột cần thiết
            required_columns = ['Họ tên', 'Ngày chốt', 'Giờ chốt']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                st.markdown(
                    f'<div class="warning-box">⚠️ Thiếu các cột: {", ".join(missing_columns)}</div>',
                    unsafe_allow_html=True
                )
                st.markdown("**Các cột hiện có:**")
                st.write(df.columns.tolist())
                return
            
            # Hiển thị dữ liệu mẫu
            if show_raw_data:
                st.markdown('<p class="sub-header">📋 Dữ liệu gốc (5 dòng đầu)</p>', unsafe_allow_html=True)
                st.dataframe(df.head(), use_container_width=True)
            
            # Xử lý dữ liệu
            with st.spinner('Đang xử lý và tính toán giờ công...'):
                df_processed = preprocess_timesheet_data(df)
                df_result = calculate_work_hours(df_processed)
            
            # Thống kê tổng quan
            st.markdown('<p class="sub-header">📊 Thống kê tổng quan</p>', unsafe_allow_html=True)
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.markdown(
                    f'<div class="metric-card"><h3>{df_result["employee_name"].nunique()}</h3><p>Nhân viên</p></div>',
                    unsafe_allow_html=True
                )
            
            with col2:
                st.markdown(
                    f'<div class="metric-card"><h3>{len(df_result)}</h3><p>Ngày làm việc</p></div>',
                    unsafe_allow_html=True
                )
            
            with col3:
                st.markdown(
                    f'<div class="metric-card"><h3>{df_result["work_hours"].sum():.1f}h</h3><p>Tổng giờ LV</p></div>',
                    unsafe_allow_html=True
                )
            
            with col4:
                st.markdown(
                    f'<div class="metric-card"><h3>{df_result["overtime_hours"].sum():.1f}h</h3><p>Tổng giờ OT</p></div>',
                    unsafe_allow_html=True
                )
            
            # Chi tiết kết quả
            st.markdown('<p class="sub-header">📈 Kết quả chi tiết</p>', unsafe_allow_html=True)
            
            # Tabs cho các loại báo cáo
            tab1, tab2, tab3, tab4 = st.tabs(["📋 Chi tiết", "👥 Theo nhân viên", "📅 Theo ngày", "🏆 Ranking OT"])
            
            with tab1:
                display_df = df_result[['employee_name', 'date', 'check_in_original', 'check_out_original', 
                                      'work_hours', 'regular_hours', 'overtime_hours']].copy()
                display_df.columns = ['Nhân viên', 'Ngày', 'Giờ vào', 'Giờ ra', 'Giờ LV', 'Giờ thường', 'Giờ OT']
                st.dataframe(display_df, use_container_width=True)
            
            with tab2:
                employee_summary = df_result.groupby('employee_name').agg({
                    'work_hours': ['sum', 'mean', 'count'],
                    'regular_hours': 'sum',
                    'overtime_hours': 'sum'
                }).round(2)
                employee_summary.columns = ['Tổng giờ LV', 'TB giờ/ngày', 'Số ngày', 'Tổng giờ thường', 'Tổng giờ OT']
                st.dataframe(employee_summary, use_container_width=True)
            
            with tab3:
                daily_summary = df_result.groupby('date').agg({
                    'work_hours': ['sum', 'mean'],
                    'regular_hours': 'sum',
                    'overtime_hours': 'sum',
                    'employee_name': 'count'
                }).round(2)
                daily_summary.columns = ['Tổng giờ LV', 'TB giờ/người', 'Tổng giờ thường', 'Tổng giờ OT', 'Số người']
                st.dataframe(daily_summary, use_container_width=True)
            
            with tab4:
                ot_ranking = df_result.groupby('employee_name')['overtime_hours'].sum().sort_values(ascending=False)
                if ot_ranking.sum() > 0:
                    st.bar_chart(ot_ranking.head(10))
                    st.dataframe(ot_ranking.head(10).to_frame('Tổng giờ OT'), use_container_width=True)
                else:
                    st.info("Không có nhân viên nào làm OT trong khoảng thời gian này.")
            
            # Tạo và download file Excel
            st.markdown('<p class="sub-header">💾 Tải xuống báo cáo</p>', unsafe_allow_html=True)
            
            excel_file = create_excel_report(df_result, df_processed)
            
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"bao_cao_gio_cong_{current_time}.xlsx"
            
            st.download_button(
                label="📥 Tải xuống báo cáo Excel",
                data=excel_file.getvalue(),
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            # Cảnh báo nếu có dữ liệu không đầy đủ
            incomplete_data = df_result[df_result['number_of_punches'] < 2]
            if len(incomplete_data) > 0:
                st.markdown(
                    f'<div class="warning-box">⚠️ Phát hiện {len(incomplete_data)} ngày có dữ liệu chấm công không đầy đủ (thiếu giờ vào hoặc giờ ra)</div>',
                    unsafe_allow_html=True
                )
                with st.expander("Xem chi tiết các ngày thiếu dữ liệu"):
                    st.dataframe(incomplete_data[['employee_name', 'date', 'number_of_punches']])
            
        except Exception as e:
            st.error(f"Lỗi khi xử lý file: {str(e)}")
            st.markdown("Vui lòng kiểm tra lại format file Excel và thử lại.")
    
    else:
        # Hiển thị hướng dẫn khi chưa upload file
        st.markdown("""
        ### 🚀 Bắt đầu sử dụng
        
        Hệ thống này giúp bạn tính toán giờ công từ dữ liệu chấm công Excel một cách tự động và chính xác.
        
        **Các tính năng chính:**
        - ✅ Tự động tính giờ làm việc và giờ OT
        - ✅ Áp dụng quy tắc giờ vào sớm nhất 7h30
        - ✅ Tự động trừ 1 tiếng nghỉ trưa  
        - ✅ Phân biệt giờ thường (8h) và giờ OT
        - ✅ Báo cáo chi tiết theo nhân viên và ngày
        - ✅ Xuất file Excel với nhiều sheet thống kê
        
        **Để bắt đầu, hãy upload file Excel của bạn ở phía trên! 👆**
        """)

if __name__ == "__main__":
    main()
