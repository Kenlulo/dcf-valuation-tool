import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import os
import io
from streamlit_option_menu import option_menu

# ===================== PAGE CONFIG =====================
st.set_page_config(page_title="DCF Valuation Tool", layout="wide", initial_sidebar_state="expanded")

# ===================== CSS THEME =====================
st.markdown("""
<style>
    [data-testid="stAppViewContainer"] { background-color: #E9ECEF; color: #1E1E1E; }
    [data-testid="stSidebar"] { background-color: #0B5C57 !important; border-right: none; }
    [data-testid="stSidebar"] p, [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] label, [data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] { color: #F8F9F9 !important; }
    [data-testid="stHeader"] { background-color: transparent; }
    .block-container { padding-top: 1rem !important; padding-bottom: 1rem !important; max-width: 98% !important; }
    .met-card {
        background: white; padding: 15px 10px; text-align: center;
        border-top: 3px solid #D5D8DC; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border-radius: 4px;
    }
    .met-label { font-size: 13px; font-weight: 600; color: #5D6D7E; text-transform: uppercase; margin-bottom: 5px; }
    .met-val { font-size: 32px; font-weight: 300; line-height: 1.1; margin-bottom: 5px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .met-val.teal { color: #1ABC9C; }
    .met-val.red { color: #E74C3C; }
    .met-val.purple { color: #8E44AD; }
    .met-goal { font-size: 11px; color: #34495E; font-weight: 500; }
    .section-title { font-size: 22px; font-weight: 700; color: #2C3E50; padding: 8px 0; border-bottom: 2px solid #1ABC9C; margin-bottom: 15px; }
    .insight-box { padding: 15px; border-radius: 4px; margin-top: 10px; }
    .insight-good { background-color: #E8F8F5; border-left: 4px solid #1ABC9C; }
    .insight-bad { background-color: #FDEDEC; border-left: 4px solid #E74C3C; }
</style>
""", unsafe_allow_html=True)

# ===================== LANGUAGE =====================
_LANG = {
    'Tổng quan Doanh nghiệp': 'Company Overview',
    'Tính WACC (Chi phí Vốn)': 'WACC Calculator',
    'Mô hình DCF (5 năm)': 'DCF Model (5 Years)',
    'Phân tích Độ nhạy': 'Sensitivity Analysis',
    'Kết luận & Xuất Excel': 'Conclusion & Export',
    'Doanh thu': 'Revenue',
    'Giá vốn Hàng bán': 'Cost of Sales',
    'Lợi nhuận Gộp': 'Gross Profit',
    'Chi phí Bán hàng & QLDN': 'SG&A Expenses',
    'Lợi nhuận Hoạt động (EBIT)': 'Operating Profit (EBIT)',
    'Lãi vay': 'Interest Expenses',
    'Lợi nhuận trước thuế': 'Profit Before Tax',
    'Thuế TNDN': 'Corporate Income Tax',
    'Lợi nhuận Ròng': 'Net Profit',
    'Khấu hao': 'Depreciation & Amortization',
    'Tổng Tài sản': 'Total Assets', 'Tổng Nợ': 'Total Liabilities',
    'Nợ ngắn hạn': 'Current Liabilities', 'Nợ dài hạn': 'Long-term Liabilities',
    'Vốn Chủ sở hữu': 'Equity', 'Tiền mặt': 'Cash',
    'Khoản phải thu': 'Receivables', 'Hàng tồn kho': 'Inventory',
    'Khoản phải trả': 'Payables', 'Tài sản Cố định': 'Fixed Assets',
    'Cấu trúc Vốn (Capital Structure)': 'Capital Structure',
    'Vốn chủ': 'Equity', 'Nợ vay': 'Debt',
    'Chỉ số Tài chính': 'Financial Ratios',
    'Lãi suất phi rủi ro (Rf)': 'Risk-free Rate (Rf)',
    'Hệ số Beta (β)': 'Beta (β)',
    'Phần bù rủi ro thị trường (Rm - Rf)': 'Market Risk Premium (Rm - Rf)',
    'Chi phí Vốn chủ (Ke)': 'Cost of Equity (Ke)',
    'Chi phí Nợ vay trước thuế (Kd)': 'Pre-tax Cost of Debt (Kd)',
    'Thuế suất (%)': 'Tax Rate (%)',
    'Tỷ trọng Vốn chủ (E/V)': 'Equity Weight (E/V)',
    'Tỷ trọng Nợ (D/V)': 'Debt Weight (D/V)',
    'Tăng trưởng Doanh thu (%)': 'Revenue Growth (%)',
    'Biên EBIT (%)': 'EBIT Margin (%)',
    'CAPEX / Doanh thu (%)': 'CAPEX / Revenue (%)',
    'D&A / Doanh thu (%)': 'D&A / Revenue (%)',
    'ΔWC / Doanh thu (%)': 'ΔWC / Revenue (%)',
    'Tăng trưởng dài hạn (g)': 'Terminal Growth (g)',
    'Giá trị Doanh nghiệp (EV)': 'Enterprise Value (EV)',
    'Giá trị Vốn chủ sở hữu': 'Equity Value',
    'Giá trị hợp lý / Cổ phiếu': 'Fair Value / Share',
    'Tên Doanh nghiệp': 'Company Name',
    'Số cổ phiếu lưu hành (triệu)': 'Shares Outstanding (million)',
    'Giá thị trường / cổ phiếu (VND)': 'Market Price / Share (VND)',
    'Ngành nghề': 'Industry',
    'Năm': 'Year',
    'Chọn nguồn dữ liệu': 'Select Data Source',
    'Upload file Excel (.xlsx)': 'Upload Excel file (.xlsx)',
    'Dùng dữ liệu mẫu (Demo)': 'Use Sample Data (Demo)',
    '📥 Tải Template Excel mẫu': '📥 Download Excel Template',
    '📤 Upload BCTC của bạn': '📤 Upload Your Financial Statements',
    'Chưa có dữ liệu. Vui lòng upload file Excel hoặc chọn dữ liệu mẫu.': 'No data loaded. Please upload an Excel file or select sample data.',
    'Lãi suất phi rủi ro Rf (%)': 'Risk-free Rate Rf (%)',
    'Phần bù rủi ro TT Rm-Rf (%)': 'Market Risk Premium Rm-Rf (%)',
    'Kd trước thuế (%)': 'Pre-tax Kd (%)',
    'Tỷ trọng Vốn chủ E/V (%)': 'Equity Weight E/V (%)',
    'Tỷ trọng Nợ D/V (%)': 'Debt Weight D/V (%)',
    'Tăng trưởng dài hạn g (%)': 'Terminal Growth g (%)',
    'Phương pháp Terminal Value': 'Terminal Value Method',
    'Cấu trúc Giá trị Doanh nghiệp (EV Composition)': 'Enterprise Value Composition',
    'Cập nhật Dữ liệu (Upload Mode): Toàn bộ dữ liệu tài chính được nhập trực tiếp từ file Excel do người dùng cung cấp. Dữ liệu Demo (31 mã VN30) được chốt cố định tại thời điểm tải snapshot nhằm phục vụ mục đích minh họa.': 'Data Update (Upload Mode): All financial data is imported directly from the user-provided Excel file. Demo data (31 VN30 tickers) is fixed at snapshot download time for illustration purposes.',
    'DỰ ÁN CÁ NHÂN (PORTFOLIO PROJECT): Công cụ DCF Valuation này là một dự án Mã nguồn mở mang tính chất Học thuật & Giáo dục (Data Science & Financial Modeling Portfolio). Tuyên bố miễn trừ trách nhiệm: Các kết quả định giá trên công cụ này chỉ mang tính chất tham khảo, mô phỏng học thuật và không phải là lời khuyên đầu tư tài chính. Nguồn dữ liệu thô: Vnstock.': 'PORTFOLIO PROJECT: This DCF Valuation Tool is an open-source project for Academic & Educational purposes (Data Science & Financial Modeling Portfolio). Disclaimer: Valuation results from this tool are for reference and academic simulation only, and do not constitute financial investment advice. Raw data source: Vnstock.',
    'Nếu bạn muốn xem thêm dự án khác hãy': 'If you want to view more projects, please',
    'nhấp vào đây': 'click here',
    'Phát triển bởi': 'Developed by',
    # Sidebar
    'Chọn mã:': 'Select ticker:',
    'Không tìm thấy thư mục data_snapshot/': 'data_snapshot/ directory not found.',
    'Bước 1: Tải Template': 'Step 1: Download Template',
    'Tải file mẫu, điền BCTC vào 3 sheets': 'Download template, fill in 3 sheets',
    'Bước 2: Upload File': 'Step 2: Upload File',
    'Chọn file Excel đã điền BCTC': 'Select your filled Excel file',
    # WACC Explanation
    'Các chỉ số trên nói lên điều gì?': 'What do these metrics mean?',
    'thấp': 'low',
    'trung bình': 'moderate',
    'cao': 'high',
    # Terminal Value Warnings
    'Terminal Value chiếm': 'Terminal Value accounts for',
    'tổng EV': 'of total EV',
    'Khuyến nghị:': 'Recommendation:',
    # Sensitivity Explanation
    'Bảng Độ nhạy nói lên điều gì?': 'What does the Sensitivity Table tell us?',
    'Cách đọc bảng:': 'How to read:',
    'Giá trị hợp lý / Cổ phiếu (VND)': 'Fair Value / Share (VND)',
    'Biên độ dao động:': 'Value range:',
    'Ý nghĩa thực tiễn:': 'Practical implications:',
    # Conclusion
    'Tải báo cáo DCF': 'Download DCF Report',
    'Giá thị trường': 'Market Price',
    'Định giá mọi Doanh nghiệp': 'Valuation for Any Business',
    # Terminal Value warnings
    'Mô hình **cực kỳ nhạy cảm** với giả định dài hạn': 'Model is **extremely sensitive** to long-term assumptions',
    'Thay đổi nhỏ trong g hoặc WACC sẽ thay đổi kết quả rất lớn.': 'Small changes in g or WACC will significantly alter results.',
    'Kiểm tra lại giả định hoặc sử dụng phương pháp Exit Multiple để cross-check.': 'Review assumptions or use Exit Multiple method to cross-check.',
    'Đây là mức bình thường trong DCF (thường 60-75%). Tuy nhiên, nhà đầu tư nên chạy Sensitivity Analysis để kiểm tra độ nhạy.': 'This is normal for DCF (typically 60-75%). However, investors should run Sensitivity Analysis to verify.',
    'Mô hình có nền tảng vững — phần lớn giá trị đến từ dòng tiền dự phóng 5 năm, không phụ thuộc quá nhiều vào giả định dài hạn.': 'Model has solid foundations — most value comes from 5-year projected cash flows, not overly dependent on long-term assumptions.',
    # Table headers
    'Thành phần': 'Component',
    'Giá trị': 'Value',
    'Tỷ trọng': 'Weight',
    'Giá trị Vốn chủ sở hữu': 'Equity Value',
}

def t(text):
    if st.session_state.get('lang', '🇻🇳 Tiếng Việt') == '🇬🇧 English':
        return _LANG.get(text, text)
    return text

# ===================== TEMPLATE GENERATOR =====================
def generate_template():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book
        
        # Formats
        header_fmt = wb.add_format({'bold': True, 'bg_color': '#2C3E50', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        item_fmt = wb.add_format({'bold': True, 'bg_color': '#EAEDED', 'border': 1})
        vi_fmt = wb.add_format({'italic': True, 'font_color': '#5D6D7E', 'border': 1})
        num_fmt = wb.add_format({'border': 1, 'num_format': '#,##0'})
        title_fmt = wb.add_format({'bold': True, 'font_size': 14, 'font_color': '#2980B9'})
        text_fmt = wb.add_format({'text_wrap': True, 'valign': 'top'})
        
        # 1. Hướng dẫn (Instructions)
        ws_inst = wb.add_worksheet('Hướng dẫn')
        ws_inst.write('A1', 'HƯỚNG DẪN NHẬP DỮ LIỆU ĐỊNH GIÁ DCF', title_fmt)
        instructions = [
            "1. Nhập Số liệu vào các ô có nền trắng. KHÔNG SỬA cột 'Item' (Cột A) vì phần mềm dùng tên này để đọc báo cáo.",
            "2. Đơn vị tiền tệ: Tỷ VNĐ (Billion VND). Ví dụ: 15,000 tỷ thì nhập 15000.",
            "3. Số lượng cổ phiếu: Tính bằng Triệu cổ phiếu. Ví dụ: 1 tỷ cổ phiếu thì nhập 1000.",
            "4. Giá cổ phiếu: Tính bằng VNĐ thực tế. Ví dụ: 50.000 VNĐ thì nhập 50000.",
            "5. Năm (Year): Bạn có thể đổi tên các cột năm ở Dòng 1 (VD: '2021', '2022', 'LTM') tùy ý.",
            "6. Các khoản chi phí (Giá vốn, Chi phí QLDN, Lãi vay, Thuế): Nhập SỐ DƯƠNG (Phần mềm tự động hiểu là chi phí để trừ ra)."
        ]
        for i, text in enumerate(instructions):
            ws_inst.write(i+2, 0, text, text_fmt)
        ws_inst.set_column('A:A', 100)
        
        # 2. Sheet: Info
        info_items = ['Company Name', 'Industry', 'Shares Outstanding (million)', 'Market Price per Share (VND)']
        info_vi = ['Tên doanh nghiệp', 'Ngành nghề', 'Số lượng cổ phiếu lưu hành (Triệu cp)', 'Giá cắt lỗ/thị trường (VNĐ)']
        info_df = pd.DataFrame({'Item': info_items, 'Vietnamese (Giải thích)': info_vi, 'Value': ['ABC Corp', 'Manufacturing', 100, 50000]})
        info_df.to_excel(writer, sheet_name='Info', index=False)
        ws_info = writer.sheets['Info']
        ws_info.set_column('A:A', 30, item_fmt)
        ws_info.set_column('B:B', 40, vi_fmt)
        ws_info.set_column('C:C', 20, num_fmt)
        for col_num, value in enumerate(info_df.columns.values):
            ws_info.write(0, col_num, value, header_fmt)
            
        # 3. Sheet: IncomeStatement
        years = [2021, 2022, 2023, 2024]
        is_items = ['Revenue', 'Cost of Sales', 'Gross Profit', 'SG&A Expenses', 'Depreciation & Amortization',
                    'Operating Profit (EBIT)', 'Interest Expenses', 'Profit Before Tax', 'Corporate Income Tax', 'Net Profit']
        is_vi = ['Doanh thu thuần', 'Giá vốn hàng bán', 'Lợi nhuận gộp', 'Chi phí BH & QLDN', 'Chi phí Khấu hao',
                 'Lợi nhuận từ HĐKD (EBIT)', 'Chi phí lãi vay', 'Lợi nhuận trước thuế', 'Thuế TNDN', 'Lợi nhuận ròng']
        is_vals = [[10000, 6000, 4000, 1500, 500, 2000, 200, 1800, 360, 1440],
                   [12000, 7200, 4800, 1800, 550, 2450, 250, 2200, 440, 1760],
                   [14000, 8400, 5600, 2100, 600, 2900, 300, 2600, 520, 2080],
                   [16000, 9600, 6400, 2400, 650, 3350, 350, 3000, 600, 2400]]
        
        is_df = pd.DataFrame({'Item': is_items, 'Vietnamese (Giải thích)': is_vi})
        for i, y in enumerate(years):
            is_df[str(y)] = is_vals[i]
            
        is_df.to_excel(writer, sheet_name='IncomeStatement', index=False)
        ws_is = writer.sheets['IncomeStatement']
        ws_is.set_column('A:A', 30, item_fmt)
        ws_is.set_column('B:B', 30, vi_fmt)
        ws_is.set_column('C:F', 18, num_fmt)
        for col_num, value in enumerate(is_df.columns.values):
            ws_is.write(0, col_num, value, header_fmt)
            
        # 4. Sheet: BalanceSheet
        bs_items = ['Cash', 'Receivables', 'Inventory', 'Other Current Assets', 'Fixed Assets', 'Other Long-term Assets',
                    'Total Assets', 'Short-term Borrowings', 'Payables', 'Other Current Liabilities',
                    'Long-term Borrowings', 'Other Long-term Liabilities', 'Total Liabilities', 'Equity']
        bs_vi = ['Tiền & Tương đương tiền', 'Phải thu ngắn hạn', 'Hàng tồn kho', 'Tài sản ngắn hạn khác', 'Tài sản cố định', 'Tài sản dài hạn khác',
                 'TỔNG TÀI SẢN', 'Vay nợ ngắn hạn', 'Phải trả người bán', 'Nợ ngắn hạn khác',
                 'Vay nợ dài hạn', 'Nợ dài hạn khác', 'TỔNG NỢ PHẢI TRẢ', 'VỐN CHỦ SỞ HỮU']
        bs_vals_base = [2000, 1500, 1000, 500, 8000, 2000, 15000, 2000, 1200, 800, 3000, 500, 7500, 7500]
        
        bs_df = pd.DataFrame({'Item': bs_items, 'Vietnamese (Giải thích)': bs_vi})
        for y in years:
            bs_df[str(y)] = [v * (1 + 0.1*(y-2021)) for v in bs_vals_base]
            
        bs_df.to_excel(writer, sheet_name='BalanceSheet', index=False)
        ws_bs = writer.sheets['BalanceSheet']
        ws_bs.set_column('A:A', 30, item_fmt)
        ws_bs.set_column('B:B', 30, vi_fmt)
        ws_bs.set_column('C:F', 18, num_fmt)
        for col_num, value in enumerate(bs_df.columns.values):
            ws_bs.write(0, col_num, value, header_fmt)
        
    output.seek(0)
    return output

# ===================== DATA PARSER =====================
def parse_uploaded_excel(uploaded_file):
    """Parse uploaded Excel into standardized dict"""
    try:
        xl = pd.ExcelFile(uploaded_file)
        data = {}
        
        # Info
        if 'Info' in xl.sheet_names:
            info_df = pd.read_excel(xl, 'Info')
            if 'Value' in info_df.columns:
                info = dict(zip(info_df.iloc[:, 0], info_df['Value']))
            else:
                info = dict(zip(info_df.iloc[:, 0], info_df.iloc[:, -1]))
            data['company_name'] = str(info.get('Company Name', 'Unknown'))
            data['industry'] = str(info.get('Industry', 'N/A'))
            data['shares'] = float(info.get('Shares Outstanding (million)', 0))
            data['market_price'] = float(info.get('Market Price per Share (VND)', 0))
        else:
            data['company_name'] = 'Unknown'
            data['industry'] = 'N/A'
            data['shares'] = 0
            data['market_price'] = 0
        
        # Income Statement
        if 'IncomeStatement' in xl.sheet_names:
            is_df = pd.read_excel(xl, 'IncomeStatement')
            is_df = is_df.set_index(is_df.columns[0])
            is_df.columns = [str(c) for c in is_df.columns]
            if 'Vietnamese (Giải thích)' in is_df.columns:
                is_df = is_df.drop(columns=['Vietnamese (Giải thích)'])
            data['is_df'] = is_df
            data['years'] = list(is_df.columns)
        
        # Balance Sheet
        if 'BalanceSheet' in xl.sheet_names:
            bs_df = pd.read_excel(xl, 'BalanceSheet')
            bs_df = bs_df.set_index(bs_df.columns[0])
            bs_df.columns = [str(c) for c in bs_df.columns]
            if 'Vietnamese (Giải thích)' in bs_df.columns:
                bs_df = bs_df.drop(columns=['Vietnamese (Giải thích)'])
            data['bs_df'] = bs_df
        
        data['valid'] = True
        return data
    except Exception as e:
        return {'valid': False, 'error': str(e)}

def parse_snapshot(ticker):
    """Parse data_snapshot file into standardized dict"""
    path = f"data_snapshot/{ticker}_snapshot.xlsx"
    if not os.path.exists(path):
        return {'valid': False, 'error': f'File {path} not found'}
    
    try:
        xl = pd.ExcelFile(path)
        data = {'company_name': ticker, 'industry': 'Listed (HOSE/HNX)', 'valid': True}
        
        # Income Statement — chỉ lấy full year (lengthReport == 4)
        raw_is = pd.read_excel(xl, 'IncomeStatement')
        raw_is = raw_is[raw_is['lengthReport'] == 4]  # Full year only
        years_raw = raw_is['yearReport'].unique()
        years = sorted([int(y) for y in years_raw if pd.notna(y)])[-4:]
        data['years'] = [str(y) for y in years]
        
        def get_is_val(col_keywords, year):
            for col in raw_is.columns:
                for kw in col_keywords:
                    if kw.lower() in col.lower():
                        row = raw_is[raw_is['yearReport'] == year]
                        if len(row) > 0:
                            val = row[col].values[0]
                            return float(val) / 1e9 if pd.notna(val) else 0  # VND → Bn VND
            return 0
        
        is_items = ['Revenue', 'Cost of Sales', 'Gross Profit', 'SG&A Expenses', 'Depreciation & Amortization',
                    'Operating Profit (EBIT)', 'Interest Expenses', 'Profit Before Tax', 'Corporate Income Tax', 'Net Profit']
        is_keys = [
            ['Net Sales', 'Revenue (Bn'],
            ['Cost of Sales'],
            ['Gross Profit'],
            ['Selling Expenses', 'General & Admin'],
            ['Depreciation'],
            ['Operating Profit'],
            ['Interest Expenses'],
            ['Profit before tax'],
            ['Business income tax'],
            ['Net Profit For the Year', 'Attributable to parent']
        ]
        
        is_dict = {'Item': is_items}
        for y in years:
            vals = []
            for keys in is_keys:
                v = get_is_val(keys, y)
                vals.append(v)
            # Fix: if COGS missing, calc from revenue - gross
            if vals[1] == 0 and vals[0] != 0 and vals[2] != 0:
                vals[1] = vals[0] - vals[2]
            # Fix: SGA = sum selling + admin
            if vals[3] == 0:
                vals[3] = get_is_val(['Selling Expenses'], y) + get_is_val(['General & Admin'], y)
            # Fix: D&A estimate
            if vals[4] == 0:
                vals[4] = abs(vals[0]) * 0.03
            is_dict[str(y)] = vals
        
        is_df = pd.DataFrame(is_dict).set_index('Item')
        data['is_df'] = is_df
        
        # Balance Sheet — chỉ lấy full year
        raw_bs = pd.read_excel(xl, 'BalanceSheet')
        raw_bs = raw_bs[raw_bs['lengthReport'] == 4]  # Full year only
        
        def get_bs_val(col_keywords, year):
            for col in raw_bs.columns:
                for kw in col_keywords:
                    if kw.lower() in col.lower():
                        row = raw_bs[raw_bs['yearReport'] == year]
                        if len(row) > 0:
                            val = row[col].values[0]
                            return float(val) / 1e9 if pd.notna(val) else 0  # VND → Bn VND
            return 0
        
        bs_items = ['Cash', 'Receivables', 'Inventory', 'Other Current Assets', 'Fixed Assets', 'Other Long-term Assets',
                    'Total Assets', 'Short-term Borrowings', 'Payables', 'Other Current Liabilities',
                    'Long-term Borrowings', 'Other Long-term Liabilities', 'Total Liabilities', 'Equity']
        bs_keys = [
            ['Cash and cash equiv'],
            ['Accounts receivable'],
            ['Inventories, Net', 'Net Inventor'],
            ['Other current assets (Bn'],
            ['Fixed assets (Bn'],
            ['Other non-current'],
            ['TOTAL ASSETS (Bn'],
            ['Short-term borrowings (Bn'],
            ['Advances from customers', 'Prepayments to suppliers'],
            ['Other current'],
            ['Long-term borrowings (Bn'],
            ['Other long-term'],
            ['LIABILITIES (Bn'],
            ["OWNER'S EQUITY", 'Capital and reserves']
        ]
        
        bs_dict = {'Item': bs_items}
        for y in years:
            vals = []
            for keys in bs_keys:
                vals.append(get_bs_val(keys, y))
            bs_dict[str(y)] = vals
        
        bs_df = pd.DataFrame(bs_dict).set_index('Item')
        data['bs_df'] = bs_df
        
        # Ratios for shares/price + HISTORICAL data
        raw_rt = pd.read_excel(xl, 'Ratios')
        raw_rt = raw_rt[raw_rt['lengthReport'] == 4].sort_values('yearReport')
        latest = raw_rt[raw_rt['yearReport'] == max(years)]
        data['shares'] = float(latest['Outstanding Share (Mil. Shares)'].values[0]) / 1e6 if 'Outstanding Share (Mil. Shares)' in raw_rt.columns and len(latest) > 0 else 0  # actual shares → millions
        data['market_price'] = 0
        if 'Market Capital (Bn. VND)' in raw_rt.columns and data['shares'] > 0 and len(latest) > 0:
            mkt_cap_vnd = float(latest['Market Capital (Bn. VND)'].values[0])  # actually VND
            data['market_price'] = mkt_cap_vnd / (data['shares'] * 1e6)  # VND / shares = VND/share
        
        # Historical trend: last 5 years from Ratios
        hist_years = sorted(raw_rt['yearReport'].unique())[-5:]
        has_ebit_margin = 'EBIT Margin (%)' in raw_rt.columns
        hist_data = []
        for hy in hist_years:
            hy_row = raw_rt[raw_rt['yearReport'] == hy].iloc[0]
            ebitda_h = float(hy_row.get('EBITDA (Bn. VND)', 0)) / 1e9 if pd.notna(hy_row.get('EBITDA (Bn. VND)', None)) else 0
            ebit_h = float(hy_row.get('EBIT (Bn. VND)', 0)) / 1e9 if pd.notna(hy_row.get('EBIT (Bn. VND)', None)) else 0
            npm_h = float(hy_row.get('Net Profit Margin (%)', 0)) if pd.notna(hy_row.get('Net Profit Margin (%)', None)) else 0
            roe_h = float(hy_row.get('ROE (%)', 0)) if pd.notna(hy_row.get('ROE (%)', None)) else 0
            
            rev_h = 0
            np_h = 0
            
            if has_ebit_margin:
                # Non-bank: Revenue = EBIT / EBIT_Margin
                ebit_margin_h = float(hy_row.get('EBIT Margin (%)', 0)) if pd.notna(hy_row.get('EBIT Margin (%)', None)) else 0
                rev_h = (ebit_h / ebit_margin_h) if ebit_margin_h > 0 else 0
                np_h = rev_h * npm_h if npm_h > 0 and rev_h > 0 else 0
            
            # Fallback: P/S method for any ticker where EBIT method fails (negative margins, banks)
            if rev_h == 0:
                mktcap_h = float(hy_row.get('Market Capital (Bn. VND)', 0)) / 1e9 if pd.notna(hy_row.get('Market Capital (Bn. VND)', None)) else 0
                ps_h = float(hy_row.get('P/S', 0)) if pd.notna(hy_row.get('P/S', None)) else 0
                eps_h = float(hy_row.get('EPS (VND)', 0)) if pd.notna(hy_row.get('EPS (VND)', None)) else 0
                shares_h = float(hy_row.get('Outstanding Share (Mil. Shares)', 0)) if pd.notna(hy_row.get('Outstanding Share (Mil. Shares)', None)) else 0
                rev_h = (mktcap_h / ps_h) if ps_h > 0 else 0
                np_h = eps_h * shares_h / 1e9 if shares_h > 0 and np_h == 0 else np_h
            de_h = float(hy_row.get('Debt/Equity', 0)) if pd.notna(hy_row.get('Debt/Equity', None)) else 0
            
            hist_data.append({
                'year': str(int(hy)), 'revenue': rev_h, 'net_profit': np_h,
                'ebitda': ebitda_h, 'ebit': ebit_h, 'roe': roe_h * 100,
                'de_ratio': de_h
            })
        data['hist_trend'] = hist_data
        data['has_ebit_margin'] = has_ebit_margin
        
        return data
    except Exception as e:
        return {'valid': False, 'error': str(e)}

def safe_get(df, row_name, col, default=0):
    try:
        if row_name in df.index and col in df.columns:
            v = df.loc[row_name, col]
            return float(v) if pd.notna(v) else default
        return default
    except:
        return default

# ===================== SIDEBAR =====================
with st.sidebar:
    lang = st.radio("Ngôn ngữ / Language:", ['🇻🇳 Tiếng Việt', '🇬🇧 English'], index=0, key='lang')
    st.markdown(f"### 📐 {t('DCF Valuation Tool')}")
    
    selected_tab = option_menu(
        menu_title=None,
        options=[t("Tổng quan Doanh nghiệp"), t("Tính WACC (Chi phí Vốn)"), t("Mô hình DCF (5 năm)"), t("Phân tích Độ nhạy"), t("Kết luận & Xuất Excel")],
        icons=["building", "calculator", "graph-up-arrow", "grid-3x3", "file-earmark-excel"],
        default_index=0,
        styles={
            "container": {"background-color": "transparent"},
            "icon": {"color": "#1ABC9C", "font-size": "16px"},
            "nav-link": {"color": "#ABB2B9", "font-size": "14px", "text-align": "left", "margin": "2px 0"},
            "nav-link-selected": {"background-color": "#1ABC9C", "color": "white", "font-weight": "600"},
        }
    )
    
    st.markdown("---")
    st.markdown(f"### {t('Chọn nguồn dữ liệu')}")
    
    data_source = st.radio("", [t("Upload file Excel (.xlsx)"), t("Dùng dữ liệu mẫu (Demo)")], index=0, label_visibility="collapsed")
    
    data = None
    
    if data_source == t("Upload file Excel (.xlsx)"):
        # Step 1: Download template
        st.markdown(f"""
<div style="background:#0E6655; border-radius:8px; padding:10px 12px; margin-bottom:6px;">
  <div style="font-size:11px; color:#A3E4D7; font-weight:600; text-transform:uppercase; letter-spacing:1px; margin-bottom:4px;">
    📋 {t("Bước 1: Tải Template")}
  </div>
  <div style="font-size:11px; color:#D1F2EB; margin-bottom:6px;">
    {t("Tải file mẫu, điền BCTC vào 3 sheets")}
  </div>
</div>""", unsafe_allow_html=True)
        st.download_button(
            label=f"📥 {t('Tải Template Excel mẫu')}",
            data=generate_template(),
            file_name="DCF_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        # Step 2: Upload file
        st.markdown(f"""
<div style="background:#0E6655; border-radius:8px; padding:10px 12px; margin-top:8px; margin-bottom:4px;">
  <div style="font-size:11px; color:#A3E4D7; font-weight:600; text-transform:uppercase; letter-spacing:1px; margin-bottom:4px;">
    📤 {t("Bước 2: Upload File")}
  </div>
  <div style="font-size:11px; color:#D1F2EB;">
    {t("Chọn file Excel đã điền BCTC")}
  </div>
</div>""", unsafe_allow_html=True)
        uploaded = st.file_uploader("", type=['xlsx'], label_visibility="collapsed")
        if uploaded:
            data = parse_uploaded_excel(uploaded)
    else:
        snapshot_dir = "data_snapshot"
        if os.path.exists(snapshot_dir):
            tickers = sorted([f.replace('_snapshot.xlsx', '') for f in os.listdir(snapshot_dir) if f.endswith('_snapshot.xlsx') and f != 'VNINDEX_snapshot.xlsx'])
            default_idx = tickers.index('REE') if 'REE' in tickers else 0
            selected_ticker = st.selectbox(t("Chọn mã:"), tickers, index=default_idx)
            data = parse_snapshot(selected_ticker)
        else:
            st.warning(t("Không tìm thấy thư mục data_snapshot/"))

# ===================== MAIN CONTENT =====================
st.markdown(f"""
<div style="text-align: center; padding: 5px 0;">
    <span style="font-size: 28px; font-weight: 700; color: #2C3E50;">📐 DCF Valuation Tool</span><br>
    <span style="font-size: 14px; color: #7F8C8D;">Discounted Cash Flow — {t("Định giá mọi Doanh nghiệp")}</span>
</div>
""", unsafe_allow_html=True)

# ===================== FIXED HEADER NOTICES =====================
st.info(t('Cập nhật Dữ liệu (Upload Mode): Toàn bộ dữ liệu tài chính được nhập trực tiếp từ file Excel do người dùng cung cấp. Dữ liệu Demo (31 mã VN30) được chốt cố định tại thời điểm tải snapshot nhằm phục vụ mục đích minh họa.'))
st.warning(t('DỰ ÁN CÁ NHÂN (PORTFOLIO PROJECT): Công cụ DCF Valuation này là một dự án Mã nguồn mở mang tính chất Học thuật & Giáo dục (Data Science & Financial Modeling Portfolio). Tuyên bố miễn trừ trách nhiệm: Các kết quả định giá trên công cụ này chỉ mang tính chất tham khảo, mô phỏng học thuật và không phải là lời khuyên đầu tư tài chính. Nguồn dữ liệu thô: Vnstock.'))
st.success(f"👉 **{t('Nếu bạn muốn xem thêm dự án khác hãy')} [{t('nhấp vào đây')}](https://portfolio-gilt-sigma-43.vercel.app)**")

if data is None or not data.get('valid', False):
    st.warning(t("Chưa có dữ liệu. Vui lòng upload file Excel hoặc chọn dữ liệu mẫu."))
    if data and not data.get('valid', False):
        st.error(f"Error: {data.get('error', 'Unknown')}")
    st.stop()

# Extract core data
is_df = data.get('is_df', pd.DataFrame())
bs_df = data.get('bs_df', pd.DataFrame())
years = data.get('years', [])
company_name = data.get('company_name', 'Unknown')
shares_mil = data.get('shares', 0)
market_price = data.get('market_price', 0)
latest_year = years[-1] if years else 'N/A'

# Helper to get latest values
def lv(df, item):
    return safe_get(df, item, latest_year)

rev = lv(is_df, 'Revenue')
cogs = abs(lv(is_df, 'Cost of Sales'))
gross = lv(is_df, 'Gross Profit')
sga = abs(lv(is_df, 'SG&A Expenses'))
ebit = lv(is_df, 'Operating Profit (EBIT)')
interest = abs(lv(is_df, 'Interest Expenses'))
pbt = lv(is_df, 'Profit Before Tax')
tax = abs(lv(is_df, 'Corporate Income Tax'))
net_profit = lv(is_df, 'Net Profit')
da = abs(lv(is_df, 'Depreciation & Amortization'))

cash = lv(bs_df, 'Cash')
total_assets = lv(bs_df, 'Total Assets')
st_debt = abs(lv(bs_df, 'Short-term Borrowings'))
lt_debt = abs(lv(bs_df, 'Long-term Borrowings'))
total_debt = st_debt + lt_debt
total_liab = abs(lv(bs_df, 'Total Liabilities'))
equity = lv(bs_df, 'Equity')
receivables = lv(bs_df, 'Receivables')
inventory = lv(bs_df, 'Inventory')
payables = abs(lv(bs_df, 'Payables'))
fixed_assets = lv(bs_df, 'Fixed Assets')

# Banks: use total_liab when borrowings = 0 but liabilities exist
if total_debt == 0 and total_liab > 0:
    total_debt = total_liab

if ebit == 0 and rev > 0:
    ebit = gross - sga
if gross == 0 and rev > 0:
    gross = rev - cogs

# ── Pre-compute default DCF so Tab 5 always has a value ──────────────────
# Uses default assumptions: rev_g=10%, ebit_m from data, capex=5%, da=3%, dwc=1%, g=3%
_wacc_default = st.session_state.get('wacc', 0.10)
_tax_default  = st.session_state.get('tax_rate', 0.20)
_ticker_key   = company_name  # use company name as cache key
_cached_ticker = st.session_state.get('_dcf_ticker', None)

# Recompute default whenever ticker changes or no value yet
if _cached_ticker != _ticker_key or st.session_state.get('ev', 0) == 0:
    _rev_g  = 0.10
    _hist = data.get('hist_trend', [])
    _hist_rev = _hist[-1].get('revenue', 0) if _hist else 0
    _base_rev = rev if rev > 0 else (_hist_rev if _hist_rev > 0 else 1000)
    _ebit_m = (ebit / rev) if rev > 0 and ebit > 0 else 0.15
    _da_r   = (da / rev)   if rev > 0 and da > 0   else 0.03
    _capex_r = 0.05
    _dwc_r   = 0.01
    _g_term  = 0.03
    _fcffs = []
    _pv_fcffs = []
    for _i in range(5):
        _p_rev   = _base_rev * ((1 + _rev_g) ** (_i + 1))
        _p_nopat = _p_rev * _ebit_m * (1 - _tax_default)
        _p_da    = _p_rev * _da_r
        _p_capex = _p_rev * _capex_r
        _p_dwc   = _p_rev * _dwc_r
        _p_fcff  = _p_nopat + _p_da - _p_capex - _p_dwc
        _pv      = _p_fcff / ((1 + _wacc_default) ** (_i + 1))
        _fcffs.append(_p_fcff)
        _pv_fcffs.append(_pv)
    _tv = _fcffs[-1] * (1 + _g_term) / (_wacc_default - _g_term) if _wacc_default > _g_term else 0
    _pv_tv = _tv / ((1 + _wacc_default) ** 5)
    _ev = sum(_pv_fcffs) + _pv_tv
    _net_debt = total_debt - cash
    _equity_value = _ev - _net_debt
    _fv_ps = (_equity_value * 1e9) / (shares_mil * 1e6) if shares_mil > 0 else 0
    st.session_state['ev']           = _ev
    st.session_state['equity_value'] = _equity_value
    st.session_state['fair_value_ps'] = _fv_ps
    st.session_state['_dcf_ticker']  = _ticker_key


if selected_tab == t("Tổng quan Doanh nghiệp"):
    st.markdown(f'<div class="section-title">🏢 {company_name} — {t("Tổng quan Doanh nghiệp")}</div>', unsafe_allow_html=True)
    
    c1, c2, c3, c4, c5 = st.columns(5)
    for col, label, val, cls in [
        (c1, t("Doanh thu"), rev, "teal"), (c2, "EBITDA", ebit + da, "teal"),
        (c3, t("Lợi nhuận Ròng"), net_profit, "teal" if net_profit >= 0 else "red"),
        (c4, t("Tổng Nợ"), total_debt, "red"), (c5, t("Vốn Chủ sở hữu"), equity, "purple")
    ]:
        col.markdown(f'<div class="met-card"><div class="met-label">{label}</div><div class="met-val {cls}">{val:,.0f}</div><div class="met-goal">Bn VND ({latest_year})</div></div>', unsafe_allow_html=True)
    
    st.markdown("---")
    col_chart, col_pie = st.columns([3, 2])
    
    with col_chart:
        st.markdown(f"### 📈 {t('Doanh thu')} & {t('Lợi nhuận Ròng')}")
        fig_trend = go.Figure()
        
        # Use historical data from Ratios if available (multi-year)
        hist_trend = data.get('hist_trend', [])
        if len(hist_trend) > 1:
            chart_years = [h['year'] for h in hist_trend]
            chart_rev = [h['revenue'] for h in hist_trend]
            chart_np = [h['net_profit'] for h in hist_trend]
        else:
            chart_years = years
            chart_rev = [safe_get(is_df, 'Revenue', y) for y in years]
            chart_np = [safe_get(is_df, 'Net Profit', y) for y in years]
        
        bar_width = 0.4 if len(chart_years) <= 2 else None
        fig_trend.add_trace(go.Bar(x=chart_years, y=chart_rev, name=t("Doanh thu"), marker_color='#1ABC9C', width=bar_width))
        fig_trend.add_trace(go.Scatter(x=chart_years, y=chart_np, name=t("Lợi nhuận Ròng"), mode='lines+markers', line=dict(color='#E74C3C', width=3), marker=dict(size=8), yaxis='y2'))
        fig_trend.update_layout(height=350, margin=dict(t=10, b=10, l=10, r=10), plot_bgcolor='rgba(0,0,0,0)',
                                xaxis=dict(type='category'),
                                yaxis2=dict(overlaying='y', side='right', showgrid=False),
                                legend=dict(orientation="h", y=1.1, x=0.5, xanchor='center'))
        st.plotly_chart(fig_trend, use_container_width=True)
    
    with col_pie:
        st.markdown(f"### 🏗️ {t('Cấu trúc Vốn (Capital Structure)')}")
        # Banks: no short/long-term borrowing breakdown → show total liab vs equity
        if st_debt == 0 and lt_debt == 0 and total_liab > 0:
            pie_labels = [t('Vốn chủ'), t('Tổng Nợ')]
            pie_values = [max(equity, 0), max(total_liab, 0)]
            pie_colors = ['#1ABC9C', '#E74C3C']
        else:
            pie_labels = [t('Vốn chủ'), t('Nợ ngắn hạn'), t('Nợ dài hạn')]
            pie_values = [max(equity, 0), max(st_debt, 0), max(lt_debt, 0)]
            pie_colors = ['#1ABC9C', '#F39C12', '#E74C3C']
        fig_pie = go.Figure(go.Pie(
            labels=pie_labels, values=pie_values,
            marker_colors=pie_colors,
            hole=0.45, textinfo='label+percent'
        ))
        fig_pie.update_layout(height=350, margin=dict(t=10, b=10, l=10, r=10), showlegend=False)
        st.plotly_chart(fig_pie, use_container_width=True)
    
    st.markdown(f"### 📋 {t('Chỉ số Tài chính')}")
    
    # Use Ratios-sheet D/E if available (more accurate for banks)
    hist_trend = data.get('hist_trend', [])
    if hist_trend:
        latest_ht = hist_trend[-1]
        de_from_ratios = latest_ht.get('de_ratio', None)
    else:
        de_from_ratios = None
    
    de_ratio = de_from_ratios if de_from_ratios is not None else (total_debt / equity if equity > 0 else 0)
    roe = (net_profit / equity * 100) if equity > 0 else 0
    roa = (net_profit / total_assets * 100) if total_assets > 0 else 0
    ebit_margin = (ebit / rev * 100) if rev > 0 else 0
    
    # Interest Coverage:
    # - Banks (no EBIT Margin column) → N/A (interest IS their revenue, metric is meaningless)
    # - Non-bank, interest=0 → N/A (no debt, ratio undefined)
    # - Non-bank, ebit<0 → negative, show 0.0x
    # - Otherwise → actual ratio capped at 999x
    _is_bank_ticker = not data.get('has_ebit_margin', True)
    if _is_bank_ticker:
        int_cov_display = "N/A"
    elif interest == 0:
        int_cov_display = "N/A"
    else:
        _int_cov = ebit / interest
        int_cov_display = f"{min(_int_cov, 999):.1f}x" if _int_cov > 0 else f"{_int_cov:.1f}x"
    
    r1, r2, r3, r4, r5 = st.columns(5)
    r1.metric("D/E Ratio", f"{de_ratio:.2f}")
    r2.metric("ROE", f"{roe:.1f}%")
    r3.metric("ROA", f"{roa:.1f}%")
    r4.metric(t("Biên EBIT (%)"), f"{ebit_margin:.1f}%")
    r5.metric("Interest Coverage", int_cov_display)

# ===================== TAB 2: WACC =====================
elif selected_tab == t("Tính WACC (Chi phí Vốn)"):
    st.markdown(f'<div class="section-title">⚖️ {t("Tính WACC (Chi phí Vốn)")}</div>', unsafe_allow_html=True)
    
    st.latex(r"WACC = \frac{E}{V} \times K_e + \frac{D}{V} \times K_d \times (1 - t)")
    
    st.markdown("---")
    col_ke, col_kd = st.columns(2)
    
    with col_ke:
        st.markdown(f"### 📈 {t('Chi phí Vốn chủ (Ke)')} — CAPM")
        st.latex(r"K_e = R_f + \beta \times (R_m - R_f)")
        ke_c1, ke_c2 = st.columns(2)
        rf = ke_c1.number_input(t("Lãi suất phi rủi ro Rf (%)"), min_value=0.0, max_value=15.0, value=2.8, step=0.1, format="%.1f") / 100
        beta = ke_c2.number_input(t("Hệ số Beta (β)"), min_value=0.3, max_value=3.0, value=1.0, step=0.05, format="%.2f")
        mrp = ke_c1.number_input(t("Phần bù rủi ro TT Rm-Rf (%)"), min_value=3.0, max_value=20.0, value=8.0, step=0.5, format="%.1f") / 100
        ke = rf + beta * mrp
        st.markdown(f"**Ke = {rf*100:.1f}% + {beta:.2f} × {mrp*100:.1f}% = <span style='color:#1ABC9C; font-size:24px; font-weight:bold;'>{ke*100:.2f}%</span>**", unsafe_allow_html=True)
    
    with col_kd:
        st.markdown(f"### 💳 {t('Chi phí Nợ vay trước thuế (Kd)')}")
        auto_kd = (interest / total_debt * 100) if total_debt > 0 else 8.0
        auto_kd = auto_kd if 1.0 < auto_kd < 20.0 else 8.0
        kd_c1, kd_c2 = st.columns(2)
        kd = kd_c1.number_input(t("Kd trước thuế (%)"), min_value=0.0, max_value=25.0, value=round(min(auto_kd, 25.0), 1), step=0.1, format="%.1f") / 100
        tax_rate = kd_c2.number_input(t("Thuế suất (%)"), min_value=0.0, max_value=40.0, value=20.0, step=1.0, format="%.0f") / 100
        kd_at = kd * (1 - tax_rate)
        st.markdown(f"**Kd after-tax = {kd*100:.1f}% × (1 - {tax_rate*100:.0f}%) = <span style='color:#E74C3C; font-size:24px; font-weight:bold;'>{kd_at*100:.2f}%</span>**", unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("### ⚖️ WACC")
    
    total_capital = equity + total_debt
    e_weight = equity / total_capital if total_capital > 0 else 0.5
    d_weight = total_debt / total_capital if total_capital > 0 else 0.5
    
    ew_c1, ew_c2 = st.columns(2)
    ew_override = ew_c1.number_input(t("Tỷ trọng Vốn chủ E/V (%)"), min_value=0.0, max_value=100.0, value=round(e_weight * 100, 0), step=1.0, format="%.0f") / 100
    dw_override = 1 - ew_override
    ew_c2.metric(t("Tỷ trọng Nợ D/V (%)"), f"{dw_override*100:.0f}%")
    
    wacc = ew_override * ke + dw_override * kd_at
    
    st.latex(f"WACC = {ew_override*100:.0f}\\% \\times {ke*100:.2f}\\% + {dw_override*100:.0f}\\% \\times {kd_at*100:.2f}\\% = {wacc*100:.2f}\\%")
    
    st.markdown(f"""
    <div style="text-align:center; padding: 20px; background: white; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
        <div style="font-size: 16px; color: #7F8C8D; text-transform: uppercase; font-weight: 600;">WACC</div>
        <div style="font-size: 48px; font-weight: 300; color: #8E44AD;">{wacc*100:.2f}%</div>
    </div>
    """, unsafe_allow_html=True)
    
    st.session_state['wacc'] = wacc
    st.session_state['ke'] = ke
    st.session_state['kd_at'] = kd_at
    st.session_state['tax_rate'] = tax_rate
    st.session_state['ew'] = ew_override
    
    # Dynamic WACC Explanation
    st.markdown("---")
    st.markdown(f"### 💡 {t('Các chỉ số trên nói lên điều gì?')}")
    
    is_en = st.session_state.get('lang', '') == '🇬🇧 English'
    
    # Ke interpretation
    ke_level = t('thấp') if ke < 0.10 else (t('trung bình') if ke < 0.15 else t('cao'))
    if is_en:
        ke_explain = f'**Cost of Equity (Ke = {ke*100:.2f}%)** is at a **{ke_level}** level. This is the minimum return investors require to compensate for holding the stock. Higher Ke → investors demand more profit → lower enterprise value.'
        kd_explain = f'**After-tax Cost of Debt (Kd = {kd_at*100:.2f}%)** reflects the effective interest rate the company pays on its borrowings after the tax shield. Kd lower than Ke shows debt is cheaper than equity — this is why companies use financial leverage.'
    else:
        ke_explain = f'**Chi phí Vốn chủ (Ke = {ke*100:.2f}%)** đang ở mức **{ke_level}**. Đây là mức lợi suất tối thiểu mà nhà đầu tư yêu cầu để bù đắp rủi ro khi nắm giữ cổ phiếu. Ke càng cao → nhà đầu tư đòi hỏi lợi nhuận càng lớn → giá trị doanh nghiệp càng thấp.'
        kd_explain = f'**Chi phí Nợ sau thuế (Kd = {kd_at*100:.2f}%)** phản ánh lãi suất thực tế mà doanh nghiệp phải trả cho các khoản vay, sau khi trừ lá chắn thuế. Kd thấp hơn Ke cho thấy nợ vay rẻ hơn vốn chủ — đây là lý do doanh nghiệp sử dụng đòn bẩy tài chính.'
    
    # WACC interpretation
    if is_en:
        if wacc < 0.08:
            wacc_verdict = '🟢 **Low WACC (<8%)**: Attractive cost of capital — the company can create value more easily if it invests in projects returning above WACC.'
        elif wacc < 0.12:
            wacc_verdict = '🟡 **Moderate WACC (8-12%)**: Normal cost of capital for the Vietnamese market. The company needs stable growth to clear this capital cost hurdle.'
        else:
            wacc_verdict = '🔴 **High WACC (>12%)**: Expensive cost of capital — the company needs projects with very high profit margins to create shareholder value.'
        wacc_desc = f'⚖️ <b>WACC = {wacc*100:.2f}%</b> is the discount rate used to convert future cash flows to present value in the DCF model. It reflects the weighted average opportunity cost of all capital sources (E: {ew_override*100:.0f}% + D: {dw_override*100:.0f}%).'
    else:
        if wacc < 0.08:
            wacc_verdict = '🟢 WACC thấp (<8%): Chi phí vốn hấp dẫn, doanh nghiệp dễ tạo giá trị khi có dự án đầu tư trên mức WACC.'
        elif wacc < 0.12:
            wacc_verdict = '🟡 WACC trung bình (8-12%): Mức chi phí vốn bình thường cho thị trường Việt Nam. Doanh nghiệp cần tăng trưởng ổn định để vượt qua rào cản chi phí vốn này.'
        else:
            wacc_verdict = '🔴 WACC cao (>12%): Chi phí vốn đắt đỏ, doanh nghiệp cần dự án có biên lợi nhuận rất cao mới tạo được giá trị cho cổ đông.'
        wacc_desc = f'⚖️ <b>WACC = {wacc*100:.2f}%</b> là tỷ suất chiết khấu dùng để quy đổi dòng tiền tương lai về hiện tại trong mô hình DCF. Nó phản ánh chi phí cơ hội trung bình có trọng số của toàn bộ nguồn vốn (E: {ew_override*100:.0f}% + D: {dw_override*100:.0f}%).'
    
    st.markdown(f"""
<div style="background: linear-gradient(135deg, #F8F9FA 0%, #E8F6F3 100%); border-radius: 10px; padding: 18px 20px; border-left: 4px solid #1ABC9C; margin-bottom: 10px;">
  <div style="font-size: 14px; line-height: 1.7; color: #2C3E50;">
    📈 {ke_explain}<br><br>
    💳 {kd_explain}<br><br>
    {wacc_desc}<br><br>
    {wacc_verdict}
  </div>
</div>
""", unsafe_allow_html=True)

# ===================== TAB 3: DCF MODEL =====================
elif selected_tab == t("Mô hình DCF (5 năm)"):
    st.markdown(f'<div class="section-title">📊 {t("Mô hình DCF (5 năm)")} — {company_name}</div>', unsafe_allow_html=True)
    
    wacc = st.session_state.get('wacc', 0.10)
    tax_rate = st.session_state.get('tax_rate', 0.20)
    
    # Assumptions
    st.markdown(f"### ⚙️ Assumptions")
    a1, a2, a3 = st.columns(3)
    with a1:
        rev_g = st.number_input(t("Tăng trưởng Doanh thu (%)"), min_value=-20.0, max_value=50.0, value=10.0, step=1.0, format="%.1f") / 100
        ebit_m = st.number_input(t("Biên EBIT (%)"), min_value=0.0, max_value=60.0, value=round((ebit/rev*100) if rev > 0 else 15.0, 1), step=1.0, format="%.1f") / 100
    with a2:
        capex_r = st.number_input(t("CAPEX / Doanh thu (%)"), min_value=0.0, max_value=30.0, value=5.0, step=0.5, format="%.1f") / 100
        da_r = st.number_input(t("D&A / Doanh thu (%)"), min_value=0.0, max_value=15.0, value=round((da/rev*100) if rev > 0 else 3.0, 1), step=0.5, format="%.1f") / 100
    with a3:
        dwc_r = st.number_input(t("ΔWC / Doanh thu (%)"), min_value=-10.0, max_value=10.0, value=1.0, step=0.5, format="%.1f") / 100
        g_term = st.number_input(t("Tăng trưởng dài hạn g (%)"), min_value=0.0, max_value=8.0, value=3.0, step=0.5, format="%.1f") / 100
    
    # Terminal Value Method
    st.markdown("---")
    tv_col1, tv_col2 = st.columns([1, 2])
    with tv_col1:
        tv_method = st.radio(
            t("Phương pháp Terminal Value"),
            ["Gordon Growth", "Exit Multiple"],
            index=0,
            help=t("Gordon Growth: TV = FCF×(1+g)/(WACC-g). Exit Multiple: TV = EBITDA×Multiple")
        )
    with tv_col2:
        if tv_method == "Exit Multiple":
            exit_multiple = st.number_input("EV/EBITDA Multiple", min_value=3.0, max_value=25.0, value=8.0, step=0.5, format="%.1f")
        else:
            exit_multiple = 0
    
    st.markdown("---")
    
    # Build projection table
    proj_years = [f"Y+{i}" for i in range(1, 6)]
    base_rev = rev if rev > 0 else 1000
    
    rows = {}
    proj_revs = []
    fcffs = []
    pv_fcffs = []
    
    for i in range(5):
        p_rev = base_rev * ((1 + rev_g) ** (i + 1))
        p_ebit = p_rev * ebit_m
        p_tax = p_ebit * tax_rate
        p_nopat = p_ebit - p_tax
        p_da = p_rev * da_r
        p_capex = p_rev * capex_r
        p_dwc = p_rev * dwc_r
        p_fcff = p_nopat + p_da - p_capex - p_dwc
        pv_factor = 1 / ((1 + wacc) ** (i + 1))
        pv_fcff = p_fcff * pv_factor

        proj_revs.append(p_rev)
        fcffs.append(p_fcff)
        pv_fcffs.append(pv_fcff)
    
    # Table display
    table_data = {
        t("Năm"): [f"Y0 ({latest_year})"] + proj_years,
        t("Doanh thu"): [f"{base_rev:,.0f}"] + [f"{r:,.0f}" for r in proj_revs],
        "EBIT": [f"{ebit:,.0f}"] + [f"{r*ebit_m:,.0f}" for r in proj_revs],
        "NOPAT": ["—"] + [f"{r*ebit_m*(1-tax_rate):,.0f}" for r in proj_revs],
        f"+ D&A": ["—"] + [f"{r*da_r:,.0f}" for r in proj_revs],
        f"- CAPEX": ["—"] + [f"{r*capex_r:,.0f}" for r in proj_revs],
        f"- ΔWC": ["—"] + [f"{r*dwc_r:,.0f}" for r in proj_revs],
        "**FCFF**": [f"—"] + [f"**{f:,.0f}**" for f in fcffs],
        "PV(FCFF)": ["—"] + [f"{p:,.0f}" for p in pv_fcffs],
    }
    st.dataframe(pd.DataFrame(table_data).set_index(t("Năm")), use_container_width=True)
    
    # Terminal Value + EV
    if tv_method == "Exit Multiple":
        ebitda_y5 = proj_revs[-1] * ebit_m + proj_revs[-1] * da_r
        tv = ebitda_y5 * exit_multiple
    else:
        tv = fcffs[-1] * (1 + g_term) / (wacc - g_term) if wacc > g_term else 0
    pv_tv = tv / ((1 + wacc) ** 5)
    sum_pv_fcff = sum(pv_fcffs)
    ev = sum_pv_fcff + pv_tv
    net_debt = total_debt - cash
    equity_value = ev - net_debt
    # equity_value đã ở Bn VND, shares ở triệu cổ phiếu
    fair_value_ps = (equity_value * 1e9) / (shares_mil * 1e6) if shares_mil > 0 else 0
    
    # TV% of EV
    tv_pct = (pv_tv / ev * 100) if ev > 0 else 0
    
    # Save these to session state so Tab 5 can read them without re-executing Tab 3
    st.session_state['ev'] = ev
    st.session_state['equity_value'] = equity_value
    st.session_state['fair_value_ps'] = fair_value_ps
    
    st.markdown("---")
    st.markdown("### 🎯 Valuation Summary")
    
    v1, v2, v3, v4 = st.columns(4)
    v1.markdown(f'<div class="met-card"><div class="met-label">Σ PV(FCFF)</div><div class="met-val teal">{sum_pv_fcff:,.0f}</div><div class="met-goal">Bn VND</div></div>', unsafe_allow_html=True)
    v2.markdown(f'<div class="met-card"><div class="met-label">PV(Terminal Value)</div><div class="met-val teal">{pv_tv:,.0f}</div><div class="met-goal">Bn VND</div></div>', unsafe_allow_html=True)
    v3.markdown(f'<div class="met-card"><div class="met-label">{t("Giá trị Doanh nghiệp (EV)")}</div><div class="met-val purple">{ev:,.0f}</div><div class="met-goal">Bn VND</div></div>', unsafe_allow_html=True)
    v4.markdown(f'<div class="met-card"><div class="met-label">{t("Giá trị Vốn chủ sở hữu")}</div><div class="met-val purple">{equity_value:,.0f}</div><div class="met-goal">Bn VND</div></div>', unsafe_allow_html=True)
    
    if shares_mil > 0:
        st.markdown(f"""
        <div style="text-align:center; padding: 20px; margin-top: 15px; background: linear-gradient(135deg, #8E44AD, #3498DB); border-radius: 8px; color: white;">
            <div style="font-size: 14px; text-transform: uppercase; letter-spacing: 2px;">{t('Giá trị hợp lý / Cổ phiếu')}</div>
            <div style="font-size: 48px; font-weight: 700;">{fair_value_ps:,.0f} VND</div>
            {'<div style="font-size:14px; margin-top:5px;">Giá thị trường: ' + f'{market_price:,.0f} VND → Upside: {((fair_value_ps/market_price)-1)*100:+.1f}%' + '</div>' if market_price > 0 else ''}
        </div>
        """, unsafe_allow_html=True)
    
    # Waterfall: EV breakdown
    fig_ev = go.Figure(go.Waterfall(
        orientation="v",
        measure=["absolute", "absolute", "total", "relative", "relative", "total"],
        x=["Σ PV(FCFF)", "PV(TV)", t("Giá trị Doanh nghiệp (EV)"), f"(-) {t('Tổng Nợ')}", f"(+) {t('Tiền mặt')}", t("Giá trị Vốn chủ sở hữu")],
        text=[f"{sum_pv_fcff:,.0f}", f"{pv_tv:,.0f}", f"{ev:,.0f}", f"-{total_debt:,.0f}", f"+{cash:,.0f}", f"{equity_value:,.0f}"],
        y=[sum_pv_fcff, pv_tv, ev, -total_debt, cash, equity_value],
        textposition="outside",
        connector={"line": {"color": "rgb(63,63,63)"}},
        increasing={"marker": {"color": "#2ECC71"}},
        decreasing={"marker": {"color": "#E74C3C"}},
        totals={"marker": {"color": "#3498DB"}}
    ))
    fig_ev.update_layout(height=400, margin=dict(t=30, b=10, l=10, r=10), showlegend=False, plot_bgcolor='rgba(0,0,0,0)')
    st.plotly_chart(fig_ev, use_container_width=True)
    
    # === TV% OF EV — PIE + WARNING ===
    st.markdown("---")
    st.markdown(f"### 📐 {t('Cấu trúc Giá trị Doanh nghiệp (EV Composition)')}")
    
    pie_col, warn_col = st.columns([1, 1])
    with pie_col:
        fig_tv = go.Figure(go.Pie(
            labels=["Σ PV(FCFF)", "PV(Terminal Value)"],
            values=[max(sum_pv_fcff, 0), max(pv_tv, 0)],
            marker_colors=['#1ABC9C', '#E67E22'],
            hole=0.5, textinfo='label+percent'
        ))
        fig_tv.update_layout(height=280, margin=dict(t=10, b=10, l=10, r=10), showlegend=False,
                             annotations=[dict(text=f"{tv_pct:.0f}%<br>TV", x=0.5, y=0.5, font_size=20, showarrow=False)])
        st.plotly_chart(fig_tv, use_container_width=True)
    
    with warn_col:
        if tv_pct > 80:
            st.error(f"""⚠️ **{t('Terminal Value chiếm')} {tv_pct:.0f}% {t('tổng EV')}!**
            
{t('Mô hình **cực kỳ nhạy cảm** với giả định dài hạn')} (Terminal Growth g = {g_term*100:.1f}%). {t('Thay đổi nhỏ trong g hoặc WACC sẽ thay đổi kết quả rất lớn.')}

**{t('Khuyến nghị:')}** {t('Kiểm tra lại giả định hoặc sử dụng phương pháp Exit Multiple để cross-check.')}""")
        elif tv_pct > 60:
            st.warning(f"""🟡 **{t('Terminal Value chiếm')} {tv_pct:.0f}% {t('tổng EV')}**
            
{t('Đây là mức bình thường trong DCF (thường 60-75%). Tuy nhiên, nhà đầu tư nên chạy Sensitivity Analysis để kiểm tra độ nhạy.')}""")
        else:
            st.success(f"""✅ **{t('Terminal Value chiếm')} {tv_pct:.0f}% {t('tổng EV')}**
            
{t('Mô hình có nền tảng vững — phần lớn giá trị đến từ dòng tiền dự phóng 5 năm, không phụ thuộc quá nhiều vào giả định dài hạn.')}""")
        
        st.markdown(f"""
| {t('Thành phần')} | {t('Giá trị')} | {t('Tỷ trọng')} |
|------------|---------|----------|
| Σ PV(FCFF) | {sum_pv_fcff:,.0f} Bn | {100-tv_pct:.0f}% |
| PV(Terminal Value) | {pv_tv:,.0f} Bn | {tv_pct:.0f}% |
| **Enterprise Value** | **{ev:,.0f} Bn** | **100%** |
| TV Method | {'Gordon Growth (g=' + f'{g_term*100:.1f}%' + ')' if tv_method == 'Gordon Growth' else 'Exit Multiple (' + f'{exit_multiple:.1f}' + 'x EBITDA)'} | |
""")
    
    st.session_state['ev'] = ev
    st.session_state['equity_value'] = equity_value
    st.session_state['fair_value_ps'] = fair_value_ps
    st.session_state['fcffs'] = fcffs
    st.session_state['wacc'] = wacc
    st.session_state['g_term'] = g_term
    st.session_state['tax_rate'] = tax_rate
    st.session_state['tv_pct'] = tv_pct
    st.session_state['tv_method'] = tv_method

# ===================== TAB 4: SENSITIVITY =====================
elif selected_tab == t("Phân tích Độ nhạy"):
    st.markdown(f'<div class="section-title">🎛️ {t("Phân tích Độ nhạy")}</div>', unsafe_allow_html=True)
    
    wacc_base = st.session_state.get('wacc', 0.10)
    g_base = st.session_state.get('g_term', 0.03)
    fcffs = st.session_state.get('fcffs', [1000]*5)
    
    wacc_range = [wacc_base + d for d in [-0.02, -0.01, 0, 0.01, 0.02]]
    g_range = [g_base + d for d in [-0.02, -0.01, 0, 0.01, 0.02]]
    
    st.markdown("### WACC × Terminal Growth → Equity Value / Share")
    
    matrix = []
    for g in reversed(g_range):
        row = []
        for w in wacc_range:
            if w <= g:
                row.append("N/A")
                continue
            tv_s = fcffs[-1] * (1 + g) / (w - g)
            pv_tv_s = tv_s / ((1 + w) ** 5)
            pv_fcff_s = sum([f / ((1 + w) ** (i+1)) for i, f in enumerate(fcffs)])
            ev_s = pv_fcff_s + pv_tv_s
            eq_s = ev_s - (total_debt - cash)
            if shares_mil > 0:
                fv_s = (eq_s * 1e9) / (shares_mil * 1e6)  # Bn VND → VND/share
                row.append(f"{fv_s:,.0f}")
            else:
                row.append(f"{eq_s:,.0f}")
        matrix.append(row)
    
    w_labels = [f"{w*100:.1f}%" for w in wacc_range]
    g_labels = [f"{g*100:.1f}%" for g in reversed(g_range)]
    
    sens_df = pd.DataFrame(matrix, index=g_labels, columns=w_labels)
    sens_df.index.name = "g \\ WACC"
    
    def color_sens(val):
        try:
            v = float(str(val).replace(',', ''))
            if market_price > 0:
                if v > market_price * 1.15:
                    return 'background-color: #D5F5E3; color: #1E8449;'
                elif v < market_price * 0.85:
                    return 'background-color: #FADBD8; color: #922B21;'
                else:
                    return 'background-color: #FEF9E7; color: #7D6608;'
            return ''
        except:
            return 'background-color: #D5D8DC;'
    
    styled = sens_df.style
    if hasattr(styled, 'map'):
        styled = styled.map(color_sens)
    else:
        styled = styled.applymap(color_sens)
    st.dataframe(styled, use_container_width=True)
    
    if market_price > 0:
        is_en = st.session_state.get('lang', '') == '🇬🇧 English'
        if is_en:
            lbl_under  = f"🟢 Undervalued (>+15% vs {market_price:,.0f} VND)"
            lbl_fair   = "🟡 Fair Value (±15%)"
            lbl_over   = "🔴 Overvalued (>-15%)"
        else:
            lbl_under  = f"🟢 Định giá thấp (>+15% so với {market_price:,.0f} VND)"
            lbl_fair   = "🟡 Giá hợp lý (±15%)"
            lbl_over   = "🔴 Định giá cao (>-15%)"
        st.markdown(f"""
        <div style="display:flex; gap:15px; justify-content:center; margin:10px 0; font-size:13px;">
            <span>{lbl_under}</span>
            <span>{lbl_fair}</span>
            <span>{lbl_over}</span>
        </div>
        """, unsafe_allow_html=True)
    
    # Dynamic Sensitivity Explanation
    st.markdown("---")
    st.markdown(f"### 💡 {t('Bảng Độ nhạy nói lên điều gì?')}")
    
    is_en = st.session_state.get('lang', '') == '🇬🇧 English'
    
    # Find min/max in the matrix
    all_vals = []
    for row in matrix:
        for v in row:
            try:
                all_vals.append(float(str(v).replace(',', '')))
            except:
                pass
    val_min = min(all_vals) if all_vals else 0
    val_max = max(all_vals) if all_vals else 0
    val_spread = val_max - val_min
    
    if market_price > 0:
        green_count = sum(1 for v in all_vals if v > market_price * 1.15)
        red_count = sum(1 for v in all_vals if v < market_price * 0.85)
        total_cells = len(all_vals)
        green_pct = green_count / total_cells * 100 if total_cells > 0 else 0
        red_pct = red_count / total_cells * 100 if total_cells > 0 else 0
        
        if is_en:
            if green_pct > 60:
                verdict = f'🟢 **Conclusion: The majority of scenarios (>{green_pct:.0f}%) suggest the stock is UNDERVALUED** vs the market. Wide margin of safety — a positive signal for long-term investors.'
            elif red_pct > 60:
                verdict = f'🔴 **Conclusion: The majority of scenarios (>{red_pct:.0f}%) suggest the stock is OVERVALUED** vs the market. High downside risk — caution advised before buying.'
            else:
                verdict = '🟡 **Conclusion: Valuation is around fair value.** Results depend heavily on WACC and growth assumptions — consider combining with qualitative analysis.'
        else:
            if green_pct > 60:
                verdict = f'🟢 **Kết luận: Phần lớn kịch bản (>{green_pct:.0f}%) cho thấy cổ phiếu đang ĐỊNH GIÁ THẤP** so với thị trường. Biên an toàn rộng — đây là tín hiệu tích cực cho nhà đầu tư dài hạn.'
            elif red_pct > 60:
                verdict = f'🔴 **Kết luận: Phần lớn kịch bản (>{red_pct:.0f}%) cho thấy cổ phiếu đang ĐỊNH GIÁ CAO** so với thị trường. Rủi ro downside lớn — cần thận trọng khi mua vào.'
            else:
                verdict = '🟡 **Kết luận: Định giá xoay quanh vùng giá hợp lý.** Kết quả phụ thuộc nhiều vào giả định WACC và tốc độ tăng trưởng — nên kết hợp thêm phân tích định tính.'
    else:
        verdict = ''
    
    if is_en:
        how_to_read = f'📊 <b>{t("Cách đọc bảng:")}</b> Each cell shows the <b>{t("Giá trị hợp lý / Cổ phiếu (VND)")}</b> for a pair of assumptions (WACC, g). Columns are WACC (cost of capital), rows are g (terminal growth rate).<br><br>'
        val_range_text = f'🎯 <b>{t("Biên độ dao động:")}</b> Fair value ranges from <b>{val_min:,.0f}</b> to <b>{val_max:,.0f} VND</b> (spread: {val_spread:,.0f} VND) — showing how sensitive the model is to input assumptions.<br><br>'
        practical_text = f'🔍 <b>{t("Ý nghĩa thực tiễn:")}</b> If you are confident WACC and g fall within the middle range, focus on the central cluster of cells. Corner cells represent extreme scenarios (most optimistic/pessimistic).<br><br>'
    else:
        how_to_read = f'📊 <b>{t("Cách đọc bảng:")}</b> Mỗi ô hiển thị <b>{t("Giá trị hợp lý / Cổ phiếu (VND)")}</b> ứng với một cặp giả định (WACC, g). Cột là WACC (chi phí vốn), hàng là g (tăng trưởng dài hạn).<br><br>'
        val_range_text = f'🎯 <b>{t("Biên độ dao động:")}</b> Giá trị hợp lý dao động từ <b>{val_min:,.0f}</b> đến <b>{val_max:,.0f} VND</b> (chênh lệch {val_spread:,.0f} VND) — cho thấy mức độ nhạy cảm của mô hình với các giả định đầu vào.<br><br>'
        practical_text = f'🔍 <b>{t("Ý nghĩa thực tiễn:")}</b> Nếu bạn tự tin WACC và g nằm trong khoảng giữa bảng, hãy tập trung vào cụm ô trung tâm. Các ô góc là kịch bản cực đoan (lạc quan/bi quan nhất).<br><br>'
    
    st.markdown(f"""
<div style="background: linear-gradient(135deg, #F8F9FA 0%, #FEF9E7 100%); border-radius: 10px; padding: 18px 20px; border-left: 4px solid #F39C12; margin-bottom: 10px;">
  <div style="font-size: 14px; line-height: 1.7; color: #2C3E50;">
    {how_to_read}{val_range_text}{practical_text}{verdict}
  </div>
</div>
""", unsafe_allow_html=True)

# ===================== TAB 5: CONCLUSION & EXPORT =====================
elif selected_tab == t("Kết luận & Xuất Excel"):
    st.markdown(f'<div class="section-title">📑 {t("Kết luận & Xuất Excel")} — {company_name}</div>', unsafe_allow_html=True)
    
    ev = st.session_state.get('ev', 0)
    equity_value = st.session_state.get('equity_value', 0)
    fair_value_ps = st.session_state.get('fair_value_ps', 0)
    wacc = st.session_state.get('wacc', 0.10)
    
    # Verdict
    if shares_mil > 0 and market_price > 0 and fair_value_ps > 0:
        upside = ((fair_value_ps / market_price) - 1) * 100
        if upside > 15:
            verdict, color, icon = "MUA (BUY)", "#27AE60", "🟢"
        elif upside < -15:
            verdict, color, icon = "BÁN (SELL)", "#E74C3C", "🔴"
        else:
            verdict, color, icon = "GIỮ (HOLD)", "#F39C12", "🟡"
        
        st.markdown(f"""
        <div style="text-align:center; padding:30px; background:white; border-radius:8px; border-left:6px solid {color}; box-shadow:0 2px 10px rgba(0,0,0,0.1);">
            <div style="font-size:60px;">{icon}</div>
            <div style="font-size:36px; font-weight:700; color:{color};">{verdict}</div>
            <div style="font-size:16px; color:#7F8C8D; margin-top:10px;">
                Fair Value: <b>{fair_value_ps:,.0f} VND</b> vs Market: <b>{market_price:,.0f} VND</b> → Upside: <b>{upside:+.1f}%</b>
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div style="text-align:center; padding:30px; background:white; border-radius:8px; border-left:6px solid #8E44AD; box-shadow:0 2px 10px rgba(0,0,0,0.1);">
            <div style="font-size:16px; color:#7F8C8D; text-transform:uppercase;">Equity Value ({t('Giá trị Vốn chủ sở hữu')})</div>
            <div style="font-size:48px; font-weight:700; color:#8E44AD;">{equity_value:,.0f} Bn VND</div>
            <div style="font-size:14px; color:#95A5A6; margin-top:5px;">Enterprise Value: {ev:,.0f} Bn VND | WACC: {wacc*100:.2f}%</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Export Excel
    st.markdown(f"### 📥 {t('Kết luận & Xuất Excel')}")
    
    def generate_report():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Sheet 1: Summary
            summary = pd.DataFrame({
                'Metric': ['Company', 'WACC', 'Enterprise Value (Bn)', 'Net Debt (Bn)', 'Equity Value (Bn)',
                           'Shares (Mil)', 'Fair Value/Share (VND)', 'Market Price (VND)', 'Upside/Downside'],
                'Value': [company_name, f"{wacc*100:.2f}%", f"{ev:,.0f}", f"{total_debt-cash:,.0f}", f"{equity_value:,.0f}",
                          f"{shares_mil:,.0f}", f"{fair_value_ps:,.0f}", f"{market_price:,.0f}",
                          f"{((fair_value_ps/market_price)-1)*100:+.1f}%" if market_price > 0 else "N/A"]
            })
            summary.to_excel(writer, sheet_name='Summary', index=False)
            
            # Sheet 2: Income Statement
            if not is_df.empty:
                is_df.to_excel(writer, sheet_name='IncomeStatement')
            
            # Sheet 3: Balance Sheet
            if not bs_df.empty:
                bs_df.to_excel(writer, sheet_name='BalanceSheet')
        
        output.seek(0)
        return output
    
    st.download_button(
        label=f"📥 Tải báo cáo DCF ({company_name}).xlsx",
        data=generate_report(),
        file_name=f"DCF_{company_name}_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

