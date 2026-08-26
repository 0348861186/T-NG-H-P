import io
import re
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins
import pandas as pd
import streamlit as st
from deep_translator import GoogleTranslator

# ============================================================
# CẤU HÌNH STREAMLIT
# ============================================================
st.set_page_config(
    page_title="Dịch & Xuất Excel Song Ngữ Tự Động 100%",
    page_icon="🌐",
    layout="wide"
)

st.title("🌐 Dịch & Xuất Excel Song Ngữ Tự Động 100% (Mọi File & Kích Thước)")
st.caption("Tự động dịch AI trực tiếp từ tiếng Trung sang tiếng Việt cho BẤT KỲ nội dung, số dòng, số cột nào mà KHÔNG CẦN từ điển cố định.")

# ============================================================
# HÀM DỊCH TỰ ĐỘNG BẰNG AI / GOOGLE TRANSLATE (CÓ CACHE TỐC ĐỘ)
# ============================================================
@st.cache_data(show_spinner=False)
def auto_translate_text(text):
    """
    Tự động dịch bất kỳ văn bản tiếng Trung nào sang tiếng Việt.
    Không phụ thuộc vào từ điển cố định.
    """
    if not text or not isinstance(text, str):
        return str(text) if text is not None else ""
    
    text_str = str(text).strip()
    
    # Nếu không chứa chữ Hán thì giữ nguyên (số, ký tự đặc biệt, tiếng Anh...)
    if not re.search(r"[\u4e00-\u9fff]", text_str):
        return text_str
    
    try:
        # Gọi thư viện dịch AI tự động 100%
        translated = GoogleTranslator(source='zh-CN', target='vi').translate(text_str)
        return translated if translated else text_str
    except Exception:
        return text_str

# ============================================================
# HÀM TẠO EXCEL SONG NGỮ TỰ ĐỘNG ĐỘNG
# ============================================================
def process_dynamic_table(raw_title_cn, raw_headers_cn, raw_rows):
    """
    Tạo workbook Excel động hoàn toàn:
    - Nhận diện đúng số dòng, số cột bất kỳ.
    - Tự động dịch AI Tiếng Trung -> Tiếng Việt cho từng ô.
    - Đặt tiếng Việt ngay bên dưới tiếng Trung trong cùng 1 ô.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Bảng song ngữ"

    # Style
    font_name = "Microsoft YaHei"
    orange_fill = PatternFill(fill_type="solid", fgColor="ED7D00")
    thin_side = Side(style="thin", color="000000")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    num_cols = len(raw_headers_cn)
    current_excel_row = 1

    # 1. TIÊU ĐỀ (Nếu có)
    if raw_title_cn and str(raw_title_cn).strip():
        title_cn = str(raw_title_cn).strip()
        title_vi = auto_translate_text(title_cn)
        
        full_title = f"{title_cn}\n{title_vi}" if title_vi and title_vi != title_cn else title_cn
        
        last_col_letter = openpyxl.utils.get_column_letter(num_cols)
        ws.merge_cells(f"A1:{last_col_letter}1")
        
        title_cell = ws["A1"]
        title_cell.value = full_title
        title_cell.font = Font(name=font_name, size=13, bold=True)
        title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 45
        current_excel_row = 2

    # 2. HEADER (Dịch AI & Căn chỉnh theo số cột)
    header_row_idx = current_excel_row
    for col_idx, cn_header in enumerate(raw_headers_cn, start=1):
        cn_text = str(cn_header).strip() if cn_header is not None else f"Cột {col_idx}"
        vi_text = auto_translate_text(cn_text)
        
        cell = ws.cell(row=header_row_idx, column=col_idx)
        if not vi_text or cn_text == vi_text:
            cell.value = cn_text
        else:
            cell.value = f"{cn_text}\n{vi_text}"
            
        cell.font = Font(name=font_name, size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = orange_fill
        cell.border = border

    ws.row_dimensions[header_row_idx].height = 38
    current_excel_row += 1

    # 3. DÒNG DỮ LIỆU (Tự động dịch từng ô & căn chỉnh)
    for row_data in raw_rows:
        row_idx = current_excel_row
        ws.row_dimensions[row_idx].height = 34
        
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            val_str = str(val).strip() if val is not None else ""
            
            if val_str and re.search(r"[\u4e00-\u9fff]", val_str):
                vi_val = auto_translate_text(val_str)
                cell.value = f"{val_str}\n{vi_val}" if vi_val and vi_val != val_str else val_str
            else:
                cell.value = val if val is not None else ""

            cell.font = Font(name=font_name, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

        current_excel_row += 1

    # 4. TỰ ĐỘNG CHỈNH ĐỘ RỘNG CỘT
    for col in ws.columns:
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or "")
            lines = val_str.split("\n")
            for line in lines:
                if len(line) > max_len:
                    max_len = len(line)
        ws.column_dimensions[col_letter].width = max(max_len + 6, 14)

    # 5. CÀI ĐẶT TRANG IN
    ws.sheet_view.showGridLines = True
    ws.freeze_panes = f"A{header_row_idx + 1}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.2, right=0.2, top=0.3, bottom=0.3, header=0.1, footer=0.1)

    # Ghi vào RAM (BytesIO)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ============================================================
# GIAO DIỆN CHÍNH STREAMLIT
# ============================================================
st.sidebar.header("⚙️ Tải File Đầu Vào")
uploaded_file = st.sidebar.file_uploader("📂 Tải lên bất kỳ file Excel (.xlsx):", type=["xlsx"])

# Dữ liệu mẫu ban đầu
default_headers = ["STT", "部门", "开几台机", "正式工", "临时工", "备注"]
default_rows = [
    [1, "连机", 5, 3, 2, ""],
    [2, "制袋机", 6, 3, 2, ""],
    [3, "连机吹膜", 5, 4, "", ""],
    [4, "制袋机吹膜", 4, 2, 1, ""],
    [5, "巡检", "", 2, "", ""]
]
default_title = "2026年8月26日员工上班"

# Tự động đọc dữ liệu khi người dùng tải file lên
if uploaded_file is not None:
    st.sidebar.success(f"Đã nhận file: {uploaded_file.name}")
    wb_in = load_workbook(uploaded_file)
    ws_in = wb_in.active
    
    data_all = list(ws_in.iter_rows(values_only=True))
    if len(data_all) >= 2:
        default_title = str(data_all[0][0]) if data_all[0][0] else ""
        default_headers = [str(c) if c is not None else f"Cột {i+1}" for i, c in enumerate(data_all[1])]
        default_rows = [[c if c is not None else "" for c in row] for row in data_all[2:]]

st.subheader("📋 Xem & Tùy Biến Bảng Dữ Liệu")
title_input = st.text_input("Tiêu đề tiếng Trung (Tự động nhận diện từ file):", value=default_title)

df_input = pd.DataFrame(default_rows, columns=default_headers)

st.write("👉 Bảng dưới đây tự động nhận diện đúng số dòng, số cột của file bạn tải lên. Bạn có thể thêm/xóa/sửa tự do:")
edited_df = st.data_editor(
    df_input,
    num_rows="dynamic",
    use_container_width=True
)

st.divider()

st.subheader("📥 Xuất File Excel Song Ngữ Dịch Tự Động")

if st.button("🚀 Bắt Đầu Dịch Tự Động & Xuất Excel", use_container_width=True):
    with st.spinner("🤖 Đang gọi AI dịch tự động từng ô tiếng Trung sang tiếng Việt..."):
        headers_list = list(edited_df.columns)
        rows_list = edited_df.values.tolist()
        
        excel_file = process_dynamic_table(title_input, headers_list, rows_list)
        
        st.download_button(
            label="⬇️ Tải Xuất File Excel Song Ngữ (.xlsx)",
            data=excel_file.getvalue(),
            file_name="Bang_Excel_Song_Ngu_Tu_Dong.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.success(f"✅ Đã dịch tự động xong toàn bộ {len(rows_list)} dòng x {len(headers_list)} cột!")
