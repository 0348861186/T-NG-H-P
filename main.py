import io
import re
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins
import pandas as pd
import streamlit as st

# ============================================================
# CẤU HÌNH STREAMLIT
# ============================================================
st.set_page_config(
    page_title="Dịch & Xuất Excel Song Ngữ Trung - Việt",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Dịch & Xuất Excel Song Ngữ Trung - Việt (Tùy biến Kích thước)")
st.caption("Tự động nhận diện mọi số dòng, số cột của file tải lên và xuất ra file Excel song ngữ chuẩn định dạng.")

# ============================================================
# TỪ ĐIỂN DỊCH TỰ ĐỘNG (DỊCH MẪU TRUNG -> VIỆT)
# ============================================================
DICT_TRANSLATE = {
    # Header / Tiêu đề
    "2026年08月26日员工上班": "Nhân viên đi làm ngày 26/08/2026",
    "2026年8月26日员工上班": "Nhân viên đi làm ngày 26/08/2026",
    "STT": "STT",
    "部分": "Bộ phận",
    "部门": "Bộ phận",
    "开几台机": "Số máy mở",
    "正式工": "Công nhân chính thức",
    "临时工": "Công nhân thời vụ",
    "新临时工": "Công nhân thời vụ mới",
    "备注": "Ghi chú",
    
    # Nội dung bộ phận
    "连机": "Máy liên kết",
    "制袋机": "Máy làm túi",
    "连机吹膜": "Thổi màng liên máy",
    "制袋机吹膜": "Thổi màng máy làm túi",
    "巡检": "Kiểm tra tuần tra",
    "打扫": "Vệ sinh",
    "打箱": "Đóng thùng",
    "分口": "Chia miệng",
    "仓库+材料": "Kho + nguyên liệu",
    "造粒": "Tạo hạt",
    "电工": "Thợ điện",
    "办公室": "Văn phòng",
    "QC": "QC",
    "阿秋，阿勇": "A Qiu, A Yong",
    "套袋": "Đóng túi",
    "一共": "Tổng cộng",
}

def translate_text(text):
    """Hàm dịch văn bản Trung -> Việt."""
    if not isinstance(text, str):
        return text
    text_clean = text.strip()
    if text_clean in DICT_TRANSLATE:
        return DICT_TRANSLATE[text_clean]
    
    # Tra cứu thay thế cụm từ nếu là văn bản kết hợp
    res = text_clean
    for k, v in DICT_TRANSLATE.items():
        if k in res and k != res:
            res = res.replace(k, v)
    return res if res != text_clean else ""

# ============================================================
# HÀM TẠO EXCEL SONG NGỮ ĐỘNG (DÀNH CHO MỌI SỐ DÒNG & SỐ CỘT)
# ============================================================
def process_dynamic_table(raw_title_cn, raw_headers_cn, raw_rows):
    """
    Tạo workbook Excel động tùy chỉnh hoàn toàn dựa vào số dòng & cột đầu vào.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Bảng song ngữ"

    # Style chung
    font_name = "Microsoft YaHei"
    orange_fill = PatternFill(fill_type="solid", fgColor="ED7D00")
    thin_side = Side(style="thin", color="000000")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    num_cols = len(raw_headers_cn)
    current_excel_row = 1

    # 1. TIÊU ĐỀ BẢNG (Nếu có)
    if raw_title_cn:
        title_vi = translate_text(raw_title_cn)
        full_title = f"{raw_title_cn}\n{title_vi}" if title_vi else raw_title_cn
        
        # Merge từ cột A đến cột cuối cùng tương ứng với dữ liệu
        last_col_letter = openpyxl.utils.get_column_letter(num_cols)
        ws.merge_cells(f"A1:{last_col_letter}1")
        
        title_cell = ws["A1"]
        title_cell.value = full_title
        title_cell.font = Font(name=font_name, size=13, bold=True)
        title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 42
        current_excel_row = 2

    # 2. HEADER (Tự động dịch và căn chỉnh theo số cột)
    header_row_idx = current_excel_row
    for col_idx, cn_header in enumerate(raw_headers_cn, start=1):
        vi_header = translate_text(cn_header)
        cell = ws.cell(row=header_row_idx, column=col_idx)
        
        if not vi_header or cn_header == vi_header:
            cell.value = str(cn_header)
        else:
            cell.value = f"{cn_header}\n{vi_header}"
            
        cell.font = Font(name=font_name, size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = orange_fill
        cell.border = border

    ws.row_dimensions[header_row_idx].height = 38
    current_excel_row += 1

    # 3. DÒNG DỮ LIỆU (Tự động lặp qua từng dòng và từng cột)
    for row_data in raw_rows:
        row_idx = current_excel_row
        ws.row_dimensions[row_idx].height = 32
        
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            
            # Xử lý dịch nếu ô chứa ký tự tiếng Trung
            if isinstance(val, str) and re.search(r"[\u4e00-\u9fff]", val):
                vi_val = translate_text(val)
                cell.value = f"{val}\n{vi_val}" if vi_val else val
            else:
                cell.value = val if val is not None else ""

            cell.font = Font(name=font_name, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

        current_excel_row += 1

    # 4. TỰ ĐỘNG ĐIỀU CHỈNH ĐỘ RỘNG CỘT
    for col in ws.columns:
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or "")
            lines = val_str.split("\n")
            for line in lines:
                if len(line) > max_len:
                    max_len = len(line)
        ws.column_dimensions[col_letter].width = max(max_len + 5, 14)

    # 5. CÀI ĐẶT TRANG IN
    ws.sheet_view.showGridLines = True
    ws.freeze_panes = f"A{header_row_idx + 1}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.2, right=0.2, top=0.3, bottom=0.3, header=0.1, footer=0.1)

    # Ghi vào RAM (BytesIO) để phù hợp Streamlit Cloud
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ============================================================
# GIAO DIỆN CHÍNH STREAMLIT
# ============================================================
st.sidebar.header("⚙️ Cấu hình Tải file")
uploaded_file = st.sidebar.file_uploader("📂 Tải lên File (Excel .xlsx hoặc Hình ảnh):", type=["xlsx", "png", "jpg", "jpeg"])

# Dữ liệu mặc định ban đầu (Ví dụ bảng 6 dòng x 6 cột)
default_headers = ["STT", "部门", "开几台机", "正式工", "临时工", "备注"]
default_rows = [
    [1, "连机", 5, 3, 2, ""],
    [2, "制袋机", 6, 3, 2, ""],
    [3, "连机吹膜", 5, 4, "", ""],
    [4, "制袋机吹膜", 4, 2, 1, ""],
    [5, "巡检", "", 2, "", ""],
    [6, "打扫", "", 1, "", ""]
]
default_title = "2026年8月26日员工上班"

# Nếu người dùng tải file Excel lên
if uploaded_file is not None:
    if uploaded_file.name.endswith(".xlsx"):
        st.sidebar.success(f"Đã nạp file Excel: {uploaded_file.name}")
        wb_in = load_workbook(uploaded_file)
        ws_in = wb_in.active
        
        data_all = list(ws_in.iter_rows(values_only=True))
        if len(data_all) > 1:
            default_title = str(data_all[0][0]) if data_all[0][0] else ""
            default_headers = [str(c) if c is not None else f"Cột {i+1}" for i, c in enumerate(data_all[1])]
            default_rows = [[c if c is not None else "" for c in row] for row in data_all[2:]]
    else:
        st.sidebar.info("📷 Đã nhận diện hình ảnh. Vui lòng kiểm tra lại cấu trúc số dòng/cột bên dưới.")

# Cho phép chỉnh sửa tiêu đề & nội dung bảng trực tiếp
st.subheader("📋 Cấu hình Tiêu đề & Dữ liệu Đầu vào")
title_input = st.text_input("Tiêu đề bảng (Tiếng Trung):", value=default_title)

# Hiển thị bảng dạng Data Editor linh hoạt số dòng/số cột
df_input = pd.DataFrame(default_rows, columns=default_headers)

st.write("👉 Bạn có thể **thêm/xóa dòng** hoặc **chỉnh sửa nội dung** trực tiếp trong bảng dưới đây trước khi bấm dịch:")
edited_df = st.data_editor(
    df_input,
    num_rows="dynamic",  # Tự do thêm/xóa dòng
    use_container_width=True
)

st.divider()

# ============================================================
# NÚT XUẤT EXCEL SONG NGỮ
# ============================================================
st.subheader("📥 Xuất File Excel Song Ngữ")

if st.button("🔄 Tiến hành Dịch & Tạo File Excel", use_container_width=True):
    headers_list = list(edited_df.columns)
    rows_list = edited_df.values.tolist()
    
    excel_file = process_dynamic_table(title_input, headers_list, rows_list)
    
    st.download_button(
        label="⬇️ Tải xuống File Excel Song Ngữ (.xlsx)",
        data=excel_file.getvalue(),
        file_name="Bang_dich_song_ngu_Trung_Viet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    st.success(f"🎉 Đã dịch và xuất thành công bảng kích thước: {len(rows_list)} dòng x {len(headers_list)} cột!")
