import io
import re
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins
from openpyxl.utils import get_column_letter
from deep_translator import GoogleTranslator

# ============================================================
# CẤU HÌNH STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Hệ thống dịch & định dạng Bảng tự động",
    page_icon="📊",
    layout="centered"
)

st.title("📊 Hệ thống dịch & định dạng Bảng tự động (Trung - Việt)")
st.caption("Xử lý tự động linh hoạt cho MỌI file Ảnh hoặc Excel tải lên")

uploaded_file = st.file_uploader(
    "📂 Tải lên file Ảnh (PNG, JPG) hoặc file Excel/CSV bất kỳ",
    type=["png", "jpg", "jpeg", "xlsx", "xls", "csv"]
)

# ============================================================
# LOGIC XỬ LÝ CHUỖI & DỊCH THUẬT ĐỘNG
# ============================================================

def is_chinese(text: str) -> bool:
    """Kiểm tra xem chuỗi có chứa ký tự Tiếng Trung hay không"""
    if not isinstance(text, str):
        return False
    return bool(re.search(r'[\u4e00-\u9fff]', text))

@st.cache_data(show_spinner=False)
def translate_zh_to_vi(text: str) -> str:
    """Dịch động mọi chuỗi Tiếng Trung sang Tiếng Việt bằng API Google Translate"""
    if not text or not is_chinese(text):
        return text
    try:
        translated = GoogleTranslator(source='zh-CN', target='vi').translate(text.strip())
        return translated if translated else text
    except Exception:
        return text

def process_bilingual_cell(val):
    """
    Chuyển đổi ô dữ liệu thành định dạng 2 dòng:
    [Tiếng Trung gốc]
    [Tiếng Việt dịch]
    """
    if pd.isna(val) or val is None:
        return ""
    
    val_str = str(val).strip()
    if not val_str:
        return ""

    # Nếu là số hoặc đã có xuống dòng hoặc là mã STT -> Giữ nguyên
    if val_str.isdigit() or val_str.replace('.', '', 1).isdigit() or val_str.upper() == "STT":
        return val_str

    if "\n" in val_str:
        return val_str

    # Nếu chứa Tiếng Trung thì dịch và ghép dòng
    if is_chinese(val_str):
        vi_trans = translate_zh_to_vi(val_str)
        if vi_trans and vi_trans.lower() != val_str.lower():
            return f"{val_str}\n{vi_trans}"

    return val_str

# ============================================================
# TRÍCH XUẤT MA TRẬN DỮ LIỆU TỪ FILE ĐẦU VÀO
# ============================================================

@st.cache_resource
def load_ocr_reader():
    import easyocr
    return easyocr.Reader(['ch_sim', 'en'])

def extract_dataframe(file_obj):
    file_ext = file_obj.name.split('.')[-1].lower()
    
    if file_ext in ['xlsx', 'xls']:
        return pd.read_excel(file_obj, header=None)
    elif file_ext == 'csv':
        return pd.read_csv(file_obj, header=None)
    
    elif file_ext in ['png', 'jpg', 'jpeg']:
        reader = load_ocr_reader()
        results = reader.readtext(file_obj.getvalue())
        if not results:
            return None
        
        # Sắp xếp các box chữ theo tọa độ Y (dòng) và X (cột)
        results_sorted = sorted(results, key=lambda x: x[0][0][1])
        lines, current_line, last_y = [], [], None
        
        for bbox, text, prob in results_sorted:
            y_center = (bbox[0][1] + bbox[2][1]) / 2
            if last_y is None or abs(y_center - last_y) < 18:
                current_line.append((bbox[0][0], text))
                last_y = y_center
            else:
                current_line.sort(key=lambda x: x[0])
                lines.append([item[1] for item in current_line])
                current_line = [(bbox[0][0], text)]
                last_y = y_center
                
        if current_line:
            current_line.sort(key=lambda x: x[0])
            lines.append([item[1] for item in current_line])
            
        if lines:
            max_cols = max(len(l) for l in lines)
            padded = [l + [""] * (max_cols - len(l)) for l in lines]
            return pd.DataFrame(padded)
            
    return None

# ============================================================
# DỰNG FILE EXCEL ĐỘNG HOÀN TOÀN TỰ ĐỘNG
# ============================================================

def generate_dynamic_excel(df_matrix):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    font_name = "Microsoft YaHei"
    orange_fill = PatternFill(fill_type="solid", fgColor="ED7D00")
    thin_side = Side(style="thin", color="000000")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    total_rows, total_cols = df_matrix.shape

    # 1. TỰ ĐỘNG PHÁT HIỆN DÒNG TIÊU ĐỀ BẢNG (TITLE ROW GỘP Ô)
    # Kiểm tra xem dòng 0 có phải là 1 tiêu đề chung (ví dụ có chứa thông tin ngày tháng/tên bảng)
    first_row_non_empty = [str(x).strip() for x in df_matrix.iloc[0].values if pd.notna(x) and str(x).strip() != ""]
    
    is_title_row = False
    if len(first_row_non_empty) <= 2 and total_cols > 2:
        # Nếu dòng đầu tiên chỉ có 1-2 ô chứa chữ nhưng bảng có nhiều cột -> Đây là Dòng Title
        is_title_row = True

    start_data_idx = 0
    current_row = 1

    # Nếu có dòng Title gộp
    if is_title_row:
        raw_title = " ".join(first_row_non_empty)
        formatted_title = process_bilingual_cell(raw_title)
        
        last_col_letter = get_column_letter(total_cols)
        ws.merge_cells(f"A1:{last_col_letter}1")
        
        cell = ws["A1"]
        cell.value = formatted_title
        cell.font = Font(name=font_name, size=11, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 36
        
        start_data_idx = 1
        current_row = 2

    # 2. XỬ LÝ DÒNG HEADER BẢNG (MÀU CAM)
    header_vals = df_matrix.iloc[start_data_idx].values
    for c_idx in range(total_cols):
        raw_val = header_vals[c_idx]
        cell_val = process_bilingual_cell(raw_val)

        cell = ws.cell(row=current_row, column=c_idx + 1, value=cell_val)
        cell.font = Font(name=font_name, size=10, bold=True)
        cell.fill = orange_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    ws.row_dimensions[current_row].height = 36
    current_row += 1

    # 3. XỬ LÝ TOÀN BỘ CÁC DÒNG DỮ LIỆU
    for r_idx in range(start_data_idx + 1, total_rows):
        ws.row_dimensions[current_row].height = 32
        row_vals = df_matrix.iloc[r_idx].values

        for c_idx in range(total_cols):
            raw_val = row_vals[c_idx]
            val_str = str(raw_val).strip() if pd.notna(raw_val) else ""

            # Nếu dữ liệu là số nguyên thì đổi về dạng int để Excel tính toán được
            if val_str.isdigit():
                cell_val = int(val_str)
            else:
                cell_val = process_bilingual_cell(raw_val)

            cell = ws.cell(row=current_row, column=c_idx + 1, value=cell_val)
            cell.font = Font(name=font_name, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

        current_row += 1

    # 4. TỰ ĐỘNG TÍNH TOÁN ĐỘ RỘNG CỘT & ĐỊNH DẠNG TRANG IN
    for c_idx in range(1, total_cols + 1):
        col_letter = get_column_letter(c_idx)
        max_len = 8
        for r_idx in range(1, current_row):
            val = str(ws.cell(row=r_idx, column=c_idx).value or "")
            for line in val.split('\n'):
                if len(line) > max_len:
                    max_len = len(line)
        ws.column_dimensions[col_letter].width = min(max_len + 5, 30)

    # Cài đặt view & trang in Excel
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_margins = PageMargins(
        left=0.2, right=0.2, top=0.3, bottom=0.3, header=0.1, footer=0.1
    )

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ============================================================
# HIỂN THỊ TRÊN STREAMLIT
# ============================================================

if uploaded_file is not None:
    with st.spinner("⏳ Đang phân tích ma trận dữ liệu và dịch tự động..."):
        df_matrix = extract_dataframe(uploaded_file)

    if df_matrix is not None and not df_matrix.empty:
        r_count, c_count = df_matrix.shape
        st.success(f"Đã nhận diện thành công ma trận: **{r_count} dòng x {c_count} cột**")

        excel_file = generate_dynamic_excel(df_matrix)

        st.download_button(
            label="⬇️ Tải xuống file Excel kết quả",
            data=excel_file.getvalue(),
            file_name="Ket_qua_dich_bang_song_ngu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.error("Không nhận diện được bảng từ file này. Vui lòng tải lên file khác!")
