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
    page_title="Bảng chấm công Trung - Việt",
    page_icon="📊",
    layout="centered"
)

# ============================================================
# TIÊU ĐỀ
# ============================================================

st.title("📊 Bảng chấm công Trung - Việt")
st.caption("Ứng dụng xử lý động cho tất cả các file tải lên (Ảnh & Excel)")

# ============================================================
# TẢI FILE ĐẦU VÀO (FILE ẢNH HOẶC FILE EXCEL)
# ============================================================

uploaded_file = st.file_uploader(
    "📂 Tải lên file Ảnh (PNG, JPG) hoặc file Excel/CSV",
    type=["png", "jpg", "jpeg", "xlsx", "xls", "csv"]
)

# ============================================================
# HÀM BỔ TRỢ & LOGIC DỊCH TIẾNG TRUNG -> TIẾNG VIỆT
# ============================================================

def is_chinese(text: str) -> bool:
    """Kiểm tra chuỗi có chứa tiếng Trung hay không"""
    if not isinstance(text, str):
        return False
    return bool(re.search(r'[\u4e00-\u9fff]', text))

@st.cache_data(show_spinner=False)
def translate_zh_to_vi(text: str) -> str:
    """Dịch tự động Tiếng Trung sang Tiếng Việt"""
    try:
        translated = GoogleTranslator(source='zh-CN', target='vi').translate(text)
        return translated if translated else ""
    except Exception:
        return ""

def process_cell_bilingual(val):
    """
    Logic cốt lõi: Nếu ô chứa Tiếng Trung thì dịch và ghép Tiếng Việt xuống dưới (Trung\nViệt)
    """
    if pd.isna(val) or val is None:
        return ""
    
    val_str = str(val).strip()
    if not val_str:
        return ""
    
    # Nếu là số thuần túy thì giữ nguyên
    if val_str.replace('.', '', 1).isdigit():
        return val_str

    # Nếu có tiếng Trung
    if is_chinese(val_str):
        # Nếu đã có dấu xuống dòng (đã là dạng Trung\nViệt) thì giữ nguyên
        if "\n" in val_str:
            return val_str
        
        # Dịch và ghép tiếng Việt xuống dòng dưới
        vi_text = translate_zh_to_vi(val_str)
        if vi_text and vi_text.lower() != val_str.lower():
            return f"{val_str}\n{vi_text}"
            
    return val_str

# ============================================================
# HÀM ĐỌC MA TRẬN DỮ LIỆU TỪ FILE LOAD LÊN
# ============================================================

@st.cache_resource
def load_ocr_reader():
    import easyocr
    return easyocr.Reader(['ch_sim', 'en'])

def read_input_matrix(file_obj):
    """Đọc dữ liệu từ file load lên thành ma trận (DataFrame)"""
    file_ext = file_obj.name.split('.')[-1].lower()
    
    # 1. Nếu là file Excel / CSV
    if file_ext in ['xlsx', 'xls']:
        return pd.read_excel(file_obj, header=None)
    elif file_ext == 'csv':
        return pd.read_csv(file_obj, header=None)
    
    # 2. Nếu là file Ảnh (dùng OCR nhận diện ma trận dòng x cột)
    elif file_ext in ['png', 'jpg', 'jpeg']:
        reader = load_ocr_reader()
        results = reader.readtext(file_obj.getvalue())
        if not results:
            return None
        
        # Nhóm chữ theo dòng dựa trên tọa độ Y
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
# HÀM TẠO FILE EXCEL (GIỮ NGUYÊN LOGIC STYLING CỦA CODE GỐC)
# ============================================================

def create_excel_dynamic(df_matrix):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bảng song ngữ"

    font_name = "Microsoft YaHei"
    orange_fill = PatternFill(fill_type="solid", fgColor="ED7D00")
    thin_side = Side(style="thin", color="000000")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    num_rows, num_cols = df_matrix.shape

    # --------------------------------------------------------
    # HEADER (Dòng 1 của file load lên)
    # --------------------------------------------------------
    for c_idx in range(num_cols):
        raw_val = df_matrix.iloc[0, c_idx]
        cell_val = process_cell_bilingual(raw_val)
        
        cell = ws.cell(row=1, column=c_idx + 1, value=cell_val)
        cell.font = Font(name=font_name, size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = orange_fill
        cell.border = border

    ws.row_dimensions[1].height = 38

    # --------------------------------------------------------
    # CÁC DÒNG DỮ LIỆU (Từ dòng 2 trở đi)
    # --------------------------------------------------------
    for r_idx in range(1, num_rows):
        excel_row = r_idx + 1
        ws.row_dimensions[excel_row].height = 32
        
        for c_idx in range(num_cols):
            raw_val = df_matrix.iloc[r_idx, c_idx]
            cell_val = process_cell_bilingual(raw_val)

            # Đổi sang số nếu giá trị là số nguyên
            if str(cell_val).isdigit():
                cell_val = int(cell_val)

            cell = ws.cell(row=excel_row, column=c_idx + 1, value=cell_val)
            cell.font = Font(name=font_name, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

    # --------------------------------------------------------
    # ĐỘ RỘNG CỘT TỰ ĐỘNG THEO KÍCH THƯỚC FILE
    # --------------------------------------------------------
    for c_idx in range(1, num_cols + 1):
        col_letter = get_column_letter(c_idx)
        max_len = 10
        for r_idx in range(1, num_rows + 1):
            val = str(ws.cell(row=r_idx, column=c_idx).value or "")
            for line in val.split('\n'):
                if len(line) > max_len:
                    max_len = len(line)
        ws.column_dimensions[col_letter].width = min(max_len + 5, 30)

    # --------------------------------------------------------
    # CÀI ĐẶT TRANG IN GIỮ NGUYÊN CỦA CODE GỐC
    # --------------------------------------------------------
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
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
# XỬ LÝ DỮ LIỆU & HIỂN THỊ TRÊN STREAMLIT
# ============================================================

if uploaded_file is not None:
    with st.spinner("⏳ Đang xử lý file..."):
        df_matrix = read_input_matrix(uploaded_file)

    if df_matrix is not None and not df_matrix.empty:
        r_count, c_count = df_matrix.shape
        st.success(f"Đã nhận diện file có **{r_count} dòng** và **{c_count} cột**.")

        # ============================================================
        # HIỂN THỊ PREVIEW DỮ LIỆU TRÊN STREAMLIT
        # ============================================================
        st.subheader("📋 Nội dung bảng")

        # Tạo dữ liệu xem trước song ngữ
        preview_matrix = df_matrix.copy()
        for r in range(preview_matrix.shape[0]):
            for c in range(preview_matrix.shape[1]):
                preview_matrix.iloc[r, c] = process_cell_bilingual(preview_matrix.iloc[r, c])

        # Đặt dòng đầu tiên làm tên cột xem trước
        preview_df = pd.DataFrame(
            preview_matrix.iloc[1:].values,
            columns=preview_matrix.iloc[0].values
        )

        st.dataframe(
            preview_df,
            use_container_width=True,
            hide_index=True
        )

        # ============================================================
        # NÚT TẠO + DOWNLOAD EXCEL
        # ============================================================
        st.divider()
        st.subheader("📥 Xuất Excel")

        excel_file = create_excel_dynamic(df_matrix)

        st.download_button(
            label="⬇️ 下载 Excel / Tải Excel",
            data=excel_file.getvalue(),
            file_name=f"Bang_cham_cong_{r_count}x{c_count}_Trung_Viet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        st.caption(
            "Excel được tạo trực tiếp trong bộ nhớ RAM, "
            "không sử dụng đường dẫn /mnt/data nên phù hợp với Streamlit Cloud."
        )
    else:
        st.error("Không thể đọc được bảng từ file này. Vui lòng thử file khác!")
else:
    st.info("👆 Vui lòng tải lên 1 file (Ảnh hoặc Excel) để bắt đầu xử lý.")
