import io
import re
import streamlit as st
import pandas as pd
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins
from openpyxl.utils import get_column_letter
from deep_translator import GoogleTranslator

# ============================================================
# CẤU HÌNH STREAMLIT
# ============================================================
st.set_page_config(
    page_title="Bảng Chấm Công Song Ngữ Động",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Bảng Chấm Công Song Ngữ Động (Trung - Việt)")
st.caption("Tự động nhận diện N dòng x M cột từ Ảnh/Excel & Tự động dịch bổ sung Tiếng Việt bên dưới dòng Tiếng Trung.")

# ============================================================
# HÀM XỬ LÝ DỊCH NỘI DUNG SONG NGỮ (LOGIC CỐT LÕI)
# ============================================================

def is_chinese(text: str) -> bool:
    """Kiểm tra xem chuỗi có chứa ký tự tiếng Trung hay không"""
    if not isinstance(text, str):
        return False
    return bool(re.search(r'[\u4e00-\u9fff]', text))

@st.cache_data(show_spinner=False)
def translate_to_vietnamese(text: str) -> str:
    """Dịch câu/từ từ Tiếng Trung sang Tiếng Việt"""
    try:
        translated = GoogleTranslator(source='zh-CN', target='vi').translate(text)
        return translated if translated else ""
    except Exception:
        return ""

def format_bilingual_cell(val):
    """
    Logic biến đổi ô:
    - Nếu có Tiếng Trung và chưa có Tiếng Việt -> Dịch và xuống dòng: "Tiếng Trung\nTiếng Việt"
    - Nếu là số, ô trống hoặc không có chữ Trung -> Giữ nguyên
    """
    if pd.isna(val) or val is None:
        return ""
    
    val_str = str(val).strip()
    if not val_str:
        return ""
    
    # Nếu là số nguyên / số thực thì giữ nguyên
    if val_str.replace('.', '', 1).isdigit():
        return val_str

    # Nếu ô chứa ký tự tiếng Trung
    if is_chinese(val_str):
        # Nếu trong ô đã có dấu xuống dòng hoặc đã có tiếng Việt thì giữ nguyên
        if "\n" in val_str:
            return val_str
        
        # Dịch tự động sang tiếng Việt
        vi_text = translate_to_vietnamese(val_str)
        if vi_text and vi_text.lower() != val_str.lower():
            # Ghép tiếng Trung ở trên, tiếng Việt ở dưới
            return f"{val_str}\n{vi_text}"
            
    return val_str

# ============================================================
# HÀM XỬ LÝ FILE ĐẦU VÀO (IMAGE & EXCEL)
# ============================================================

@st.cache_resource
def load_ocr_reader():
    import easyocr
    return easyocr.Reader(['ch_sim', 'en'])

def process_image_file(uploaded_file):
    """Trích xuất ma trận bảng từ file ảnh bằng OCR"""
    reader = load_ocr_reader()
    results = reader.readtext(uploaded_file.getvalue())
    
    if not results:
        return None
    
    # Sắp xếp các box nhận diện được theo dòng Y
    results_sorted = sorted(results, key=lambda x: x[0][0][1])
    lines = []
    current_line = []
    last_y = None
    
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
        max_cols = max(len(line) for line in lines)
        padded_lines = [line + [""] * (max_cols - len(line)) for line in lines]
        return pd.DataFrame(padded_lines)
    return None

def process_excel_file(uploaded_file):
    """Đọc dữ liệu từ file Excel/CSV tải lên"""
    try:
        if uploaded_file.name.endswith('.csv'):
            return pd.read_csv(uploaded_file, header=None)
        else:
            return pd.read_excel(uploaded_file, header=None)
    except Exception as e:
        st.error(f"Lỗi khi đọc file: {e}")
        return None

# ============================================================
# HÀM TẠO EXCEL SONG NGỮ CHUYÊN NGHIỆP (ĐỘNG N x M)
# ============================================================

def create_bilingual_excel(df_data, title_cn="员工上班", title_vi="Nhân viên đi làm"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bảng song ngữ"

    # Định dạng Font & Style chuẩn như code gốc
    font_name = "Microsoft YaHei"
    orange_fill = PatternFill(fill_type="solid", fgColor="ED7D00")
    thin_side = Side(style="thin", color="000000")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    num_rows, num_cols = df_data.shape

    # --------------------------------------------------------
    # 1. TIÊU ĐỀ BẢNG (Gộp từ cột 1 đến cột M)
    # --------------------------------------------------------
    last_col_letter = get_column_letter(num_cols)
    ws.merge_cells(f"A1:{last_col_letter}1")
    
    ws["A1"] = f"{title_cn}\n{title_vi}" if title_vi else title_cn
    ws["A1"].font = Font(name=font_name, size=13, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 42

    # --------------------------------------------------------
    # 2. DÒNG HEADER (Row 2 Excel) - Tự động tạo Song ngữ
    # --------------------------------------------------------
    headers = df_data.iloc[0].fillna("").tolist()
    for col_idx, h_val in enumerate(headers, start=1):
        cell_value = format_bilingual_cell(h_val)
        
        cell = ws.cell(row=2, column=col_idx, value=cell_value)
        cell.font = Font(name=font_name, size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = orange_fill
        cell.border = border
        
    ws.row_dimensions[2].height = 38

    # --------------------------------------------------------
    # 3. NỘI DUNG DỮ LIỆU (Row 3 đến N+1 Excel)
    # --------------------------------------------------------
    for r_idx in range(1, num_rows):
        excel_row = r_idx + 2
        ws.row_dimensions[excel_row].height = 32
        
        for c_idx in range(num_cols):
            raw_val = df_data.iloc[r_idx, c_idx]
            
            # Xử lý tạo nội dung Trung \n Việt
            cell_value = format_bilingual_cell(raw_val)

            # Ép kiểu số nếu nội dung chỉ là số thuần túy
            if str(cell_value).isdigit():
                cell_value = int(cell_value)

            cell = ws.cell(row=excel_row, column=c_idx + 1, value=cell_value)
            cell.font = Font(name=font_name, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

    # --------------------------------------------------------
    # 4. TỰ ĐỘNG CÂN CHỈNH ĐỘ RỘNG CỘT
    # --------------------------------------------------------
    for col_idx in range(1, num_cols + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 10
        for r_idx in range(1, num_rows + 2):
            val = str(ws.cell(row=r_idx, column=col_idx).value or "")
            for line in val.split("\n"):
                if len(line) > max_len:
                    max_len = len(line)
        ws.column_dimensions[col_letter].width = min(max_len + 6, 35)

    # --------------------------------------------------------
    # 5. THIẾT LẬP TRANG IN (FIT 1 PAGE, LANDSCAPE)
    # --------------------------------------------------------
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.2, right=0.2, top=0.3, bottom=0.3, header=0.1, footer=0.1)

    # Trả về ô nhớ RAM
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ============================================================
# GIAO DIỆN STREAMLIT
# ============================================================

uploaded_file = st.file_uploader(
    "📂 Tải lên File Ảnh (PNG, JPG) hoặc File Excel (XLSX, CSV)",
    type=["png", "jpg", "jpeg", "xlsx", "xls", "csv"]
)

col_t1, col_t2 = st.columns(2)
with col_t1:
    title_cn_in = st.text_input("Tiêu đề Tiếng Trung:", "2026年8月26日员工上班")
with col_t2:
    title_vi_in = st.text_input("Tiêu đề Tiếng Việt:", "Nhân viên đi làm ngày 26/08/2026")

if uploaded_file is not None:
    file_ext = uploaded_file.name.split('.')[-1].lower()
    
    with st.spinner("⏳ Đang nhận diện dữ liệu & tự động dịch Tiếng Việt song ngữ..."):
        if file_ext in ["png", "jpg", "jpeg"]:
            df_raw = process_image_file(uploaded_file)
        else:
            df_raw = process_excel_file(uploaded_file)

    if df_raw is not None and not df_raw.empty:
        st.success(f"✅ Đã đọc thành công ma trận dữ liệu: **{df_raw.shape[0]} dòng x {df_raw.shape[1]} cột**")

        # Xem trước dữ liệu sau khi tự động tạo song ngữ Trung \n Việt
        df_preview = df_raw.copy()
        for r in range(df_preview.shape[0]):
            for c in range(df_preview.shape[1]):
                df_preview.iloc[r, c] = format_bilingual_cell(df_preview.iloc[r, c])

        st.subheader("📋 Xem trước nội dung song ngữ sẽ xuất ra Excel")
        st.dataframe(df_preview, use_container_width=True)

        st.divider()

        # Tạo file Excel bằng RAM BytesIO
        excel_file = create_bilingual_excel(df_raw, title_cn=title_cn_in, title_vi=title_vi_in)

        st.download_button(
            label="⬇️ 下载 Excel / Tải Excel Song Ngữ",
            data=excel_file.getvalue(),
            file_name="Bang_Song_Ngu_Trung_Viet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.error("Không thể đọc được bảng từ file này, vui lòng kiểm tra lại file đầu vào!")
