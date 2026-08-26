import io
import streamlit as st
import pandas as pd
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins
from openpyxl.utils import get_column_letter

# ============================================================
# CẤU HÌNH STREAMLIT
# ============================================================
st.set_page_config(
    page_title="Chuyển Đổi Bảng Song Ngữ Động",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Chuyển Đổi Bảng Tùy Biến (Ảnh & Excel ➔ Excel Chuyên Nghiệp)")
st.caption("Tự động nhận diện số lượng dòng & cột từ file tải lên, giữ nguyên logic thiết kế/định dạng chuyên nghiệp.")

# ============================================================
# HÀM XỬ LÝ FILE ĐẦU VÀO (IMAGE & EXCEL)
# ============================================================

@st.cache_resource
def load_ocr_reader():
    import easyocr
    # Khởi tạo EasyOCR cho tiếng Trung giản thể (ch_sim), tiếng Anh (en)
    return easyocr.Reader(['ch_sim', 'en'])

def process_image_file(uploaded_file):
    """Trích xuất ma trận bảng từ file ảnh bằng OCR"""
    image = Image.open(uploaded_file)
    reader = load_ocr_reader()
    
    # Đọc text kèm tọa độ từ ảnh
    results = reader.readtext(uploaded_file.getvalue())
    
    if not results:
        return None
    
    # Nhóm các text theo dòng dựa trên tọa độ Y
    lines = []
    # Tự động gom dòng dựa vào tọa độ bbox
    results_sorted = sorted(results, key=lambda x: x[0][0][1]) # Sắp xếp theo y-min
    
    current_line = []
    last_y = None
    
    for bbox, text, prob in results_sorted:
        y_center = (bbox[0][1] + bbox[2][1]) / 2
        if last_y is None or abs(y_center - last_y) < 15: # Ngưỡng dòng
            current_line.append((bbox[0][0], text)) # Thêm x-min và text
            last_y = y_center
        else:
            # Sắp xếp các từ trong dòng theo x-min
            current_line.sort(key=lambda x: x[0])
            lines.append([item[1] for item in current_line])
            current_line = [(bbox[0][0], text)]
            last_y = y_center
            
    if current_line:
        current_line.sort(key=lambda x: x[0])
        lines.append([item[1] for item in current_line])
        
    # Tạo DataFrame từ các dòng nhận diện được
    if lines:
        max_cols = max(len(line) for line in lines)
        padded_lines = [line + [""] * (max_cols - len(line)) for line in lines]
        df = pd.DataFrame(padded_lines)
        return df
    return None

def process_excel_file(uploaded_file):
    """Đọc dữ liệu từ file Excel/CSV tải lên"""
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, header=None)
        else:
            df = pd.read_excel(uploaded_file, header=None)
        return df
    except Exception as e:
        st.error(f"Lỗi khi đọc file Excel: {e}")
        return None

# ============================================================
# HÀM TẠO FILE EXCEL TỰ ĐỘNG THEO MA TRẬN DÒNG x CỘT
# ============================================================

def create_dynamic_excel(df_data, title_text="BẢNG DỮ LIỆU / 数据表"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bảng dữ liệu"

    # Style Configurations
    font_name = "Microsoft YaHei"
    orange_fill = PatternFill(fill_type="solid", fgColor="ED7D00")
    thin_side = Side(style="thin", color="000000")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    num_rows, num_cols = df_data.shape

    # --------------------------------------------------------
    # 1. TIÊU ĐỀ (Tự động Gộp A1 đến Cột cuối)
    # --------------------------------------------------------
    last_col_letter = get_column_letter(num_cols)
    ws.merge_cells(f"A1:{last_col_letter}1")
    
    ws["A1"] = title_text
    ws["A1"].font = Font(name=font_name, size=13, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 42

    # --------------------------------------------------------
    # 2. DÒNG HEADER (Dòng 1 của Data -> Row 2 của Excel)
    # --------------------------------------------------------
    headers = df_data.iloc[0].fillna("").tolist()
    for col_idx, header_val in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col_idx, value=str(header_val))
        cell.font = Font(name=font_name, size=10, bold=True, color="FFFFFF" if orange_fill.fgColor.rgb == "ED7D00" else "000000")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = orange_fill
        cell.border = border
    ws.row_dimensions[2].height = 38

    # --------------------------------------------------------
    # 3. NỘI DUNG BẢNG (Dòng 2..N của Data -> Row 3..N+1 của Excel)
    # --------------------------------------------------------
    for r_idx in range(1, num_rows):
        excel_row = r_idx + 2
        ws.row_dimensions[excel_row].height = 32
        
        for c_idx in range(num_cols):
            val = df_data.iloc[r_idx, c_idx]
            val = "" if pd.isna(val) else str(val)
            
            # Thử chuyển đổi kiểu số nếu có thể
            if val.isdigit():
                val = int(val)
            else:
                try:
                    val = float(val)
                except ValueError:
                    pass

            cell = ws.cell(row=excel_row, column=c_idx + 1, value=val)
            cell.font = Font(name=font_name, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

    # --------------------------------------------------------
    # 4. TỰ ĐỘNG TÍNH VÀ CÀI ĐẶT ĐỘ RỘNG CỘT
    # --------------------------------------------------------
    for col_idx in range(1, num_cols + 1):
        col_letter = get_column_letter(col_idx)
        # Tính độ dài max của nội dung cột
        max_len = 12
        for r_idx in range(1, num_rows + 1):
            cell_val = str(ws.cell(row=r_idx, column=col_idx).value or "")
            lines = cell_val.split('\n')
            for l in lines:
                if len(l) > max_len:
                    max_len = len(l)
        ws.column_dimensions[col_letter].width = min(max_len + 5, 40)

    # --------------------------------------------------------
    # 5. CÀI ĐẶT TRANG IN CHUYÊN NGHIỆP
    # --------------------------------------------------------
    ws.sheet_view.showGridLines = True
    ws.freeze_panes = "A3"
    ws.page_setup.orientation = "landscape" if num_cols > 5 else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.2, right=0.2, top=0.3, bottom=0.3, header=0.1, footer=0.1)

    # Xuất file ra bộ nhớ RAM
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ============================================================
# GIAO DIỆN STREAMLIT
# ============================================================

uploaded_file = st.file_drop_target if hasattr(st, 'file_drop_target') else st.file_uploader(
    "📂 Tải lên File Ảnh (PNG, JPG) hoặc File Excel (XLSX, CSV)",
    type=["png", "jpg", "jpeg", "xlsx", "xls", "csv"]
)

title_input = st.text_input("📝 Tiêu đề bảng Excel:", "BẢNG TỔNG HỢP / 数据汇总表")

if uploaded_file is not None:
    file_ext = uploaded_file.name.split('.')[-1].lower()
    df_data = None
    
    with st.spinner("⏳ Đang đọc và phân tích dữ liệu bảng..."):
        if file_ext in ["png", "jpg", "jpeg"]:
            df_data = process_image_file(uploaded_file)
        elif file_ext in ["xlsx", "xls", "csv"]:
            df_data = process_excel_file(uploaded_file)

    if df_data is not None and not df_data.empty:
        rows_count, cols_count = df_data.shape
        st.success(f"✅ Đã nhận diện thành công bảng dữ liệu kích thước: **{rows_count} dòng x {cols_count} cột**")

        st.subheader("📋 Xem trước dữ liệu trích xuất")
        st.dataframe(df_data, use_container_width=True)

        st.divider()
        
        # Nút Tạo & Download
        excel_bytes = create_dynamic_excel(df_data, title_text=title_input)
        
        st.download_button(
            label="⬇️ Tải xuống File Excel Đã Định Dạng",
            data=excel_bytes.getvalue(),
            file_name=f"Bang_Xu_Ly_{rows_count}x{cols_count}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.error("❌ Không thể trích xuất dữ liệu từ file đã tải lên. Vui lòng kiểm tra lại chất lượng ảnh hoặc định dạng file.")
