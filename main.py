import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from PIL import Image
import streamlit as st
from deep_translator import GoogleTranslator

# Tùy chỉnh cấu hình trang Streamlit
st.set_page_config(
    page_title="Dịch Bảng Song Ngữ Trung - Việt",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Ứng Dụng Dịch Bảng Song Ngữ (Trung ↔ Việt)")
st.write(
    "Tải lên file **Excel** hoặc **Ảnh bảng biểu**. Hệ thống sẽ dịch và trình bày dạng song ngữ (Dòng dịch nằm ngay dưới dòng gốc trong cùng 1 ô)."
)


# Hàng hàm bổ trợ dịch thuật
def translate_text(text, src_lang, target_lang):
    """Dịch chuỗi văn bản nếu chứa chữ cái hoặc chữ Hán."""
    if not text or pd.isna(text):
        return ""
    text_str = str(text).strip()
    # Nếu là số pure hoặc rỗng thì không dịch
    if not text_str or text_str.isdigit():
        return text_str

    try:
        translated = GoogleTranslator(
            source=src_lang, target=target_lang
        ).translate(text_str)
        return translated if translated else text_str
    except Exception:
        return text_str


def process_bilingual_cell(original_val, src_lang, target_lang):
    """Tạo nội dung song ngữ cho ô: Dòng gốc / Dòng dịch."""
    if not original_val or pd.isna(original_val):
        return ""

    val_str = str(original_val).strip()

    # Nếu chỉ có số hoặc không chứa chữ thì giữ nguyên
    if not re.search(r"[\u4e00-\u9fff a-zA-Zà-ỹÀ-Ỹ]", val_str):
        return val_str

    translated_val = translate_text(val_str, src_lang, target_lang)

    # Nếu kết quả dịch giống hệt câu gốc thì trả về câu gốc
    if val_str.lower() == translated_val.lower():
        return val_str

    return f"{val_str}\n{translated_val}"


# Tạo file Excel với định dạng chuẩn từ DataFrame song ngữ
def create_styled_excel(df_bilingual):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bảng Song Ngữ"

    # Style cơ bản
    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    header_fill = PatternFill(
        start_color="E67E22", end_color="E67E22", fill_type="solid"
    )  # Màu cam như mẫu
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    cell_font = Font(name="Arial", size=10)
    align_center = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )

    # Ghi dữ liệu vào sheet
    for r_idx, row in enumerate(
        df_bilingual.itertuples(index=False), start=1
    ):
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_border
            cell.alignment = align_center

            # Nếu là dòng tiêu đề (Dòng 1)
            if r_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
            else:
                cell.font = cell_font

    # Tự động chỉnh độ rộng cột và chiều cao dòng
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                lines = str(cell.value).split("\n")
                for l in lines:
                    if len(l) > max_len:
                        max_len = len(l)
        ws.column_dimensions[col_letter].width = max(max_len + 5, 12)

    for row in ws.iter_rows():
        ws.row_dimensions[row[0].row].height = 35  # Độ cao tối thiểu cho 2 dòng

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# Cấu hình thanh điều hướng bên trái (Sidebar)
st.sidebar.header("⚙️ Cấu hình Dịch")
direction = st.sidebar.selectbox(
    "Chọn hướng dịch:",
    [
        "Trung sang Việt (Chinese -> Vietnamese)",
        "Việt sang Trung (Vietnamese -> Chinese)",
    ],
)

src_lang = "zh-CN" if "Trung sang Việt" in direction else "vi"
target_lang = "vi" if "Trung sang Việt" in direction else "zh-CN"

# Tải file lên
uploaded_file = st.file_uploader(
    "Tải lên file Excel (.xlsx, .xls) hoặc File Ảnh (.png, .jpg, .jpeg)",
    type=["xlsx", "xls", "png", "jpg", "jpeg"],
)

if uploaded_file is not None:
    file_type = uploaded_file.name.split(".")[-1].lower()

    df_raw = None

    if file_type in ["xlsx", "xls"]:
        st.info("📂 Đã nhận file Excel. Đang đọc dữ liệu...")
        df_raw = pd.read_excel(uploaded_file, header=None)

    elif file_type in ["png", "jpg", "jpeg"]:
        st.info("🖼️ Đã nhận file Ảnh. Đang quét chữ bằng OCR...")
        try:
            import easyocr
            import numpy as np

            # Đọc ảnh
            image = Image.open(uploaded_file)
            st.image(image, caption="Ảnh gốc tải lên", use_container_width=True)

            # Khởi tạo OCR
            reader = easyocr.Reader(["zh_sim", "en"])
            img_np = np.array(image)
            results = reader.readtext(img_np)

            # Gom nhóm tọa độ để tạo cấu trúc dòng/cột đơn giản
            if results:
                # Sắp xếp kết quả theo trục Y (dòng) rồi tới X (cột)
                lines = []
                # Phân nhóm dòng dựa vào khoảng cách Y
                sorted_res = sorted(results, key=lambda x: x[0][0][1])
                curr_line = []
                last_y = -1

                for bbox, text, prob in sorted_res:
                    y_top = bbox[0][1]
                    if last_y == -1 or abs(y_top - last_y) < 20:
                        curr_line.append((bbox[0][0], text))
                    else:
                        curr_line.sort(key=lambda x: x[0])
                        lines.append([t[1] for t in curr_line])
                        curr_line = [(bbox[0][0], text)]
                    last_y = y_top

                if curr_line:
                    curr_line.sort(key=lambda x: x[0])
                    lines.append([t[1] for t in curr_line])

                df_raw = pd.DataFrame(lines)
            else:
                st.error("Không nhận diện được văn bản trong ảnh!")
        except ImportError:
            st.error(
                "Để xử lý ảnh trên Streamlit Cloud, hãy chắc chắn bạn đã thêm `easyocr` vào file `requirements.txt`."
            )

    # Tiến hành xử lý dịch bảng dữ liệu
    if df_raw is not None and not df_raw.empty:
        st.subheader("📋 Bảng dữ liệu gốc")
        st.dataframe(df_raw)

        if st.button("🚀 Tiến hành Dịch Song Ngữ"):
            with st.spinner("Đang dịch toàn bộ bảng... Vui lòng đợi trong giây lát."):
                # Tạo DataFrame mới lưu kết quả dịch
                df_bilingual = df_raw.copy()

                for row_idx in range(df_raw.shape[0]):
                    for col_idx in range(df_raw.shape[1]):
                        cell_val = df_raw.iloc[row_idx, col_idx]
                        df_bilingual.iloc[row_idx, col_idx] = (
                            process_bilingual_cell(
                                cell_val, src_lang, target_lang
                            )
                        )

                st.success("✅ Dịch hoàn tất!")

                st.subheader("✨ Kết quả xem trước:")
                st.dataframe(df_bilingual)

                # Xuất ra file Excel
                excel_bytes = create_styled_excel(df_bilingual)

                # Nút Download File Excel
                st.download_button(
                    label="📥 Tải xuống File Excel Song Ngữ",
                    data=excel_bytes,
                    file_name="Bảng_Dịch_Song_Ngữ.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
