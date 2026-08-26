import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from PIL import Image
import streamlit as st
from deep_translator import GoogleTranslator

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Dịch Bảng Song Ngữ Trung - Việt",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Ứng Dụng Dịch Bảng Song Ngữ (Trung ↔ Việt)")
st.write(
    "Tải lên file **Excel** hoặc **File Ảnh**. Hệ thống sẽ dịch và trình bày dạng song ngữ (Dòng tiếng Việt/Trung dịch nằm ngay bên dưới dòng gốc)."
)


# Cache bộ đọc OCR để tối ưu RAM Streamlit Cloud
@st.cache_resource
def load_ocr_reader(chinese_type):
    import easyocr

    lang_code = "ch_sim" if chinese_type == "Giản thể (ch_sim)" else "ch_tra"
    return easyocr.Reader([lang_code, "en"], download_enabled=True)


# Hàm dịch văn bản
def translate_text(text, src_lang, target_lang):
    if not text or pd.isna(text):
        return ""
    text_str = str(text).strip()

    if not text_str or text_str.isdigit():
        return text_str

    try:
        translated = GoogleTranslator(
            source=src_lang, target=target_lang
        ).translate(text_str)
        return translated if translated else text_str
    except Exception:
        return text_str


# Tạo nội dung song ngữ: Dòng gốc \n Dòng dịch
def process_bilingual_cell(original_val, src_lang, target_lang):
    if not original_val or pd.isna(original_val):
        return ""

    val_str = str(original_val).strip()

    # Nếu không phải chữ cái hay chữ Hán thì giữ nguyên (ví dụ: số 1, 2, 3...)
    if not re.search(r"[\u4e00-\u9fff a-zA-Zà-ỹÀ-Ỹ]", val_str):
        return val_str

    translated_val = translate_text(val_str, src_lang, target_lang)

    if val_str.lower() == translated_val.lower():
        return val_str

    # Ghép dòng gốc và dòng dịch bằng ký tự xuống dòng
    return f"{val_str}\n{translated_val}"


# Tạo File Excel chuẩn định dạng xuống dòng và viền bảng
def create_styled_excel(df_bilingual):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bảng Song Ngữ"

    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    header_fill = PatternFill(
        start_color="E67E22", end_color="E67E22", fill_type="solid"
    )
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    cell_font = Font(name="Arial", size=10)
    align_center = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )

    for r_idx, row in enumerate(
        df_bilingual.itertuples(index=False), start=1
    ):
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_border
            cell.alignment = align_center

            if r_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
            else:
                cell.font = cell_font

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                lines = str(cell.value).split("\n")
                for line in lines:
                    if len(line) > max_len:
                        max_len = len(line)
        ws.column_dimensions[col_letter].width = max(max_len + 6, 14)

    for row in ws.iter_rows():
        ws.row_dimensions[row[0].row].height = 40

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# Hàm hiển thị Bảng HTML trên Streamlit hỗ trợ xuống dòng \n chuẩn
def display_html_table(df):
    html = '<table style="width:100%; border-collapse: collapse; text-align: center; font-family: Arial;">'
    for r_idx, row in enumerate(df.values):
        html += "<tr>"
        for val in row:
            val_str = (
                str(val).replace("\n", "<br>")
                if pd.notna(val) and val != ""
                else ""
            )
            if r_idx == 0:
                html += f'<th style="border: 1px solid #333; padding: 8px; background-color: #E67E22; color: white; font-weight: bold;">{val_str}</th>'
            else:
                html += f'<td style="border: 1px solid #ccc; padding: 8px;">{val_str}</td>'
        html += "</tr>"
    html += "</table>"
    st.markdown(html, unsafe_allow_html=True)


# ---------------------------------------------------------
# STREAMLIT UI
# ---------------------------------------------------------
st.sidebar.header("⚙️ Cấu hình Dịch")
direction = st.sidebar.selectbox(
    "Chọn hướng dịch:",
    [
        "Trung sang Việt (Chinese -> Vietnamese)",
        "Việt sang Trung (Vietnamese -> Chinese)",
    ],
)

chinese_type = st.sidebar.selectbox(
    "Loại chữ Trung trong ảnh (dành cho file Ảnh):",
    ["Giản thể (ch_sim)", "Phồn thể (ch_tra)"],
)

src_lang = "zh-CN" if "Trung sang Việt" in direction else "vi"
target_lang = "vi" if "Trung sang Việt" in direction else "zh-CN"

uploaded_file = st.file_uploader(
    "Tải lên file Excel (.xlsx, .xls) hoặc File Ảnh (.png, .jpg, .jpeg)",
    type=["xlsx", "xls", "png", "jpg", "jpeg"],
)

if uploaded_file is not None:
    file_type = uploaded_file.name.split(".")[-1].lower()
    df_raw = None

    if file_type in ["xlsx", "xls"]:
        st.info("📂 Đã nhận file Excel.")
        df_raw = pd.read_excel(uploaded_file, header=None)

    elif file_type in ["png", "jpg", "jpeg"]:
        st.info("🖼️ Đã nhận file Ảnh. Đang quét chữ...")
        try:
            import numpy as np

            image = Image.open(uploaded_file)
            st.image(image, caption="Ảnh gốc tải lên", use_container_width=True)

            reader = load_ocr_reader(chinese_type)
            img_np = np.array(image)
            results = reader.readtext(img_np)

            if results:
                sorted_res = sorted(results, key=lambda x: x[0][0][1])
                lines = []
                curr_line = []
                last_y = -1

                for bbox, text, prob in sorted_res:
                    y_top = bbox[0][1]
                    if last_y == -1 or abs(y_top - last_y) < 22:
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
                st.error("Không phát hiện được chữ trong ảnh!")
        except Exception as e:
            st.error(f"Lỗi khi xử lý ảnh: {e}")

    if df_raw is not None and not df_raw.empty:
        st.subheader("📋 Bảng dữ liệu gốc nhận diện được")
        st.dataframe(df_raw)

        if st.button("🚀 Tiến hành Dịch Song Ngữ"):
            with st.spinner("Đang dịch dữ liệu..."):
                df_bilingual = df_raw.copy()

                for r in range(df_raw.shape[0]):
                    for c in range(df_raw.shape[1]):
                        cell_val = df_raw.iloc[r, c]
                        df_bilingual.iloc[r, c] = process_bilingual_cell(
                            cell_val, src_lang, target_lang
                        )

                st.success("✅ Dịch hoàn tất!")

                st.subheader("✨ Kết quả hiển thị Song Ngữ (Xem trước):")
                # Hiển thị trực tiếp bảng song ngữ xuống dòng dạng HTML
                display_html_table(df_bilingual)

                st.write("")
                # Tải file Excel có định dạng chuẩn
                excel_bytes = create_styled_excel(df_bilingual)
                st.download_button(
                    label="📥 Tải xuống File Excel Song Ngữ",
                    data=excel_bytes,
                    file_name="Bang_Dich_Song_Ngu.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
