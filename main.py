import copy
import io
import os
import cv2
import numpy as np
from PIL import Image
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, range_boundaries
from deep_translator import GoogleTranslator
import easyocr
import streamlit as st

st.set_page_config(
    page_title="App Dịch File Song Ngữ (Excel & Hình Ảnh)",
    page_icon="🌐",
    layout="wide",
)

st.title(
    "🌐 Ứng dụng Dịch File Song Ngữ (Hỗ trợ cả Excel & Hình Ảnh với AI OCR)"
)
st.markdown("---")

st.sidebar.header("Cấu hình chiều dịch")
direction = st.sidebar.selectbox(
    "Chọn chế độ dịch:", ("Trung - Việt", "Việt - Trung")
)

# Sử dụng mã 'zh-CN' chuẩn xác theo danh sách hỗ trợ của deep_translator
if direction == "Trung - Việt":
    source_lang, target_lang = "zh-CN", "vi"
    st.sidebar.info(
        "📌 Chế độ: Trung - Việt (Tiếng Việt luôn nằm ngay bên dưới)."
    )
    ocr_langs = ["ch_sim", "en"]  # EasyOCR nhận diện chữ Trung giản thể
else:
    source_lang, target_lang = "vi", "zh-CN"
    st.sidebar.info(
        "📌 Chế độ: Việt - Trung (Tiếng Việt luôn nằm ngay bên dưới)."
    )
    ocr_langs = ["vi", "en"]

# Khởi tạo Translator an toàn
try:
    translator = GoogleTranslator(source=source_lang, target=target_lang)
except Exception as e:
    st.sidebar.error(f"Lỗi khởi tạo bộ dịch: {e}")


# Khởi tạo EasyOCR (có caching để load nhanh hơn trên Streamlit)
@st.cache_resource
def load_ocr_reader(langs):
    return easyocr.Reader(langs, gpu=False)


def translate_text(text):
    if text is None:
        return ""
    text_str = str(text).strip()
    if not text_str:
        return ""
    if text_str.startswith("=") or text_str.replace(".", "", 1).isdigit():
        return text_str
    try:
        return translator.translate(text_str)
    except Exception:
        return text_str


# --- TẢI FILE LÊN ---
uploaded_file = st.file_uploader(
    "Tải lên file của bạn (Excel .xlsx hoặc Hình ảnh .png/.jpg/.jpeg)",
    type=["xlsx", "png", "jpg", "jpeg"],
)

if uploaded_file is not None:
    file_extension = uploaded_file.name.split(".")[-1].lower()

    # ==========================================
    # XỬ LÝ FILE EXCEL (.xlsx)
    # ==========================================
    if file_extension == "xlsx":
        st.success(
            "Đã nhận diện file Excel! Đang tiến hành xử lý cấu trúc và dịch..."
        )

        try:
            wb = openpyxl.load_workbook(uploaded_file, data_only=False)

            with st.spinner(
                "Đang xử lý dịch, chèn dòng và bảo toàn định dạng..."
            ):
                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]

                    col_widths = {
                        col: ws.column_dimensions[col].width
                        for col in ws.column_dimensions
                    }
                    max_row = ws.max_row
                    max_col = ws.max_column

                    row_heights_bak = {}
                    for r in range(1, max_row + 1):
                        if ws.row_dimensions[r].height:
                            row_heights_bak[r] = ws.row_dimensions[r].height

                    original_merged_ranges = list(ws.merged_cells.ranges)
                    for cr in original_merged_ranges:
                        ws.unmerge_cells(str(cr))

                    for r in range(max_row, 0, -1):
                        has_content = any(
                            ws.cell(row=r, column=c).value is not None
                            for c in range(1, max_col + 1)
                        )
                        if not has_content:
                            continue

                        ws.insert_rows(r + 1, amount=1)

                        if r in row_heights_bak:
                            ws.row_dimensions[r + 1].height = row_heights_bak[r]
                            ws.row_dimensions[r].height = row_heights_bak[r]

                        for c in range(1, max_col + 1):
                            src_cell = ws.cell(row=r, column=c)
                            dst_cell = ws.cell(row=r + 1, column=c)

                            val = src_cell.value
                            dst_cell.value = val

                            if src_cell.font:
                                dst_cell.font = copy.copy(src_cell.font)
                            if src_cell.fill:
                                dst_cell.fill = copy.copy(src_cell.fill)
                            if src_cell.alignment:
                                dst_cell.alignment = copy.copy(
                                    src_cell.alignment
                                )
                            if src_cell.border:
                                dst_cell.border = copy.copy(src_cell.border)
                            dst_cell.number_format = src_cell.number_format

                            if val is not None:
                                translated_val = translate_text(val)
                                if direction == "Trung - Việt":
                                    dst_cell.value = (
                                        translated_val  # Việt ở dưới
                                    )
                                else:
                                    dst_cell.value = val  # Việt ở dưới
                                    src_cell.value = (
                                        translated_val  # Trung ở trên
                                    )

                    for cr in original_merged_ranges:
                        min_col, min_row, max_col_merged, max_row_merged = (
                            range_boundaries(str(cr))
                        )
                        new_min_row = (min_row * 2) - 1
                        new_max_row = max_row_merged * 2
                        new_range_str = (
                            f"{get_column_letter(min_col)}{new_min_row}:"
                            f"{get_column_letter(max_col_merged)}{new_max_row}"
                        )
                        try:
                            ws.merge_cells(new_range_str)
                        except Exception:
                            pass

                    for col, width in col_widths.items():
                        if width:
                            ws.column_dimensions[col].width = width

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)

            st.success("🎉 Xử lý Excel hoàn tất!")
            st.download_button(
                label="📥 Tải xuống file Excel song ngữ",
                data=output,
                file_name=f"dich_song_ngu_{uploaded_file.name}",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )

        except Exception as e:
            st.error(f"Đã xảy ra lỗi khi xử lý file Excel: {e}")

    # ==========================================
    # XỬ LÝ FILE HÌNH ẢNH (.png, .jpg, .jpeg)
    # ==========================================
    elif file_extension in ["png", "jpg", "jpeg"]:
        st.success("Đã nhận diện file Hình ảnh! Đang tiến hành trích xuất chữ OCR...")

        try:
            # Đọc ảnh gốc bằng OpenCV / PIL
            image = Image.open(uploaded_file)
            image_np = np.array(image)

            # Hiển thị ảnh gốc
            st.image(image, caption="Hình ảnh gốc", use_container_width=True)

            with st.spinner("AI đang đọc chữ (OCR) và dịch song ngữ..."):
                reader = load_ocr_reader(ocr_langs)
                # Kết quả OCR trả về các tuple: (bbox, text, probability)
                results = reader.readtext(image_np)

            if not results:
                st.warning(
                    "Không tìm thấy văn bản nào trong hình ảnh này. Hãy thử ảnh"
                    " khác rõ nét hơn."
                )
            else:
                st.subheader("📝 Kết quả trích xuất & Dịch song ngữ:")

                # Hiển thị dạng bảng kết quả chi tiết theo từng dòng nhận diện
                extracted_data = []
                for bbox, text, prob in results:
                    translated = translate_text(text)

                    # Sắp xếp hiển thị: Tiếng Việt luôn nằm ngay dưới Tiếng Trung
                    if direction == "Trung - Việt":
                        line_top = f"🀄 Gốc (Trung): {text}"
                        line_bottom = f"🇻🇳 Dịch (Việt): {translated}"
                    else:
                        line_top = f"🀄 Dịch (Trung): {translated}"
                        line_bottom = f"🇻🇳 Gốc (Việt): {text}"

                    extracted_data.append(
                        {
                            "Văn bản gốc": text,
                            "Bản dịch": translated,
                            "Độ chính xác": f"{prob * 100:.1f}%",
                        }
                    )

                    # In ra giao diện đúng yêu cầu: Nội dung tiếng Việt nằm ngay bên dưới
                    st.markdown(
                        f"""
                        <div style="background-color: #f0f2f6; padding: 10px; border-radius: 5px; margin-bottom: 8px;">
                            <div style="font-weight: bold; color: #333;">{line_top}</div>
                            <div style="font-weight: bold; color: #0066cc; margin-top: 4px;">{line_bottom}</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                # Cung cấp file text/csv để download kết quả dịch hình ảnh nếu cần
                import pandas as pd

                df_result = pd.DataFrame(extracted_data)
                csv_data = df_result.to_csv(index=False).encode("utf-8")

                st.download_button(
                    label="📥 Tải xuống kết quả dịch hình ảnh (CSV)",
                    data=csv_data,
                    file_name=f"ket_qua_dich_anh_{uploaded_file.name}.csv",
                    mime="text/csv",
                )

        except Exception as e:
            st.error(f"Đã xảy ra lỗi trong quá trình xử lý hình ảnh: {e}")
