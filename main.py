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
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="App Dịch File Song Ngữ (Excel & Hình Ảnh -> Excel)",
    page_icon="🌐",
    layout="wide",
)

st.title("🌐 Ứng dụng Dịch File Song Ngữ & Xuất File Excel")
st.markdown("---")

st.sidebar.header("Cấu hình chiều dịch")
direction = st.sidebar.selectbox(
    "Chọn chế độ dịch:", ("Trung - Việt", "Việt - Trung")
)

if direction == "Trung - Việt":
    source_lang, target_lang = "zh-CN", "vi"
    st.sidebar.info(
        "📌 Chế độ: Trung - Việt (Tiếng Việt luôn nằm ngay bên dưới)."
    )
    ocr_langs = ["ch_sim", "en"]
else:
    source_lang, target_lang = "vi", "zh-CN"
    st.sidebar.info(
        "📌 Chế độ: Việt - Trung (Tiếng Việt luôn nằm ngay bên dưới)."
    )
    ocr_langs = ["vi", "en"]

# Khởi tạo Translator
try:
    translator = GoogleTranslator(source=source_lang, target=target_lang)
except Exception as e:
    st.sidebar.error(f"Lỗi khởi tạo bộ dịch: {e}")


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


uploaded_file = st.file_uploader(
    "Tải lên file của bạn (Excel .xlsx hoặc Hình ảnh .png/.jpg/.jpeg)",
    type=["xlsx", "png", "jpg", "jpeg"],
)

if uploaded_file is not None:
    file_extension = uploaded_file.name.split(".")[-1].lower()

    # ==========================================
    # 1. XỬ LÝ FILE EXCEL (.xlsx)
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
                                    dst_cell.value = translated_val
                                else:
                                    dst_cell.value = val
                                    src_cell.value = translated_val

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
    # 2. XỬ LÝ FILE HÌNH ẢNH (.png, .jpg, .jpeg) -> XUẤT RA EXCEL
    # ==========================================
    elif file_extension in ["png", "jpg", "jpeg"]:
        st.success(
            "Đã nhận diện file Hình ảnh! Đang trích xuất OCR và tạo file"
            " Excel..."
        )

        try:
            image = Image.open(uploaded_file)
            image_np = np.array(image)

            st.image(image, caption="Hình ảnh gốc", use_container_width=True)

            with st.spinner("AI đang đọc chữ (OCR) và biên tập file Excel..."):
                reader = load_ocr_reader(ocr_langs)
                results = reader.readtext(image_np)

            if not results:
                st.warning(
                    "Không tìm thấy văn bản nào trong hình ảnh này. Hãy thử ảnh"
                    " khác rõ nét hơn."
                )
            else:
                # Tạo một file Excel mới bằng openpyxl để lưu kết quả OCR song ngữ
                wb_out = openpyxl.Workbook()
                ws_out = wb_out.active
                ws_out.title = "KetQua_OCR_SongNgu"

                # Thiết lập tiêu đề bảng Excel
                ws_out["A1"] = "STT"
                ws_out["B1"] = (
                    "Nội dung Gốc (Tiếng Trung)"
                    if direction == "Trung - Việt"
                    else "Nội dung Gốc (Tiếng Việt)"
                )
                ws_out["C1"] = (
                    "Bản Dịch (Tiếng Việt)"
                    if direction == "Trung - Việt"
                    else "Bản Dịch (Tiếng Trung)"
                )
                ws_out["D1"] = "Độ chính xác OCR"

                # Định dạng Header
                header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
                header_fill = PatternFill(
                    start_color="4F81BD", end_color="4F81BD", fill_type="solid"
                )
                thin_border = Border(
                    left=Side(style="thin", color="CCCCCC"),
                    right=Side(style="thin", color="CCCCCC"),
                    top=Side(style="thin", color="CCCCCC"),
                    bottom=Side(style="thin", color="CCCCCC"),
                )

                for col_idx in range(1, 5):
                    cell = ws_out.cell(row=1, column=col_idx)
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center"
                    )
                    cell.border = thin_border

                row_idx = 2
                for idx, (bbox, text, prob) in enumerate(results, 1):
                    translated = translate_text(text)

                    # Ghi dòng gốc và dòng dịch ngay bên dưới (hoặc theo cặp dòng trong Excel)
                    # Theo yêu cầu: tiếng Việt nằm ngay bên dưới tiếng Trung
                    if direction == "Trung - Việt":
                        line_upper = text  # Trung trên
                        line_lower = translated  # Việt dưới
                    else:
                        line_upper = translated  # Trung trên
                        line_lower = text  # Việt dưới

                    # Dòng 1: Tiếng Trung
                    ws_out.cell(row=row_idx, column=1, value=idx)
                    ws_out.cell(row=row_idx, column=2, value=line_upper)
                    ws_out.cell(row=row_idx, column=3, value="")
                    ws_out.cell(
                        row=row_idx, column=4, value=f"{prob * 100:.1f}%"
                    )

                    # Dòng 2: Tiếng Việt nằm ngay bên dưới
                    ws_out.cell(row=row_idx + 1, column=1, value="")
                    ws_out.cell(row=row_idx + 1, column=2, value="")
                    ws_out.cell(row=row_idx + 1, column=3, value=line_lower)
                    ws_out.cell(row=row_idx + 1, column=4, value="")

                    # Style các ô
                    for r in [row_idx, row_idx + 1]:
                        for c in range(1, 5):
                            cell = ws_out.cell(row=r, column=c)
                            cell.border = thin_border
                            if c == 3 and r == row_idx + 1:
                                cell.font = Font(
                                    name="Arial", size=10, bold=True, color="0066CC"
                                >
                            else:
                                cell.font = Font(name="Arial", size=10)

                    row_idx += 2

                # Tự động điều chỉnh độ rộng cột cho đẹp
                for col in ws_out.columns:
                    max_len = max(len(str(cell.value or "")) for cell in col)
                    col_letter = get_column_letter(col[0].column)
                    ws_out.column_dimensions[col_letter].width = max(
                        max_len + 4, 15
                    )

                # Lưu vào bộ nhớ đệm
                output_excel = io.BytesIO()
                wb_out.save(output_excel)
                output_excel.seek(0)

                st.success(
                    "🎉 Đã trích xuất hình ảnh và biên tập thành file Excel thành"
                    " công!"
                )

                # Nút download file Excel
                st.download_button(
                    label="📥 Tải xuống file Excel kết quả dịch hình ảnh",
                    data=output_excel,
                    file_name=f"dich_anh_sang_excel_{uploaded_file.name}.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ),
                )

        except Exception as e:
            st.error(f"Đã xảy ra lỗi trong quá trình xử lý hình ảnh: {e}")
