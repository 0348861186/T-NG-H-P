import io
import os
from PIL import Image, ImageDraw, ImageFont
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from deep_translator import GoogleTranslator
import streamlit as st
import easyocr

# Cấu hình giao diện Streamlit
st.set_page_config(
    page_title="App Dịch Song Ngữ (Excel & Hình Ảnh)", layout="wide"
)


# Khởi tạo EasyOCR (Fix lỗi xung đột ngôn ngữ bằng cách tách riêng hoặc dùng chung chuẩn của EasyOCR)
@st.cache_resource
def load_ocr_readers():
    # EasyOCR yêu cầu tiếng Trung chỉ đi kèm với tiếng Anh, do đó ta khởi tạo 2 reader riêng biệt để quét linh hoạt cả Trung lẫn Việt
    reader_zh = easyocr.Reader(["ch_sim", "en"], gpu=False)
    reader_vi = easyocr.Reader(["vi", "en"], gpu=False)
    return reader_zh, reader_vi


reader_zh, reader_vi = load_ocr_readers()


# Hàm dịch văn bản sử dụng deep-translator
def translate_text(text, direction):
    if not text or not str(text).strip():
        return text

    try:
        if direction == "Trung - Việt":
            translated = GoogleTranslator(
                source="zh-CN", target="vi"
            ).translate(str(text))
        else:  # Việt - Trung
            translated = GoogleTranslator(
                source="vi", target="zh-CN"
            ).translate(str(text))
        return translated
    except Exception as e:
        return f"[Lỗi dịch: {str(e)}]"


# Xử lý dịch file Excel
def process_excel(file, direction):
    wb = openpyxl.load_workbook(file)
    output_wb = openpyxl.Workbook()
    output_wb.remove(output_wb.active)  # Xóa sheet mặc định

    for sheet_name in wb.sheetnames:
        sheet = wb[sheetname]
        new_sheet = output_wb.create_sheet(title=sheet_name)

        max_row = sheet.max_row
        max_col = sheet.max_column

        # Copy độ rộng cột
        for col in sheet.columns:
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            new_sheet.column_dimensions[col_letter].width = (
                sheet.column_dimensions[col_letter].width
            )

        # Copy chiều cao hàng
        for row_idx in range(1, max_row + 1):
            if sheet.row_dimensions[row_idx].height:
                new_sheet.row_dimensions[row_idx].height = sheet.row_dimensions[
                    row_idx
                ].height

        # Copy merged cells
        for merged_cell in sheet.merged_cells.ranges:
            new_sheet.merge_cells(str(merged_cell))

        # Duyệt qua từng ô để dịch và ghép nội dung (Tiếng Việt luôn nằm dưới)
        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                cell = sheet.cell(row=row, column=col)
                new_cell = new_sheet.cell(row=row, column=col)

                val = cell.value
                if val is not None and str(val).strip() != "":
                    translated_val = translate_text(val, direction)

                    if direction == "Trung - Việt":
                        combined_val = f"{val}\n{translated_val}"
                    else:
                        combined_val = f"{translated_val}\n{val}"

                    new_cell.value = combined_val
                else:
                    new_cell.value = val

                # Copy định dạng (Style)
                if cell.has_style:
                    new_cell.font = copy_font(cell.font)
                    new_cell.alignment = copy_alignment(cell.alignment)
                    new_cell.fill = copy_fill(cell.fill)
                    new_cell.border = copy_border(cell.border)
                    if new_cell.alignment:
                        new_cell.alignment = Alignment(
                            horizontal=new_cell.alignment.horizontal,
                            vertical=new_cell.alignment.vertical,
                            wrap_text=True,
                            indent=new_cell.alignment.indent,
                            shrink_to_fit=new_cell.alignment.shrink_to_fit,
                        )
                    else:
                        new_cell.alignment = Alignment(wrap_text=True)

    output = io.BytesIO()
    output_wb.save(output)
    output.seek(0)
    return output


# Các hàm hỗ trợ copy style của openpyxl
def copy_font(font):
    if not font:
        return None
    return Font(
        name=font.name,
        size=font.size,
        bold=font.bold,
        italic=font.italic,
        vertAlign=font.vertAlign,
        underline=font.underline,
        strike=font.strike,
        color=font.color,
    )


def copy_alignment(alignment):
    if not alignment:
        return Alignment(wrap_text=True)
    return Alignment(
        horizontal=alignment.horizontal,
        vertical=alignment.vertical,
        text_rotation=alignment.text_rotation,
        wrap_text=True,
        shrink_to_fit=alignment.shrink_to_fit,
        indent=alignment.indent,
    )


def copy_fill(fill):
    if not fill:
        return None
    return PatternFill(
        fill_type=fill.fill_type,
        start_color=fill.start_color,
        end_color=fill.end_color,
    )


def copy_border(border):
    if not border:
        return None
    return Border(
        left=border.left,
        right=border.right,
        top=border.top,
        bottom=border.bottom,
        diagonal=border.diagonal,
        diagonal_direction=border.diagonal_direction,
        outline=border.outline,
    )


# Xử lý dịch hình ảnh
def process_image(image_file, direction):
    image = Image.open(image_file).convert("RGB")
    img_np = np.array(image)

    # Dùng cả 2 reader để quét chữ chính xác cho cả tiếng Trung và tiếng Việt
    results_zh = reader_zh.readtext(img_np)
    results_vi = reader_vi.readtext(img_np)
    results = results_zh + results_vi  # Gộp kết quả nhận diện

    draw_image = image.copy()
    draw = ImageDraw.Draw(draw_image)

    font_path = None
    for path in [
        "C:/Windows/Fonts/arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    ]:
        if os.path.exists(path):
            font_path = path
            break

    for bbox, text, prob in results:
        if prob < 0.2:
            continue

        translated = translate_text(text, direction)

        if direction == "Trung - Việt":
            display_text = f"{text}\n{translated}"
        else:
            display_text = f"{translated}\n{text}"

        top_left = bbox[0]
        bottom_left = bbox[3]
        bottom_right = bbox[2]

        box_height = int(bottom_left[1] - top_left[1])
        font_size = max(12, int(box_height * 0.4))
        try:
            font = (
                ImageFont.truetype(font_path, font_size)
                if font_path
                else ImageFont.load_default()
            )
        except:
            font = ImageFont.load_default()

        text_x = int(top_left[0])
        text_y = int(bottom_left[1] + 2)

        draw.text(
            (text_x, text_y), display_text, fill=(255, 0, 0), font=font
        )

    return draw_image


# --- Giao diện Streamlit Dashboard ---
st.title("🔤 Ứng Dụng Dịch Song Ngữ Thông Minh (Giữ Nguyên Định Dạng)")
st.markdown(
    "Hỗ trợ dịch file **Excel (.xlsx)** giữ nguyên 100% định dạng và file **Hình ảnh**. Tiếng Việt luôn nằm ngay bên dưới tiếng Trung."
)

# Sidebar cấu hình
st.sidebar.header("⚙️ Tùy Chọn Cấu Hình")
direction = st.sidebar.selectbox(
    "Chọn chiều dịch:", ("Trung - Việt", "Việt - Trung")
)

# Khu vực tải file lên
uploaded_file = st.file_uploader(
    "Tải lên file của bạn (Hình ảnh hoặc Excel)", type=["xlsx", "png", "jpg", "jpeg"]
)

if uploaded_file is not None:
    file_extension = uploaded_file.name.split(".")[-1].lower()

    if file_extension == "xlsx":
        st.info(
            "📁 Đã nhận file Excel. Đang tiến hành dịch và giữ nguyên định dạng..."
        )

        if st.button("🚀 Bắt đầu dịch Excel"):
            with st.spinner("Đang xử lý cấu trúc và dịch nội dung..."):
                try:
                    output_excel = process_excel(uploaded_file, direction)
                    st.success("✨ Dịch file Excel hoàn tất!")

                    st.download_button(
                        label="📥 Tải xuống file Excel đã dịch",
                        data=output_excel,
                        file_name=f"dich_song_ngu_{uploaded_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi trong quá trình xử lý file: {e}")

    elif file_extension in ["png", "jpg", "jpeg"]:
        st.info("🖼️ Đã nhận file hình ảnh. Đang nhận diện và dịch văn bản...")

        if st.button("🚀 Bắt đầu dịch Hình Ảnh"):
            with st.spinner("Đang quét OCR và dịch..."):
                try:
                    result_img = process_image(uploaded_file, direction)
                    st.success("✨ Dịch hình ảnh hoàn tất!")

                    st.image(
                        result_img,
                        caption="Ảnh sau khi dịch (Tiếng Việt nằm dưới Tiếng Trung)",
                        use_column_width=True,
                    )

                    buf = io.BytesIO()
                    result_img.save(buf, format="PNG")
                    byte_im = buf.getvalue()

                    st.download_button(
                        label="📥 Tải xuống hình ảnh đã dịch",
                        data=byte_im,
                        file_name=f"dich_{uploaded_file.name}",
                        mime="image/png",
                    )
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi trong quá trình xử lý ảnh: {e}")
