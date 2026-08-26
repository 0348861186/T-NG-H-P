import io
import os
import re
import tempfile
from pathlib import Path

import streamlit as st
import numpy as np
import cv2
from PIL import Image, ImageDraw, ImageFont

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

from deep_translator import GoogleTranslator


# ============================================================
# CẤU HÌNH STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Trình dịch Excel & Ảnh Trung - Việt",
    page_icon="🌐",
    layout="wide"
)


# ============================================================
# GIAO DIỆN
# ============================================================

st.title("🌐 Trình dịch Excel & Ảnh Trung - Việt")

st.caption(
    "Tự nhận diện file Excel / hình ảnh • OCR • dịch tự động • "
    "giữ nguyên bố cục tối đa"
)


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.header("⚙️ Cài đặt")

translation_mode = st.sidebar.radio(
    "Kiểu hiển thị",
    [
        "Song ngữ - Trung + Việt",
        "Chỉ tiếng Việt"
    ]
)

source_language = st.sidebar.selectbox(
    "Ngôn ngữ nguồn",
    [
        "Tiếng Trung",
        "Tự động nhận diện"
    ]
)

target_language = st.sidebar.selectbox(
    "Ngôn ngữ đích",
    [
        "Tiếng Việt",
        "English"
    ]
)

st.sidebar.divider()

st.sidebar.info(
    "Excel sẽ được xử lý trực tiếp trên workbook gốc để cố gắng "
    "giữ nguyên định dạng."
)


# ============================================================
# BỘ NHỚ CACHE TRANSLATION
# ============================================================

if "translation_cache" not in st.session_state:
    st.session_state.translation_cache = {}


# ============================================================
# HÀM KIỂM TRA CHUỖI CÓ CHỮ
# ============================================================

def contains_text(value):
    if value is None:
        return False

    if not isinstance(value, str):
        return False

    value = value.strip()

    if not value:
        return False

    return True


# ============================================================
# LOẠI BỎ NHỮNG CHUỖI KHÔNG NÊN DỊCH
# ============================================================

def should_translate(text):
    if not contains_text(text):
        return False

    text = str(text).strip()

    # URL
    if re.match(r"^(https?|ftp)://", text, re.I):
        return False

    # Email
    if re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", text):
        return False

    # Chỉ là số
    if re.match(r"^[\d\s.,:/+\-()%]+$", text):
        return False

    # Công thức Excel
    if text.startswith("="):
        return False

    return True


# ============================================================
# DỊCH
# ============================================================

def translate_text(text, target="vi"):
    """
    Dịch text bằng Google Translate thông qua deep-translator.
    Có cache để không dịch lặp.
    """

    if not should_translate(text):
        return text

    original = str(text)

    cache_key = (
        original,
        target
    )

    if cache_key in st.session_state.translation_cache:
        return st.session_state.translation_cache[cache_key]

    try:
        translator = GoogleTranslator(
            source="auto",
            target=target
        )

        result = translator.translate(original)

        if result is None:
            result = original

        result = str(result).strip()

        st.session_state.translation_cache[cache_key] = result

        return result

    except Exception:
        return original


# ============================================================
# TẠO NỘI DUNG SONG NGỮ
# ============================================================

def make_translated_text(original, translated):
    if translation_mode == "Chỉ tiếng Việt":
        return translated

    if original.strip() == translated.strip():
        return original

    return f"{original}\n{translated}"


# ============================================================
# PHÂN TÍCH EXTENSION
# ============================================================

def detect_file_type(filename):
    ext = Path(filename).suffix.lower()

    excel_extensions = {
        ".xlsx",
        ".xlsm",
        ".xltx",
        ".xltm",
        ".xls"
    }

    image_extensions = {
        ".png",
        ".jpg",
        ".jpeg",
        ".bmp",
        ".tif",
        ".tiff",
        ".webp"
    }

    if ext in excel_extensions:
        return "excel"

    if ext in image_extensions:
        return "image"

    return "unknown"


# ============================================================
# ============================================================
#                    EXCEL PROCESSING
# ============================================================
# ============================================================


# ============================================================
# DỊCH EXCEL
# ============================================================

def translate_excel(uploaded_file):
    """
    Mở workbook gốc và dịch trực tiếp các cell chứa text.
    Không tạo workbook mới.
    """

    file_bytes = uploaded_file.getvalue()

    filename = uploaded_file.name
    ext = Path(filename).suffix.lower()

    keep_vba = ext in {".xlsm", ".xltm"}

    input_stream = io.BytesIO(file_bytes)

    # --------------------------------------------------------
    # XLSX / XLSM / XLTX / XLTM
    # --------------------------------------------------------

    if ext != ".xls":
        wb = load_workbook(
            input_stream,
            keep_vba=keep_vba,
            data_only=False
        )

    else:
        raise ValueError(
            "File .xls đời cũ không được openpyxl hỗ trợ trực tiếp. "
            "Hãy lưu file thành .xlsx trước khi upload."
        )

    translated_count = 0

    progress_bar = st.progress(0)

    # Đếm tổng số ô cần xử lý
    total_cells = 0

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:

                if isinstance(cell, MergedCell):
                    continue

                if should_translate(cell.value):
                    total_cells += 1

    processed = 0

    # --------------------------------------------------------
    # DỊCH TỪNG SHEET
    # --------------------------------------------------------

    for ws in wb.worksheets:

        st.write(f"🔄 Đang xử lý sheet: **{ws.title}**")

        for row in ws.iter_rows():

            for cell in row:

                if isinstance(cell, MergedCell):
                    continue

                value = cell.value

                if not should_translate(value):
                    continue

                original = str(value)

                translated = translate_text(
                    original,
                    target=(
                        "vi"
                        if target_language == "Tiếng Việt"
                        else "en"
                    )
                )

                new_value = make_translated_text(
                    original,
                    translated
                )

                cell.value = new_value

                # Giữ wrap text nếu có chữ song ngữ
                if "\n" in new_value:
                    current_alignment = cell.alignment.copy(
                        wrap_text=True
                    )
                    cell.alignment = current_alignment

                translated_count += 1
                processed += 1

                if total_cells > 0:
                    progress_bar.progress(
                        min(
                            processed / total_cells,
                            1.0
                        )
                    )

    progress_bar.progress(1.0)

    # --------------------------------------------------------
    # GHI RA RAM
    # --------------------------------------------------------

    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    return output, translated_count


# ============================================================
# ============================================================
#                    OCR IMAGE PROCESSING
# ============================================================
# ============================================================


@st.cache_resource
def load_easyocr():
    """
    EasyOCR được load một lần.
    """

    import easyocr

    reader = easyocr.Reader(
        [
            "ch_sim",
            "en"
        ],
        gpu=False
    )

    return reader


# ============================================================
# OCR
# ============================================================

def perform_ocr(image):
    reader = load_easyocr()

    image_np = np.array(image)

    result = reader.readtext(
        image_np,
        detail=1,
        paragraph=False
    )

    return result


# ============================================================
# TÍNH FONT SIZE
# ============================================================

def estimate_font_size(box):
    """
    Ước lượng font theo chiều cao vùng OCR.
    """

    ys = [point[1] for point in box]

    height = max(ys) - min(ys)

    size = int(height * 0.75)

    size = max(size, 12)
    size = min(size, 80)

    return size


# ============================================================
# TÌM FONT HỖ TRỢ TIẾNG TRUNG / VIỆT
# ============================================================

def find_font():
    possible_fonts = [
        # Windows
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\msyhbd.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
        r"C:\Windows\Fonts\arial.ttf",

        # Linux
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.otf",

        # macOS
        "/System/Library/Fonts/PingFang.ttc"
    ]

    for font_path in possible_fonts:
        if os.path.exists(font_path):
            return font_path

    return None


# ============================================================
# TEXT BOX
# ============================================================

def get_box_coordinates(box):
    xs = [int(point[0]) for point in box]
    ys = [int(point[1]) for point in box]

    return (
        min(xs),
        min(ys),
        max(xs),
        max(ys)
    )


# ============================================================
# INPAINT VÙNG CHỮ
# ============================================================

def remove_text_background(image_np, rectangle):
    """
    Xóa chữ cũ bằng OpenCV inpainting.
    """

    x1, y1, x2, y2 = rectangle

    h, w = image_np.shape[:2]

    x1 = max(0, x1)
    y1 = max(0, y1)
    x2 = min(w - 1, x2)
    y2 = min(h - 1, y2)

    mask = np.zeros(
        (h, w),
        dtype=np.uint8
    )

    # Nới vùng mask một chút
    padding = 3

    mx1 = max(0, x1 - padding)
    my1 = max(0, y1 - padding)
    mx2 = min(w - 1, x2 + padding)
    my2 = min(h - 1, y2 + padding)

    cv2.rectangle(
        mask,
        (mx1, my1),
        (mx2, my2),
        255,
        -1
    )

    result = cv2.inpaint(
        image_np,
        mask,
        3,
        cv2.INPAINT_TELEA
    )

    return result


# ============================================================
# VẼ TEXT DỊCH
# ============================================================

def draw_text_inside_box(
    image,
    text,
    rectangle,
    font_path
):
    """
    Vẽ text dịch vào đúng vùng OCR.
    """

    draw = ImageDraw.Draw(image)

    x1, y1, x2, y2 = rectangle

    box_width = max(
        20,
        x2 - x1
    )

    box_height = max(
        20,
        y2 - y1
    )

    font_size = max(
        12,
        min(
            estimate_font_size(
                [
                    (x1, y1),
                    (x2, y1),
                    (x2, y2),
                    (x1, y2)
                ]
            ),
            60
        )
    )

    # --------------------------------------------------------
    # Load font
    # --------------------------------------------------------

    try:

        if font_path:
            font = ImageFont.truetype(
                font_path,
                font_size
            )

        else:
            font = ImageFont.load_default()

    except Exception:

        font = ImageFont.load_default()

    # --------------------------------------------------------
    # Text wrapping
    # --------------------------------------------------------

    words = text.split()

    lines = []

    current = ""

    for word in words:

        test = (
            word
            if not current
            else current + " " + word
        )

        bbox = draw.textbbox(
            (0, 0),
            test,
            font=font
        )

        width = bbox[2] - bbox[0]

        if width <= box_width:

            current = test

        else:

            if current:
                lines.append(current)

            current = word

    if current:
        lines.append(current)

    if not lines:
        lines = [text]

    # --------------------------------------------------------
    # Giảm font nếu quá cao
    # --------------------------------------------------------

    while len(lines) * (font_size + 4) > box_height and font_size > 8:

        font_size -= 1

        try:

            if font_path:
                font = ImageFont.truetype(
                    font_path,
                    font_size
                )

            else:
                font = ImageFont.load_default()

        except Exception:

            break

    # --------------------------------------------------------
    # Vẽ
    # --------------------------------------------------------

    total_height = len(lines) * (
        font_size + 4
    )

    start_y = y1 + max(
        0,
        (box_height - total_height) // 2
    )

    for line in lines:

        bbox = draw.textbbox(
            (0, 0),
            line,
            font=font
        )

        text_width = bbox[2] - bbox[0]

        x = x1 + max(
            0,
            (box_width - text_width) // 2
        )

        draw.text(
            (x, start_y),
            line,
            font=font,
            fill=(0, 0, 0)
        )

        start_y += font_size + 4

    return image


# ============================================================
# DỊCH ẢNH
# ============================================================

def translate_image(uploaded_file):
    """
    OCR -> dịch -> thay chữ.
    """

    original_bytes = uploaded_file.getvalue()

    image = Image.open(
        io.BytesIO(original_bytes)
    ).convert("RGB")

    original_size = image.size

    st.write(
        f"📐 Kích thước ảnh gốc: "
        f"**{original_size[0]} × {original_size[1]} px**"
    )

    # --------------------------------------------------------
    # OCR
    # --------------------------------------------------------

    with st.spinner("🔍 Đang nhận diện chữ trong ảnh..."):

        ocr_results = perform_ocr(image)

    st.write(
        f"🔎 Phát hiện **{len(ocr_results)}** vùng chữ."
    )

    if not ocr_results:

        return (
            io.BytesIO(original_bytes),
            0
        )

    # --------------------------------------------------------
    # Chuyển sang OpenCV
    # --------------------------------------------------------

    image_np = np.array(image)

    # --------------------------------------------------------
    # Xử lý từng vùng
    # --------------------------------------------------------

    progress = st.progress(0)

    font_path = find_font()

    processed = 0

    for index, item in enumerate(ocr_results):

        try:

            box = item[0]
            original_text = str(item[1]).strip()
            confidence = float(item[2])

        except Exception:
            continue

        # Bỏ OCR quá thấp
        if confidence < 0.25:
            continue

        if not should_translate(original_text):
            continue

        # ----------------------------------------------------
        # Dịch
        # ----------------------------------------------------

        translated = translate_text(
            original_text,
            target=(
                "vi"
                if target_language == "Tiếng Việt"
                else "en"
            )
        )

        # ----------------------------------------------------
        # Nội dung cuối
        # ----------------------------------------------------

        if translation_mode == "Song ngữ - Trung + Việt":

            display_text = translated

        else:

            display_text = translated

        rectangle = get_box_coordinates(box)

        # ----------------------------------------------------
        # Xóa chữ gốc
        # ----------------------------------------------------

        image_np = remove_text_background(
            image_np,
            rectangle
        )

        # ----------------------------------------------------
        # Convert PIL
        # ----------------------------------------------------

        image = Image.fromarray(
            cv2.cvtColor(
                image_np,
                cv2.COLOR_BGR2RGB
            )
        )

        # ----------------------------------------------------
        # Vẽ chữ mới
        # ----------------------------------------------------

        image = draw_text_inside_box(
            image,
            display_text,
            rectangle,
            font_path
        )

        processed += 1

        progress.progress(
            min(
                (index + 1) / len(ocr_results),
                1.0
            )
        )

    # --------------------------------------------------------
    # Xuất ảnh
    # --------------------------------------------------------

    output = io.BytesIO()

    output_format = "PNG"

    image.save(
        output,
        format=output_format
    )

    output.seek(0)

    return output, processed


# ============================================================
# ============================================================
#                       UPLOAD
# ============================================================
# ============================================================

st.divider()

st.subheader("📤 Tải file cần dịch")

uploaded_file = st.file_uploader(
    "Chọn file Excel hoặc hình ảnh",
    type=[
        "xlsx",
        "xlsm",
        "xltx",
        "xltm",
        "xls",
        "png",
        "jpg",
        "jpeg",
        "bmp",
        "tif",
        "tiff",
        "webp"
    ]
)


# ============================================================
# XỬ LÝ
# ============================================================

if uploaded_file:

    file_type = detect_file_type(
        uploaded_file.name
    )

    st.success(
        f"📄 File: **{uploaded_file.name}**"
    )

    if file_type == "excel":

        st.info(
            "📊 Đã nhận diện đây là file Excel."
        )

        if st.button(
            "🚀 DỊCH FILE EXCEL",
            use_container_width=True
        ):

            try:

                with st.spinner(
                    "⏳ Đang dịch Excel..."
                ):

                    output, count = translate_excel(
                        uploaded_file
                    )

                st.success(
                    f"✅ Hoàn thành! Đã xử lý "
                    f"**{count} ô chứa văn bản**."
                )

                original_ext = Path(
                    uploaded_file.name
                ).suffix.lower()

                if original_ext in {
                    ".xlsm",
                    ".xltm"
                }:

                    output_name = (
                        Path(uploaded_file.name).stem
                        + "_Trung_Viet"
                        + original_ext
                    )

                    mime = (
                        "application/vnd.ms-excel."
                        "sheet.macroEnabled.12"
                    )

                else:

                    output_name = (
                        Path(uploaded_file.name).stem
                        + "_Trung_Viet.xlsx"
                    )

                    mime = (
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    )

                st.download_button(
                    label="⬇️ TẢI FILE EXCEL ĐÃ DỊCH",
                    data=output.getvalue(),
                    file_name=output_name,
                    mime=mime,
                    use_container_width=True
                )

            except Exception as e:

                st.error(
                    "❌ Không thể xử lý file Excel."
                )

                st.exception(e)

    elif file_type == "image":

        st.info(
            "🖼️ Đã nhận diện đây là file hình ảnh."
        )

        # ----------------------------------------------------
        # Preview
        # ----------------------------------------------------

        col1, col2 = st.columns(2)

        with col1:

            st.subheader("Ảnh gốc")

            st.image(
                uploaded_file,
                use_container_width=True
            )

        # ----------------------------------------------------
        # Button
        # ----------------------------------------------------

        if st.button(
            "🚀 DỊCH HÌNH ẢNH",
            use_container_width=True
        ):

            try:

                output, count = translate_image(
                    uploaded_file
                )

                st.success(
                    f"✅ Hoàn thành! Đã xử lý "
                    f"**{count} vùng chữ**."
                )

                with col2:

                    st.subheader(
                        "Ảnh sau khi dịch"
                    )

                    st.image(
                        output.getvalue(),
                        use_container_width=True
                    )

                output_name = (
                    Path(
                        uploaded_file.name
                    ).stem
                    + "_Trung_Viet.png"
                )

                st.download_button(
                    label="⬇️ TẢI ẢNH ĐÃ DỊCH",
                    data=output.getvalue(),
                    file_name=output_name,
                    mime="image/png",
                    use_container_width=True
                )

            except Exception as e:

                st.error(
                    "❌ Không thể xử lý hình ảnh."
                )

                st.exception(e)

    else:

        st.error(
            "❌ Định dạng file chưa được hỗ trợ."
        )


# ============================================================
# THÔNG TIN
# ============================================================

st.divider()

st.caption(
    "Ứng dụng xử lý file trực tiếp trong bộ nhớ RAM. "
    "Không cần lưu file đầu ra vào /mnt/data."
)
