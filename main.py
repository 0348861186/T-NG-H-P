import streamlit as st
import os
import re
import io
import tempfile
import hashlib
from pathlib import Path
from copy import copy

import numpy as np
import pandas as pd

from PIL import Image

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins
from openpyxl.utils import get_column_letter

from deep_translator import GoogleTranslator


# ============================================================
# CẤU HÌNH
# ============================================================

st.set_page_config(
    page_title="Dịch Excel / Ảnh Trung - Việt",
    page_icon="🌏",
    layout="wide"
)

APP_TITLE = "🌏 DỊCH EXCEL / ẢNH TRUNG ↔ VIỆT"

SUPPORTED_EXCEL = [
    ".xlsx",
    ".xlsm",
    ".xls",
    ".xlsb",
    ".ods",
    ".csv"
]

SUPPORTED_IMAGE = [
    ".png",
    ".jpg",
    ".jpeg",
    ".webp",
    ".bmp",
    ".tif",
    ".tiff"
]


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
    <style>
    .main-title {
        font-size: 30px;
        font-weight: 700;
        margin-bottom: 5px;
    }

    .sub-title {
        color: #666;
        margin-bottom: 20px;
    }

    .info-box {
        padding: 12px 15px;
        border-radius: 8px;
        background: #f4f6f8;
        border: 1px solid #ddd;
        margin-bottom: 15px;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# SESSION STATE
# ============================================================

if "translation_cache" not in st.session_state:
    st.session_state.translation_cache = {}

if "ocr_reader" not in st.session_state:
    st.session_state.ocr_reader = None


# ============================================================
# HÀM NHẬN DIỆN NGÔN NGỮ
# ============================================================

CHINESE_RE = re.compile(
    r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]"
)

VIETNAMESE_RE = re.compile(
    r"[ÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚĂĐĨŨƠàáâãèéêìíòóôõùúăđĩũơƯĂÂÊÔƠưăâêôơ]"
)


def count_characters(text):
    """
    Đếm số lượng ký tự Trung / Việt / Latin.
    """

    if not isinstance(text, str):
        text = str(text)

    chinese = len(CHINESE_RE.findall(text))
    vietnamese = len(VIETNAMESE_RE.findall(text))

    latin = len(
        re.findall(
            r"[A-Za-z]",
            text
        )
    )

    return chinese, vietnamese, latin


def detect_language(text):
    """
    Nhận diện tương đối:
    - zh: Trung
    - vi: Việt
    - other: khác
    """

    if text is None:
        return "other"

    text = str(text).strip()

    if not text:
        return "other"

    chinese, vietnamese, latin = count_characters(text)

    if chinese > 0 and chinese >= vietnamese:
        return "zh"

    if vietnamese > 0:
        return "vi"

    return "other"


def detect_document_language(texts):
    """
    Nhận diện ngôn ngữ chính của toàn bộ file.
    """

    total_chinese = 0
    total_vietnamese = 0
    total_latin = 0

    for text in texts:

        if text is None:
            continue

        text = str(text).strip()

        if not text:
            continue

        chinese, vietnamese, latin = count_characters(text)

        total_chinese += chinese
        total_vietnamese += vietnamese
        total_latin += latin

    if total_chinese > 0 and total_chinese >= total_vietnamese:
        return "zh"

    if total_vietnamese > 0:
        return "vi"

    return "other"


# ============================================================
# XÁC ĐỊNH CHẾ ĐỘ DỊCH
# ============================================================

def get_direction_from_selection(selection, detected_language):

    if selection == "Trung → Việt":
        return "zh_to_vi"

    if selection == "Việt → Trung":
        return "vi_to_zh"

    # Tự động
    if detected_language == "zh":
        return "zh_to_vi"

    if detected_language == "vi":
        return "vi_to_zh"

    return "zh_to_vi"


# ============================================================
# KIỂM TRA NỘI DUNG CÓ CẦN DỊCH KHÔNG
# ============================================================

def is_formula(value):
    return (
        isinstance(value, str)
        and value.startswith("=")
    )


def is_number(value):

    if value is None:
        return True

    if isinstance(
        value,
        (
            int,
            float,
            complex
        )
    ):
        return True

    return False


def is_date_like(text):

    if not isinstance(text, str):
        return False

    patterns = [
        r"^\d{1,4}[/-]\d{1,2}[/-]\d{1,4}$",
        r"^\d{1,2}:\d{2}(:\d{2})?$",
        r"^\d{1,2}[/-]\d{1,2}$",
        r"^\d{4}$"
    ]

    for pattern in patterns:

        if re.match(pattern, text.strip()):
            return True

    return False


def is_code_like(text):

    if not isinstance(text, str):
        return False

    value = text.strip()

    if not value:
        return True

    # Mã sản phẩm / mã nhân viên / số hiệu
    if re.fullmatch(
        r"[A-Z0-9_\-./]+",
        value
    ):
        return True

    return False


def should_translate(value):

    if value is None:
        return False

    if is_number(value):
        return False

    if not isinstance(value, str):
        return False

    text = value.strip()

    if not text:
        return False

    if is_formula(text):
        return False

    if is_date_like(text):
        return False

    if is_code_like(text):
        return False

    zh_count, vi_count, latin_count = count_characters(text)

    # Không có chữ Trung / Việt
    if zh_count == 0 and vi_count == 0:
        return False

    return True


# ============================================================
# CACHE DỊCH
# ============================================================

def cache_key(text, direction):

    raw = f"{direction}|{text}"

    return hashlib.sha256(
        raw.encode("utf-8")
    ).hexdigest()


# ============================================================
# DỊCH
# ============================================================

def translate_text(text, direction):

    if not should_translate(text):
        return text

    text = str(text).strip()

    # Nếu đang dịch Trung -> Việt
    if direction == "zh_to_vi":

        # Không dịch nếu nội dung đã là tiếng Việt
        if detect_language(text) == "vi":
            return text

        source = "zh-CN"
        target = "vi"

    else:

        # Việt -> Trung
        if detect_language(text) == "zh":
            return text

        source = "vi"
        target = "zh-CN"

    key = cache_key(
        text,
        direction
    )

    if key in st.session_state.translation_cache:
        return st.session_state.translation_cache[key]

    try:

        translator = GoogleTranslator(
            source=source,
            target=target
        )

        result = translator.translate(text)

        if result is None:
            result = text

        result = str(result).strip()

        st.session_state.translation_cache[key] = result

        return result

    except Exception as e:

        st.warning(
            f"Không thể dịch nội dung: {text[:80]}"
        )

        return text


# ============================================================
# TẠO TEXT SONG NGỮ
# ============================================================

def make_bilingual_text(
    original,
    translated
):

    original = "" if original is None else str(original)
    translated = "" if translated is None else str(translated)

    if not translated:
        return original

    if original.strip() == translated.strip():
        return original

    return (
        f"{original}\n"
        f"{translated}"
    )


# ============================================================
# COPY STYLE CELL
# ============================================================

def copy_cell_style(source, target):

    if source.has_style:

        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)
        target.number_format = source.number_format
        target.protection = copy(source.protection)


# ============================================================
# COPY WIDTH / HEIGHT
# ============================================================

def copy_sheet_dimensions(
    source_ws,
    target_ws
):

    for key, dimension in source_ws.column_dimensions.items():

        target_ws.column_dimensions[key].width = dimension.width
        target_ws.column_dimensions[key].hidden = dimension.hidden

    for key, dimension in source_ws.row_dimensions.items():

        target_ws.row_dimensions[key].height = dimension.height
        target_ws.row_dimensions[key].hidden = dimension.hidden


# ============================================================
# XỬ LÝ MERGED CELLS
# ============================================================

def get_merged_range(
    ws,
    row,
    column
):

    for merged in ws.merged_cells.ranges:

        if (
            merged.min_row <= row <= merged.max_row
            and
            merged.min_col <= column <= merged.max_col
        ):
            return merged

    return None


# ============================================================
# XỬ LÝ EXCEL OPENPYXL
# ============================================================

def process_openpyxl_workbook(
    uploaded_file,
    direction
):

    uploaded_file.seek(0)

    keep_vba = uploaded_file.name.lower().endswith(
        ".xlsm"
    )

    wb = load_workbook(
        uploaded_file,
        data_only=False,
        keep_vba=keep_vba
    )

    progress = st.progress(0)

    sheets = wb.worksheets

    total_cells = 0

    for ws in sheets:

        for row in ws.iter_rows():

            for cell in row:

                if cell.value is not None:
                    total_cells += 1

    processed = 0

    # --------------------------------------------------------
    # DỊCH TỪNG CELL
    # --------------------------------------------------------

    for ws in sheets:

        # Không hiện grid
        ws.sheet_view.showGridLines = False

        for row in ws.iter_rows():

            for cell in row:

                value = cell.value

                if should_translate(value):

                    original = str(value)

                    translated = translate_text(
                        original,
                        direction
                    )

                    # Giữ nguyên công thức / số
                    cell.value = make_bilingual_text(
                        original,
                        translated
                    )

                    # Giữ format gốc
                    old_alignment = copy(
                        cell.alignment
                    )

                    cell.alignment = Alignment(
                        horizontal=old_alignment.horizontal or "center",
                        vertical=old_alignment.vertical or "center",
                        wrap_text=True,
                        text_rotation=old_alignment.text_rotation,
                        shrink_to_fit=old_alignment.shrink_to_fit,
                        indent=old_alignment.indent
                    )

                processed += 1

                if total_cells > 0:
                    progress.progress(
                        min(
                            processed / total_cells,
                            1.0
                        )
                    )

    progress.empty()

    # --------------------------------------------------------
    # THIẾT LẬP IN
    # --------------------------------------------------------

    for ws in wb.worksheets:

        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0

        ws.sheet_properties.pageSetUpPr.fitToPage = True

        ws.page_margins = PageMargins(
            left=0.2,
            right=0.2,
            top=0.3,
            bottom=0.3,
            header=0.1,
            footer=0.1
        )

    return wb


# ============================================================
# EXCEL CŨ: XLS / XLSB / ODS / CSV
# ============================================================

def read_legacy_excel(
    uploaded_file,
    extension
):

    uploaded_file.seek(0)

    if extension == ".xls":

        return pd.read_excel(
            uploaded_file,
            sheet_name=None,
            engine="xlrd",
            header=None
        )

    if extension == ".xlsb":

        return pd.read_excel(
            uploaded_file,
            sheet_name=None,
            engine="pyxlsb",
            header=None
        )

    if extension == ".ods":

        return pd.read_excel(
            uploaded_file,
            sheet_name=None,
            engine="odf",
            header=None
        )

    if extension == ".csv":

        df = pd.read_csv(
            uploaded_file,
            header=None,
            dtype=str
        )

        return {
            "Sheet1": df
        }

    raise ValueError(
        f"Không hỗ trợ định dạng {extension}"
    )


def dataframe_to_workbook(
    sheets,
    direction
):

    wb = Workbook()

    # Xóa sheet mặc định
    default_ws = wb.active
    wb.remove(default_ws)

    for sheet_name, df in sheets.items():

        # Excel giới hạn tên sheet 31 ký tự
        safe_name = str(sheet_name)[:31]

        if not safe_name:
            safe_name = "Sheet"

        ws = wb.create_sheet(
            title=safe_name
        )

        for r_idx, row in enumerate(
            df.itertuples(
                index=False,
                name=None
            ),
            1
        ):

            for c_idx, value in enumerate(
                row,
                1
            ):

                if pd.isna(value):
                    value = ""

                cell = ws.cell(
                    row=r_idx,
                    column=c_idx
                )

                original = str(value)

                if should_translate(value):

                    translated = translate_text(
                        original,
                        direction
                    )

                    cell.value = make_bilingual_text(
                        original,
                        translated
                    )

                else:

                    cell.value = value

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True
                )

                cell.font = Font(
                    name="Microsoft YaHei",
                    size=10
                )

        # Auto width
        for col in range(
            1,
            ws.max_column + 1
        ):

            max_length = 0

            for row in range(
                1,
                ws.max_row + 1
            ):

                value = ws.cell(
                    row=row,
                    column=col
                ).value

                if value is not None:

                    max_length = max(
                        max_length,
                        len(
                            str(value).split("\n")[0]
                        )
                    )

            ws.column_dimensions[
                get_column_letter(col)
            ].width = min(
                max(max_length + 3, 10),
                40
            )

        ws.sheet_view.showGridLines = False

    return wb


# ============================================================
# OCR ENGINE
# ============================================================

@st.cache_resource
def load_ocr_reader():

    import easyocr

    reader = easyocr.Reader(
        [
            "ch_sim",
            "en"
        ],
        gpu=False,
        verbose=False
    )

    return reader


# ============================================================
# OCR ẢNH
# ============================================================

def preprocess_image(image):

    img = np.array(image)

    if len(img.shape) == 3:

        # RGB
        gray = (
            0.299 * img[:, :, 0]
            +
            0.587 * img[:, :, 1]
            +
            0.114 * img[:, :, 2]
        )

    else:

        gray = img

    gray = gray.astype(
        np.uint8
    )

    # Phóng to để OCR chính xác hơn
    scale = 2

    resized = np.array(
        Image.fromarray(gray).resize(
            (
                gray.shape[1] * scale,
                gray.shape[0] * scale
            ),
            Image.Resampling.LANCZOS
        )
    )

    return resized


def perform_ocr(
    image
):

    reader = load_ocr_reader()

    processed = preprocess_image(
        image
    )

    results = reader.readtext(
        processed,
        detail=1,
        paragraph=False,
        width_ths=0.7,
        link_threshold=0.3,
        low_text=0.3,
        text_threshold=0.6
    )

    return results


# ============================================================
# OCR -> EXCEL
# ============================================================

def ocr_results_to_workbook(
    results,
    direction
):

    wb = Workbook()

    ws = wb.active
    ws.title = "Dữ liệu OCR"

    # --------------------------------------------------------
    # HEADER
    # --------------------------------------------------------

    headers = [
        ("STT", "STT"),
        ("原文", "Nguyên bản"),
        ("翻译", "Bản dịch"),
        ("位置", "Vị trí"),
        ("置信度", "Độ tin cậy")
    ]

    for col, (cn, vi) in enumerate(
        headers,
        1
    ):

        cell = ws.cell(
            row=1,
            column=col
        )

        if cn == vi:
            cell.value = cn
        else:
            cell.value = (
                f"{cn}\n"
                f"{vi}"
            )

        cell.font = Font(
            name="Microsoft YaHei",
            size=11,
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

        cell.fill = PatternFill(
            "solid",
            fgColor="ED7D00"
        )

    # --------------------------------------------------------
    # OCR RESULTS
    # --------------------------------------------------------

    row_number = 2

    for result in results:

        if len(result) < 3:
            continue

        box = result[0]
        text = result[1]
        confidence = result[2]

        text = str(text).strip()

        if not text:
            continue

        translated = translate_text(
            text,
            direction
        )

        # Vị trí
        x1 = min(
            int(point[0])
            for point in box
        )

        y1 = min(
            int(point[1])
            for point in box
        )

        x2 = max(
            int(point[0])
            for point in box
        )

        y2 = max(
            int(point[1])
            for point in box
        )

        position = (
            f"X={x1}, Y={y1}, "
            f"W={x2-x1}, H={y2-y1}"
        )

        values = [
            row_number - 1,
            text,
            translated,
            position,
            round(
                float(confidence) * 100,
                2
            )
        ]

        for col, value in enumerate(
            values,
            1
        ):

            cell = ws.cell(
                row=row_number,
                column=col,
                value=value
            )

            cell.font = Font(
                name="Microsoft YaHei",
                size=10
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

        row_number += 1

    # --------------------------------------------------------
    # BORDER
    # --------------------------------------------------------

    thin = Side(
        style="thin",
        color="000000"
    )

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    for row in ws.iter_rows(
        min_row=1,
        max_row=ws.max_row,
        min_col=1,
        max_col=5
    ):

        for cell in row:
            cell.border = border

    # --------------------------------------------------------
    # COLUMN WIDTH
    # --------------------------------------------------------

    widths = {
        "A": 8,
        "B": 35,
        "C": 35,
        "D": 25,
        "E": 15
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.row_dimensions[1].height = 38

    ws.freeze_panes = "A2"

    ws.sheet_view.showGridLines = False

    return wb


# ============================================================
# LƯU WORKBOOK AN TOÀN
# ============================================================

def safe_save_workbook(
    wb,
    original_name
):

    extension = Path(
        original_name
    ).suffix.lower()

    if extension not in [
        ".xlsx",
        ".xlsm"
    ]:
        extension = ".xlsx"

    temp_dir = tempfile.mkdtemp(
        prefix="translator_"
    )

    safe_name = (
        Path(original_name).stem
        +
        "_song_ngu"
        +
        extension
    )

    output_path = os.path.join(
        temp_dir,
        safe_name
    )

    # Đảm bảo thư mục tồn tại
    os.makedirs(
        temp_dir,
        exist_ok=True
    )

    wb.save(
        output_path
    )

    # Kiểm tra file sau khi save
    if not os.path.exists(
        output_path
    ):
        raise FileNotFoundError(
            f"Không tạo được file: {output_path}"
        )

    if os.path.getsize(
        output_path
    ) == 0:
        raise IOError(
            "File Excel được tạo nhưng có kích thước 0 byte."
        )

    return output_path


# ============================================================
# THU THẬP TEXT TỪ EXCEL ĐỂ NHẬN DIỆN NGÔN NGỮ
# ============================================================

def collect_excel_texts_openpyxl(
    uploaded_file
):

    uploaded_file.seek(0)

    keep_vba = uploaded_file.name.lower().endswith(
        ".xlsm"
    )

    wb = load_workbook(
        uploaded_file,
        read_only=True,
        data_only=False,
        keep_vba=keep_vba
    )

    texts = []

    for ws in wb.worksheets:

        for row in ws.iter_rows():

            for cell in row:

                value = cell.value

                if isinstance(
                    value,
                    str
                ):

                    texts.append(
                        value
                    )

    return texts


def collect_dataframe_texts(
    sheets
):

    texts = []

    for df in sheets.values():

        for column in df.columns:

            for value in df[column]:

                if not pd.isna(value):

                    texts.append(
                        str(value)
                    )

    return texts


# ============================================================
# GIAO DIỆN
# ============================================================

st.markdown(
    f'<div class="main-title">{APP_TITLE}</div>',
    unsafe_allow_html=True
)

st.markdown(
    """
    <div class="sub-title">
    Giữ logic Excel song ngữ của phiên bản gốc,
    đồng thời mở rộng xử lý Excel và ảnh.
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header("⚙️ Tùy chọn")

    translation_mode = st.radio(
        "Chế độ dịch",
        [
            "Tự động",
            "Trung → Việt",
            "Việt → Trung"
        ],
        index=0
    )

    st.markdown("---")

    st.info(
        """
        **Tự động nhận diện:**

        🇨🇳 File chủ yếu tiếng Trung
        → Trung → Việt

        🇻🇳 File chủ yếu tiếng Việt
        → Việt → Trung
        """
    )

    st.markdown("---")

    st.caption(
        "OCR ảnh sử dụng nhận diện tiếng Trung + Latin. "
        "Ảnh có độ phân giải cao và chữ rõ sẽ cho kết quả tốt hơn."
    )


# ============================================================
# UPLOAD
# ============================================================

uploaded_file = st.file_uploader(
    "📂 Chọn file Excel hoặc file ảnh",
    type=[
        x.replace(".", "")
        for x in (
            SUPPORTED_EXCEL
            +
            SUPPORTED_IMAGE
        )
    ]
)


# ============================================================
# XỬ LÝ FILE
# ============================================================

if uploaded_file is not None:

    file_name = uploaded_file.name

    extension = Path(
        file_name
    ).suffix.lower()

    st.success(
        f"Đã tải: **{file_name}**"
    )

    # --------------------------------------------------------
    # XỬ LÝ EXCEL
    # --------------------------------------------------------

    if extension in SUPPORTED_EXCEL:

        st.subheader(
            "📊 Nhận diện file Excel"
        )

        st.write(
            f"Định dạng: `{extension}`"
        )

        detected_language = "other"

        try:

            if extension in [
                ".xlsx",
                ".xlsm"
            ]:

                texts = collect_excel_texts_openpyxl(
                    uploaded_file
                )

            else:

                sheets = read_legacy_excel(
                    uploaded_file,
                    extension
                )

                texts = collect_dataframe_texts(
                    sheets
                )

            detected_language = detect_document_language(
                texts
            )

        except Exception as e:

            st.error(
                f"Không thể phân tích file: {e}"
            )

            st.stop()

        # ----------------------------------------------------
        # HIỂN THỊ NHẬN DIỆN
        # ----------------------------------------------------

        if detected_language == "zh":

            language_name = (
                "🇨🇳 Tiếng Trung"
            )

        elif detected_language == "vi":

            language_name = (
                "🇻🇳 Tiếng Việt"
            )

        else:

            language_name = (
                "🌐 Không xác định rõ"
            )

        direction = get_direction_from_selection(
            translation_mode,
            detected_language
        )

        if direction == "zh_to_vi":

            direction_text = (
                "🇨🇳 Trung → 🇻🇳 Việt"
            )

        else:

            direction_text = (
                "🇻🇳 Việt → 🇨🇳 Trung"
            )

        col1, col2 = st.columns(2)

        with col1:

            st.info(
                f"Ngôn ngữ nhận diện: **{language_name}**"
            )

        with col2:

            st.info(
                f"Chế độ dịch: **{direction_text}**"
            )

        # ----------------------------------------------------
        # NÚT DỊCH
        # ----------------------------------------------------

        if st.button(
            "🚀 BẮT ĐẦU DỊCH",
            type="primary",
            use_container_width=True
        ):

            try:

                with st.spinner(
                    "Đang xử lý và dịch file..."
                ):

                    # ----------------------------------------
                    # XLSX / XLSM
                    # ----------------------------------------

                    if extension in [
                        ".xlsx",
                        ".xlsm"
                    ]:

                        wb = process_openpyxl_workbook(
                            uploaded_file,
                            direction
                        )

                    # ----------------------------------------
                    # XLS / XLSB / ODS / CSV
                    # ----------------------------------------

                    else:

                        uploaded_file.seek(0)

                        sheets = read_legacy_excel(
                            uploaded_file,
                            extension
                        )

                        wb = dataframe_to_workbook(
                            sheets,
                            direction
                        )

                    output_path = safe_save_workbook(
                        wb,
                        file_name
                    )

                st.success(
                    "✅ Đã dịch xong!"
                )

                st.download_button(
                    label="⬇️ TẢI FILE EXCEL SONG NGỮ",
                    data=open(
                        output_path,
                        "rb"
                    ).read(),
                    file_name=os.path.basename(
                        output_path
                    ),
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    use_container_width=True
                )

                st.caption(
                    f"File đã tạo thành công: {output_path}"
                )

            except Exception as e:

                st.error(
                    "❌ Có lỗi khi xử lý file."
                )

                st.exception(e)

    # --------------------------------------------------------
    # XỬ LÝ ẢNH
    # --------------------------------------------------------

    elif extension in SUPPORTED_IMAGE:

        st.subheader(
            "🖼️ Nhận diện file ảnh"
        )

        image = Image.open(
            uploaded_file
        )

        st.image(
            image,
            caption=file_name,
            use_container_width=True
        )

        # ----------------------------------------------------
        # ẢNH LUÔN ĐƯỢC OCR
        # ----------------------------------------------------

        direction = get_direction_from_selection(
            translation_mode,
            "zh"
        )

        if translation_mode == "Tự động":

            st.info(
                "Ảnh sẽ được OCR trước, sau đó hệ thống "
                "phân tích ngôn ngữ nhận diện được."
            )

        if st.button(
            "🔍 OCR + DỊCH ẢNH",
            type="primary",
            use_container_width=True
        ):

            try:

                with st.spinner(
                    "Đang nhận diện chữ trong ảnh..."
                ):

                    results = perform_ocr(
                        image
                    )

                if not results:

                    st.warning(
                        "Không nhận diện được chữ trong ảnh."
                    )

                    st.stop()

                # ------------------------------------------------
                # NHẬN DIỆN NGÔN NGỮ SAU OCR
                # ------------------------------------------------

                ocr_texts = []

                for result in results:

                    if len(result) >= 2:

                        ocr_texts.append(
                            str(result[1])
                        )

                detected_language = detect_document_language(
                    ocr_texts
                )

                direction = get_direction_from_selection(
                    translation_mode,
                    detected_language
                )

                if detected_language == "zh":

                    lang_text = "🇨🇳 Tiếng Trung"

                elif detected_language == "vi":

                    lang_text = "🇻🇳 Tiếng Việt"

                else:

                    lang_text = "🌐 Không xác định"

                if direction == "zh_to_vi":

                    direction_text = (
                        "🇨🇳 Trung → 🇻🇳 Việt"
                    )

                else:

                    direction_text = (
                        "🇻🇳 Việt → 🇨🇳 Trung"
                    )

                st.success(
                    f"Nhận diện: **{lang_text}** | "
                    f"Dịch: **{direction_text}**"
                )

                # ------------------------------------------------
                # HIỂN THỊ OCR
                # ------------------------------------------------

                st.subheader(
                    "📝 Nội dung OCR"
                )

                preview_data = []

                for result in results:

                    box = result[0]
                    text = result[1]
                    confidence = result[2]

                    translated = translate_text(
                        text,
                        direction
                    )

                    preview_data.append(
                        {
                            "Nguyên bản": text,
                            "Bản dịch": translated,
                            "Độ tin cậy": round(
                                float(confidence) * 100,
                                2
                            )
                        }
                    )

                st.dataframe(
                    pd.DataFrame(
                        preview_data
                    ),
                    use_container_width=True
                )

                # ------------------------------------------------
                # TẠO EXCEL
                # ------------------------------------------------

                with st.spinner(
                    "Đang tạo Excel song ngữ..."
                ):

                    wb = ocr_results_to_workbook(
                        results,
                        direction
                    )

                    output_path = safe_save_workbook(
                        wb,
                        Path(file_name).stem + ".xlsx"
                    )

                st.success(
                    "✅ Đã OCR và tạo Excel song ngữ!"
                )

                st.download_button(
                    label="⬇️ TẢI EXCEL SONG NGỮ",
                    data=open(
                        output_path,
                        "rb"
                    ).read(),
                    file_name=os.path.basename(
                        output_path
                    ),
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    use_container_width=True
                )

            except Exception as e:

                st.error(
                    "❌ OCR / dịch ảnh thất bại."
                )

                st.exception(e)

    else:

        st.error(
            f"Định dạng {extension} chưa được hỗ trợ."
        )


# ============================================================
# FOOTER
# ============================================================

st.markdown("---")

st.caption(
    "Hệ thống giữ nguyên logic dịch song ngữ của code gốc, "
    "đồng thời tự nhận diện định dạng file và ngôn ngữ."
)
