import io
import os
import re
import time
from copy import copy
from pathlib import Path
from typing import List

import streamlit as st
from PIL import Image, ImageOps
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field
from google import genai
from google.genai import types


# ============================================================
# TRUNG <-> VIỆT EXCEL / IMAGE TRANSLATOR - V2
#
# Excel:
#   - Sửa trực tiếp bản sao workbook
#   - Giữ merge, style, border, màu, font, width, height,
#     freeze panes và các thuộc tính sheet chính
#   - XLSM: keep_vba=True
#
# Image:
#   - Gemini Vision đọc bảng
#   - Nhận dạng cells + rowspan + colspan
#   - Dựng lại merge cells
#   - OCR -> dịch -> Excel
#   - Song ngữ: nguyên văn dòng trên + bản dịch dòng dưới
#
# Lưu ý:
#   OpenPyXL không đảm bảo bảo toàn 100% mọi đối tượng Excel
#   đặc biệt như một số drawing/chart/slicer/embedded object.
# ============================================================

APP_TITLE = "🇨🇳 ↔ 🇻🇳 Trung ↔ Việt Excel & Image Translator"

MODEL_DEFAULT = "gemini-3.7-flash"
BATCH_SIZE = 30
MAX_RETRIES = 3

HEADER_ORANGE = "F28C00"
WHITE = "FFFFFF"
BLACK = "000000"
THIN_GRAY = "808080"


# ============================================================
# PYDANTIC SCHEMAS
# ============================================================

class TranslationItem(BaseModel):
    id: int
    translation: str = ""


class TranslationBatch(BaseModel):
    items: List[TranslationItem] = Field(default_factory=list)


class ImageCell(BaseModel):
    row: int = Field(ge=0)
    col: int = Field(ge=0)
    value: str = ""
    rowspan: int = Field(default=1, ge=1)
    colspan: int = Field(default=1, ge=1)


class ImageTable(BaseModel):
    title: str = ""
    cells: List[ImageCell] = Field(default_factory=list)
    row_count: int = Field(default=0, ge=0)
    col_count: int = Field(default=0, ge=0)


# ============================================================
# API / CLIENT
# ============================================================

def get_api_key():
    """Lấy GEMINI_API_KEY từ Streamlit Secrets hoặc environment."""
    try:
        key = st.secrets.get("GEMINI_API_KEY", "")
        if key:
            return key
    except Exception:
        pass

    return os.getenv("GEMINI_API_KEY", "")


@st.cache_resource(show_spinner=False)
def get_client(api_key: str):
    return genai.Client(api_key=api_key)


# ============================================================
# TEXT HELPERS
# ============================================================

def is_formula(value):
    return isinstance(value, str) and value.startswith("=")


def contains_chinese(text):
    return bool(re.search(r"[\u3400-\u4dbf\u4e00-\u9fff]", text or ""))


def contains_vietnamese(text):
    return bool(
        re.search(
            r"[ăâđêôơưĂÂĐÊÔƠƯ"
            r"àáạảãằắặẳẵầấậẩẫ"
            r"èéẹẻẽềếệểễ"
            r"ìíịỉĩ"
            r"òóọỏõồốộổổờớợởỡ"
            r"ùúụủũừứựửữ"
            r"ỳýỵỷỹ]",
            text or "",
        )
    )


def is_translatable_text(value):
    if value is None or not isinstance(value, str):
        return False

    text = value.strip()

    if not text:
        return False

    if is_formula(text):
        return False

    # Chỉ có số / ký hiệu / dấu câu cơ bản.
    if re.fullmatch(r"[\d\s.,:/\\%+\-*=()]+", text):
        return False

    return True


def already_bilingual(text, source_lang, target_lang):
    """
    Kiểm tra tương đối thông minh xem ô đã chứa cả hai ngôn ngữ chưa.
    Không coi mọi ô nhiều dòng là song ngữ.
    """
    if not isinstance(text, str):
        return False

    lines = [x.strip() for x in text.splitlines() if x.strip()]

    if len(lines) < 2:
        return False

    if source_lang == "zh" and target_lang == "vi":
        return (
            any(contains_chinese(x) for x in lines)
            and any(contains_vietnamese(x) for x in lines)
        )

    if source_lang == "vi" and target_lang == "zh":
        return (
            any(contains_vietnamese(x) for x in lines)
            and any(contains_chinese(x) for x in lines)
        )

    return False


def bilingual_text(original, translation):
    if not translation:
        return original

    if not isinstance(original, str):
        return translation

    return f"{original}\n{translation}"


def preserve_alignment_with_wrap(cell):
    """Giữ alignment cũ, chỉ bật wrap_text."""
    old = copy(cell.alignment)

    cell.alignment = Alignment(
        horizontal=old.horizontal,
        vertical=old.vertical or "center",
        textRotation=old.textRotation,
        wrap_text=True,
        shrink_to_fit=old.shrink_to_fit,
        indent=old.indent,
    )


# ============================================================
# GEMINI TRANSLATION
# ============================================================

def translate_batch_once(client, model, texts, source_lang, target_lang):
    if not texts:
        return {}

    language_name = {
        "zh": "tiếng Trung",
        "vi": "tiếng Việt",
    }

    source = language_name[source_lang]
    target = language_name[target_lang]

    numbered = "\n".join(
        f"[{i}] {text}" for i, text in enumerate(texts)
    )

    prompt = f"""
Bạn là biên dịch viên chuyên nghiệp Trung - Việt trong môi trường
nhà máy, sản xuất, máy móc, nhân sự, bảng chấm công, biểu mẫu,
quản lý sản xuất và Excel.

Nhiệm vụ:
Dịch từ {source} sang {target}.

YÊU CẦU BẮT BUỘC:
1. Giữ nguyên ý nghĩa.
2. Ưu tiên thuật ngữ kỹ thuật và thuật ngữ nhà máy tại Việt Nam.
3. Không dịch số, mã máy, mã sản phẩm, ký hiệu và đơn vị nếu không cần.
4. Không thêm giải thích.
5. Không thêm dấu ngoặc kép.
6. Không tự ý bỏ nội dung.
7. Giữ đúng ID.
8. Mỗi ID chỉ có đúng một bản dịch.
9. Nếu nội dung là tên cột/bảng, dịch ngắn gọn, tự nhiên.
10. Giữ nguyên tên riêng, mã số và ký hiệu.
11. Không tạo thêm ID.
12. Không gộp các ID.
13. Nếu một ô có nhiều dòng thì giữ cấu trúc dòng khi có thể.

Danh sách cần dịch:
{numbered}
"""

    response = client.models.generate_content(
        model=model,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=TranslationBatch,
        ),
    )

    data = response.parsed

    if data is None:
        data = TranslationBatch.model_validate_json(response.text)

    result = {}

    for item in data.items:
        if 0 <= item.id < len(texts):
            result[item.id] = item.translation.strip()

    return result


def translate_batch(client, model, texts, source_lang, target_lang):
    """
    Retry batch khi API lỗi hoặc response thiếu ID.
    """
    if not texts:
        return {}

    last_error = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            result = translate_batch_once(
                client,
                model,
                texts,
                source_lang,
                target_lang,
            )

            missing = [
                i for i in range(len(texts))
                if i not in result
            ]

            if missing:
                raise ValueError(
                    "Gemini trả thiếu bản dịch cho ID: "
                    + ", ".join(map(str, missing[:20]))
                )

            return result

        except Exception as exc:
            last_error = exc

            if attempt < MAX_RETRIES:
                time.sleep(2 ** (attempt - 1))

    raise RuntimeError(
        f"Dịch batch thất bại sau {MAX_RETRIES} lần thử: {last_error}"
    )


def translate_texts(
    client,
    model,
    texts,
    source_lang,
    target_lang,
    progress=None,
):
    """Dịch toàn bộ danh sách theo batch."""
    all_results = {}

    total = len(texts)

    if total == 0:
        if progress:
            progress(1.0)
        return all_results

    for start in range(0, total, BATCH_SIZE):
        batch = texts[start:start + BATCH_SIZE]

        result = translate_batch(
            client,
            model,
            batch,
            source_lang,
            target_lang,
        )

        for local_id, translation in result.items():
            all_results[start + local_id] = translation

        if progress:
            done = min(start + BATCH_SIZE, total)
            progress(done / total)

    return all_results


# ============================================================
# EXCEL
# ============================================================

def copy_sheet_properties(src, dst):
    """Sao chép các thuộc tính sheet phổ biến."""

    dst.sheet_view.showGridLines = src.sheet_view.showGridLines
    dst.freeze_panes = src.freeze_panes

    dst.sheet_format.defaultRowHeight = (
        src.sheet_format.defaultRowHeight
    )
    dst.sheet_format.defaultColWidth = (
        src.sheet_format.defaultColWidth
    )

    if src.sheet_properties.pageSetUpPr:
        dst.sheet_properties.pageSetUpPr = copy(
            src.sheet_properties.pageSetUpPr
        )

    dst.page_margins = copy(src.page_margins)
    dst.page_setup = copy(src.page_setup)
    dst.print_options = copy(src.print_options)

    for key, value in src.column_dimensions.items():
        dst.column_dimensions[key].width = value.width
        dst.column_dimensions[key].hidden = value.hidden
        dst.column_dimensions[key].bestFit = value.bestFit

    for key, value in src.row_dimensions.items():
        dst.row_dimensions[key].height = value.height
        dst.row_dimensions[key].hidden = value.hidden


def translate_excel(
    uploaded_bytes,
    filename,
    client,
    model,
    source_lang,
    target_lang,
    mode="bilingual",
    progress=None,
):
    """
    Dịch workbook trực tiếp trên bản sao.
    Không tạo workbook mới.
    """

    keep_vba = filename.lower().endswith(".xlsm")

    wb = load_workbook(
        io.BytesIO(uploaded_bytes),
        data_only=False,
        keep_vba=keep_vba,
    )

    locations = []
    texts = []

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value

                if not is_translatable_text(value):
                    continue

                if already_bilingual(
                    value,
                    source_lang,
                    target_lang,
                ):
                    continue

                locations.append(
                    (ws.title, cell.coordinate)
                )
                texts.append(value.strip())

    if not texts:
        if progress:
            progress(1.0)
        return save_workbook(wb)

    results = translate_texts(
        client,
        model,
        texts,
        source_lang,
        target_lang,
        progress=progress,
    )

    missing = [
        i for i in range(len(texts))
        if i not in results
    ]

    if missing:
        raise RuntimeError(
            "Một số ô chưa có bản dịch: "
            + ", ".join(map(str, missing[:20]))
        )

    for idx, (sheet_name, coordinate) in enumerate(locations):
        ws = wb[sheet_name]
        cell = ws[coordinate]

        translation = results[idx]

        if mode == "bilingual":
            cell.value = bilingual_text(
                cell.value,
                translation,
            )

            preserve_alignment_with_wrap(cell)

            # Chỉ tăng chiều cao nếu chưa được cố định.
            row_dim = ws.row_dimensions[cell.row]

            if row_dim.height is None:
                old_height = 15
                row_dim.height = max(
                    old_height,
                    30,
                )

        else:
            cell.value = translation

    return save_workbook(wb)


def save_workbook(wb):
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ============================================================
# IMAGE OCR
# ============================================================

def prepare_image_bytes(image_bytes, filename):
    """
    Đọc ảnh, tự động xoay theo EXIF và xuất PNG.
    Giúp Gemini nhận ảnh ổn định hơn.
    """
    image = Image.open(io.BytesIO(image_bytes))
    image = ImageOps.exif_transpose(image)

    if image.mode not in ("RGB", "RGBA"):
        image = image.convert("RGB")

    out = io.BytesIO()
    image.save(out, format="PNG")

    return out.getvalue()


def extract_table_from_image_once(
    client,
    model,
    image_bytes,
):
    prompt = r"""
Bạn là chuyên gia OCR bảng biểu Trung - Việt.

Hãy đọc chính xác bảng trong ảnh và trả về JSON theo schema.

MỤC TIÊU:
1. Nhận dạng số hàng và số cột.
2. Nhận dạng từng ô.
3. Giữ nguyên thứ tự hàng/cột.
4. Giữ nguyên chữ, số, mã, ngày, đơn vị.
5. Ô trống phải là "".
6. Không tự sáng tạo dữ liệu.
7. Không đưa chữ trang trí bên ngoài bảng vào.
8. Nhận dạng tiêu đề bảng nếu có.
9. Nếu bảng có merge cell, dùng rowspan và colspan.
10. row và col bắt đầu từ 0.
11. rowspan mặc định là 1.
12. colspan mặc định là 1.
13. Không tạo hai cell đại diện cho cùng một vùng merge.
14. Nếu một ô có nhiều dòng, giữ bằng ký tự xuống dòng.
15. Nếu không chắc chắn về dữ liệu, ưu tiên giữ nguyên những gì nhìn thấy,
    không tự đoán.
16. rows không cần trả về; chỉ trả cells.

Ví dụ:

{
  "title": "2026 年08月26日员工上班",
  "row_count": 4,
  "col_count": 6,
  "cells": [
    {
      "row": 0,
      "col": 0,
      "value": "STT",
      "rowspan": 2,
      "colspan": 1
    },
    {
      "row": 0,
      "col": 1,
      "value": "部门",
      "rowspan": 2,
      "colspan": 1
    },
    {
      "row": 0,
      "col": 2,
      "value": "人员",
      "rowspan": 1,
      "colspan": 2
    }
  ]
}

Chỉ trả JSON theo schema.
"""

    image = Image.open(io.BytesIO(image_bytes))

    response = client.models.generate_content(
        model=model,
        contents=[
            image,
            prompt,
        ],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=ImageTable,
        ),
    )

    data = response.parsed

    if data is not None:
        return data

    return ImageTable.model_validate_json(
        response.text.strip()
    )


def validate_image_table(table):
    if not table.cells:
        raise ValueError(
            "Gemini không nhận dạng được ô nào trong bảng."
        )

    max_row = 0
    max_col = 0

    for cell in table.cells:
        if cell.rowspan < 1 or cell.colspan < 1:
            raise ValueError("OCR trả về rowspan/colspan không hợp lệ.")

        max_row = max(
            max_row,
            cell.row + cell.rowspan,
        )

        max_col = max(
            max_col,
            cell.col + cell.colspan,
        )

    row_count = max(table.row_count, max_row)
    col_count = max(table.col_count, max_col)

    if row_count <= 0 or col_count <= 0:
        raise ValueError("Không xác định được kích thước bảng.")

    return row_count, col_count


def extract_table_from_image(
    client,
    model,
    image_bytes,
):
    last_error = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            table = extract_table_from_image_once(
                client,
                model,
                image_bytes,
            )

            validate_image_table(table)

            return table

        except Exception as exc:
            last_error = exc

            if attempt < MAX_RETRIES:
                time.sleep(2 ** (attempt - 1))

    raise RuntimeError(
        f"OCR ảnh thất bại sau {MAX_RETRIES} lần thử: {last_error}"
    )


# ============================================================
# IMAGE -> EXCEL HELPERS
# ============================================================

def normalize_image_cells(table):
    """
    Loại cell trùng vị trí bắt đầu.
    Nếu Gemini vô tình trả trùng, giữ cell đầu tiên.
    """
    unique = {}
    for cell in table.cells:
        key = (cell.row, cell.col)

        if key not in unique:
            unique[key] = cell

    return list(unique.values())


def collect_image_texts(
    table,
    source_lang,
    target_lang,
    mode,
):
    texts = []
    locations = []

    cells = normalize_image_cells(table)

    for cell in cells:
        value = cell.value

        if not is_translatable_text(value):
            continue

        if already_bilingual(
            value,
            source_lang,
            target_lang,
        ):
            continue

        texts.append(value.strip())
        locations.append((cell.row, cell.col))

    title_index = None

    if (
        table.title
        and is_translatable_text(table.title)
        and not already_bilingual(
            table.title,
            source_lang,
            target_lang,
        )
    ):
        title_index = len(texts)
        texts.append(table.title.strip())

    return texts, locations, title_index


def make_image_excel(
    table,
    translations,
    cell_translation_map,
    title_translation,
    mode,
):
    row_count, col_count = validate_image_table(table)

    wb = Workbook()
    ws = wb.active
    ws.title = "Translated"

    # --------------------------------------------------------
    # TITLE
    # --------------------------------------------------------

    has_title = bool(table.title.strip())

    start_row = 1

    if has_title:
        title_value = table.title

        if mode == "bilingual" and title_translation:
            title_value = bilingual_text(
                table.title,
                title_translation,
            )

        elif mode == "translated" and title_translation:
            title_value = title_translation

        ws.merge_cells(
            start_row=1,
            start_column=1,
            end_row=1,
            end_column=col_count,
        )

        title_cell = ws.cell(
            1,
            1,
            title_value,
        )

        title_cell.font = Font(
            name="Arial",
            size=16,
            bold=True,
            color=BLACK,
        )

        title_cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

        ws.row_dimensions[1].height = (
            44 if mode == "bilingual" else 30
        )

        start_row = 2

    # --------------------------------------------------------
    # TABLE
    # --------------------------------------------------------

    thin = Side(
        style="thin",
        color=THIN_GRAY,
    )

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin,
    )

    orange_fill = PatternFill(
        fill_type="solid",
        fgColor=HEADER_ORANGE,
    )

    cells = normalize_image_cells(table)

    # Tạo toàn bộ cell trước.
    for r in range(row_count):
        for c in range(col_count):
            excel_cell = ws.cell(
                start_row + r,
                c + 1,
            )

            excel_cell.border = border
            excel_cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

    # Ghi dữ liệu.
    for image_cell in cells:
        r = image_cell.row
        c = image_cell.col

        if r >= row_count or c >= col_count:
            continue

        excel_row = start_row + r
        excel_col = c + 1

        value = image_cell.value

        translation = cell_translation_map.get(
            (r, c),
            "",
        )

        if translation:
            if mode == "bilingual":
                output_value = bilingual_text(
                    value,
                    translation,
                )
            else:
                output_value = translation
        else:
            output_value = value

        excel_cell = ws.cell(
            excel_row,
            excel_col,
            output_value,
        )

        excel_cell.border = border
        excel_cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

        # Header:
        # các cell nằm ở hàng đầu tiên của bảng được tô cam.
        if r == 0:
            excel_cell.fill = orange_fill
            excel_cell.font = Font(
                name="Arial",
                size=11,
                bold=True,
                color=WHITE,
            )
        else:
            excel_cell.font = Font(
                name="Arial",
                size=11,
                color=BLACK,
            )

    # --------------------------------------------------------
    # MERGE
    # --------------------------------------------------------

    merged_ranges = set()

    for image_cell in cells:
        if (
            image_cell.rowspan == 1
            and image_cell.colspan == 1
        ):
            continue

        start_excel_row = start_row + image_cell.row
        start_excel_col = image_cell.col + 1

        end_excel_row = (
            start_excel_row
            + image_cell.rowspan
            - 1
        )

        end_excel_col = (
            start_excel_col
            + image_cell.colspan
            - 1
        )

        if (
            end_excel_row > start_row + row_count - 1
            or end_excel_col > col_count
        ):
            continue

        range_ref = (
            f"{get_column_letter(start_excel_col)}"
            f"{start_excel_row}:"
            f"{get_column_letter(end_excel_col)}"
            f"{end_excel_row}"
        )

        if range_ref in merged_ranges:
            continue

        try:
            ws.merge_cells(range_ref)
            merged_ranges.add(range_ref)
        except Exception:
            # Nếu Gemini nhận dạng merge xung đột,
            # không làm hỏng toàn bộ file.
            pass

    # --------------------------------------------------------
    # WIDTH / HEIGHT
    # --------------------------------------------------------

    for col in range(1, col_count + 1):
        letter = get_column_letter(col)
        max_len = 0

        for row in range(
            start_row,
            start_row + row_count,
        ):
            value = ws.cell(row, col).value

            if value is None:
                continue

            longest = max(
                [len(line) for line in str(value).splitlines()]
                or [0]
            )

            max_len = max(
                max_len,
                longest,
            )

        ws.column_dimensions[letter].width = min(
            max(8, max_len + 3),
            40,
        )

    for row in range(
        start_row,
        start_row + row_count,
    ):
        max_lines = 1
        max_chars = 0

        for col in range(1, col_count + 1):
            value = ws.cell(row, col).value

            if value is None:
                continue

            text = str(value)

            max_lines = max(
                max_lines,
                text.count("\n") + 1,
            )

            max_chars = max(
                max_chars,
                max(
                    [len(x) for x in text.splitlines()]
                    or [0]
                ),
            )

        height = 20 * max_lines

        if max_chars > 45:
            height += 10

        ws.row_dimensions[row].height = min(
            max(22, height),
            90,
        )

    ws.sheet_view.showGridLines = False

    if row_count > 1:
        ws.freeze_panes = f"A{start_row + 1}"

    return save_workbook(wb)


def build_excel_from_image(
    image_bytes,
    original_filename,
    client,
    model,
    source_lang,
    target_lang,
    mode="bilingual",
    progress=None,
):
    prepared_bytes = prepare_image_bytes(
        image_bytes,
        original_filename,
    )

    if progress:
        progress(0.05)

    table = extract_table_from_image(
        client,
        model,
        prepared_bytes,
    )

    if progress:
        progress(0.20)

    texts, locations, title_index = collect_image_texts(
        table,
        source_lang,
        target_lang,
        mode,
    )

    translations = translate_texts(
        client,
        model,
        texts,
        source_lang,
        target_lang,
        progress=(
            lambda p: progress(
                0.20 + p * 0.65
            )
            if progress
            else None
        ),
    )

    title_translation = ""

    if title_index is not None:
        title_translation = translations.get(
            title_index,
            "",
        )

    cell_translation_map = {}

    for index, location in enumerate(locations):
        if index in translations:
            cell_translation_map[location] = translations[index]

    if progress:
        progress(0.90)

    output_bytes = make_image_excel(
        table,
        translations,
        cell_translation_map,
        title_translation,
        mode,
    )

    if progress:
        progress(1.0)

    return output_bytes, table


# ============================================================
# FILE TYPE
# ============================================================

def extension_type(filename):
    lower = filename.lower()

    if lower.endswith((".xlsx", ".xlsm")):
        return "excel"

    if lower.endswith(
        (".png", ".jpg", ".jpeg", ".webp")
    ):
        return "image"

    return None


# ============================================================
# UI
# ============================================================

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🌐",
    layout="wide",
)

st.title(APP_TITLE)

st.caption(
    "Dịch Trung ↔ Việt bằng Gemini • "
    "Excel giữ tối đa cấu trúc/định dạng gốc • "
    "Ảnh bảng OCR + nhận dạng merge + xuất Excel"
)

with st.sidebar:
    st.header("⚙️ Cài đặt")

    source_label = st.selectbox(
        "Ngôn ngữ nguồn",
        [
            "中文 — Tiếng Trung",
            "Tiếng Việt",
        ],
        index=0,
    )

    source_lang = (
        "zh"
        if source_label.startswith("中文")
        else "vi"
    )

    target_lang = (
        "vi"
        if source_lang == "zh"
        else "zh"
    )

    direction = (
        "Trung → Việt"
        if source_lang == "zh"
        else "Việt → Trung"
    )

    st.info(
        f"Chiều dịch: **{direction}**"
    )

    mode_label = st.radio(
        "Kiểu xuất",
        [
            "Song ngữ — nguyên văn + bản dịch",
            "Chỉ bản dịch",
        ],
        index=0,
    )

    mode = (
        "bilingual"
        if mode_label.startswith("Song ngữ")
        else "translated"
    )

    model = st.selectbox(
        "Gemini model",
        [
            "gemini-3.7-flash",
            "gemini-3.6-flash",
            "gemini-3.5-flash",
        ],
        index=0,
    )

    api_key = get_api_key()

    if not api_key:
        api_key = st.text_input(
            "Gemini API Key",
            type="password",
            help=(
                "Có thể nhập tạm thời hoặc cấu hình "
                "GEMINI_API_KEY trong Streamlit Secrets."
            ),
        )
    else:
        st.success("Đã tìm thấy GEMINI_API_KEY.")

    st.divider()

    st.markdown(
        """
### Định dạng hỗ trợ

- Excel: `.xlsx`, `.xlsm`
- Ảnh: `.png`, `.jpg`, `.jpeg`, `.webp`

### Excel

- Sửa trực tiếp trên bản sao
- Giữ merge và style của ô
- Giữ font, màu, border
- Giữ width/height
- Giữ freeze panes
- Giữ các thuộc tính sheet chính
- XLSM: cố gắng bảo toàn VBA bằng `keep_vba=True`

> OpenPyXL không đảm bảo bảo toàn 100% mọi đối tượng Excel
> phức tạp như chart/drawing/slicer/embedded object.

### Ảnh

- Gemini Vision OCR
- Nhận dạng hàng/cột
- Nhận dạng rowspan/colspan
- Dựng merge cell
- Dịch song ngữ
- Xuất Excel
        """
    )


uploaded = st.file_uploader(
    "📤 Kéo thả file vào đây",
    type=[
        "xlsx",
        "xlsm",
        "png",
        "jpg",
        "jpeg",
        "webp",
    ],
)

if uploaded:
    file_type = extension_type(uploaded.name)

    if file_type == "image":
        st.subheader("🖼️ Ảnh nguồn")

        image_bytes = uploaded.getvalue()

        col1, col2 = st.columns([1, 1])

        with col1:
            st.image(
                image_bytes,
                caption="Ảnh gốc",
                use_container_width=True,
            )

        with col2:
            st.markdown(
                """
### Quy trình

1. Chuẩn hóa ảnh theo EXIF.
2. Gemini Vision đọc bảng.
3. Nhận dạng cell, hàng, cột.
4. Nhận dạng merge bằng `rowspan/colspan`.
5. Dịch nội dung chữ.
6. Dựng Excel.
7. Xuất file `.xlsx`.

### Song ngữ

Ví dụ:

```text
员工姓名
Họ tên nhân viên
"""
            )

        if st.button(
            "🚀 OCR + DỊCH + XUẤT EXCEL",
            type="primary",
            use_container_width=True,
        ):
            if not api_key:
                st.error(
                    "Chưa có GEMINI_API_KEY. "
                    "Hãy nhập API key ở sidebar "
                    "hoặc thêm vào Streamlit Secrets."
                )
                st.stop()

            try:
                client = get_client(api_key)

                progress = st.progress(0)

                with st.spinner(
                    "Gemini đang OCR, nhận dạng bảng và dịch..."
                ):
                    (
                        output_bytes,
                        table,
                    ) = build_excel_from_image(
                        image_bytes,
                        uploaded.name,
                        client,
                        model,
                        source_lang,
                        target_lang,
                        mode,
                        progress=progress.progress,
                    )

                row_count, col_count = (
                    validate_image_table(table)
                )

                st.success(
                    f"Hoàn tất. Nhận dạng khoảng "
                    f"{row_count} hàng × {col_count} cột, "
                    f"{len(table.cells)} vùng ô."
                )

                st.subheader(
                    "🔎 Dữ liệu OCR"
                )

                preview = [
                    [
                        ""
                        for _ in range(col_count)
                    ]
                    for _ in range(row_count)
                ]

                for cell in normalize_image_cells(table):
                    if (
                        cell.row < row_count
                        and cell.col < col_count
                    ):
                        preview[
                            cell.row
                        ][
                            cell.col
                        ] = cell.value

                st.dataframe(
                    preview,
                    use_container_width=True,
                    hide_index=True,
                )

                out_name = (
                    Path(uploaded.name).stem
                    + "_"
                    + (
                        "song_ngu"
                        if mode == "bilingual"
                        else "da_dich"
                    )
                    + ".xlsx"
                )

                st.download_button(
                    "⬇️ TẢI FILE EXCEL",
                    data=output_bytes,
                    file_name=out_name,
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    type="primary",
                    use_container_width=True,
                )

            except Exception as exc:
                st.error(
                    f"Lỗi xử lý ảnh: {exc}"
                )
                st.exception(exc)

    elif file_type == "excel":
        st.subheader("📊 Excel nguồn")

        excel_bytes = uploaded.getvalue()

        st.info(
            "Chương trình sửa trực tiếp trên bản sao workbook: "
            "giữ tối đa cấu trúc và định dạng Excel gốc; "
            "chỉ thay nội dung chữ bằng bản dịch."
        )

        if uploaded.name.lower().endswith(".xlsm"):
            st.warning(
                "Đây là XLSM. Chương trình dùng keep_vba=True "
                "để cố gắng bảo toàn VBA. Hãy kiểm tra macro sau khi xuất."
            )

        if st.button(
            "🚀 DỊCH EXCEL + GIỮ ĐỊNH DẠNG",
            type="primary",
            use_container_width=True,
        ):
            if not api_key:
                st.error(
                    "Chưa có GEMINI_API_KEY. "
                    "Hãy nhập API key ở sidebar "
                    "hoặc thêm vào Streamlit Secrets."
                )
                st.stop()

            try:
                client = get_client(api_key)

                progress = st.progress(0)

                with st.spinner(
                    "Đang đọc Excel và dịch..."
                ):
                    output_bytes = translate_excel(
                        excel_bytes,
                        uploaded.name,
                        client,
                        model,
                        source_lang,
                        target_lang,
                        mode,
                        progress=progress.progress,
                    )

                progress.progress(1.0)

                st.success(
                    "Dịch Excel hoàn tất."
                )

                suffix = (
                    "_song_ngu"
                    if mode == "bilingual"
                    else "_da_dich"
                )

                original_ext = (
                    ".xlsm"
                    if uploaded.name.lower().endswith(".xlsm")
                    else ".xlsx"
                )

                out_name = (
                    Path(uploaded.name).stem
                    + suffix
                    + original_ext
                )

                mime = (
                    "application/vnd.ms-excel.sheet."
                    "macroEnabled.12"
                    if original_ext == ".xlsm"
                    else
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                )

                st.download_button(
                    "⬇️ TẢI EXCEL SAU KHI DỊCH",
                    data=output_bytes,
                    file_name=out_name,
                    mime=mime,
                    type="primary",
                    use_container_width=True,
                )

            except Exception as exc:
                st.error(
                    f"Lỗi xử lý Excel: {exc}"
                )
                st.exception(exc)
else:
    st.markdown("")
