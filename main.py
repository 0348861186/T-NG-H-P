import io
import os
import re
import json
import time
import copy
import math
import zipfile
from pathlib import Path

import streamlit as st
from PIL import Image, ImageDraw, ImageFont

import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.cell.cell import MergedCell

from google import genai
from google.genai import types


# ============================================================
# STREAMLIT CONFIG
# ============================================================

st.set_page_config(
    page_title="Gemini AI - Dịch Excel Song Ngữ",
    page_icon="🌐",
    layout="wide",
)


# ============================================================
# CONSTANTS
# ============================================================

PRIMARY_MODEL = "gemini-2.5-flash"
FALLBACK_MODEL = "gemini-2.5-flash-lite"

SUPPORTED_EXTENSIONS = [
    "xlsx",
    "xlsm",
    "png",
    "jpg",
    "jpeg",
]

IMAGE_EXTENSIONS = {"png", "jpg", "jpeg"}
EXCEL_EXTENSIONS = {"xlsx", "xlsm"}

MAX_TEXT_BATCH = 40
MAX_RETRIES = 3


# ============================================================
# UI
# ============================================================

st.title("🌐 Dịch File Song Ngữ Trung ↔ Việt bằng Gemini AI")

st.caption(
    "Excel: giữ cấu trúc/định dạng tối đa và thêm bản dịch bên dưới. "
    "Ảnh: OCR bằng Gemini và thêm tiếng Việt bên dưới tiếng Trung."
)

st.markdown("---")


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.header("⚙️ Cấu hình")

direction = st.sidebar.selectbox(
    "Chọn hướng dịch:",
    [
        "Trung → Việt",
        "Việt → Trung",
    ],
)

st.sidebar.markdown("---")

st.sidebar.subheader("🤖 Gemini AI")

api_key_input = st.sidebar.text_input(
    "Gemini API Key",
    type="password",
    placeholder="AIza...",
    help="Có thể để trống nếu đã cấu hình GEMINI_API_KEY trong Streamlit Secrets.",
)

if direction == "Trung → Việt":
    st.sidebar.success(
        "Kết quả: tiếng Trung ở trên, tiếng Việt ở dưới."
    )
else:
    st.sidebar.success(
        "Kết quả cuối: tiếng Trung ở trên, tiếng Việt ở dưới."
    )


# ============================================================
# GET API KEY
# ============================================================

def get_api_key():
    """
    Ưu tiên:
    1. API key nhập trên sidebar
    2. st.secrets["GEMINI_API_KEY"]
    3. biến môi trường GEMINI_API_KEY
    """
    if api_key_input and api_key_input.strip():
        return api_key_input.strip()

    try:
        secret_key = st.secrets.get("GEMINI_API_KEY")
        if secret_key:
            return str(secret_key).strip()
    except Exception:
        pass

    env_key = os.getenv("GEMINI_API_KEY")

    if env_key:
        return env_key.strip()

    return ""


API_KEY = get_api_key()


# ============================================================
# GEMINI CLIENT
# ============================================================

@st.cache_resource(show_spinner=False)
def create_gemini_client(api_key):
    if not api_key:
        return None
    return genai.Client(api_key=api_key)


client = create_gemini_client(API_KEY)


# ============================================================
# HELPERS
# ============================================================

def clean_text(value):
    if value is None:
        return ""
    return str(value).strip()


def is_formula(value):
    if not isinstance(value, str):
        return False
    return value.startswith("=")


def is_number(value):
    if value is None:
        return False

    if isinstance(value, (int, float)):
        return True

    text = str(value).strip()

    if not text:
        return False

    patterns = [
        r"^-?\d+$",
        r"^-?\d+\.\d+$",
        r"^-?\d+,\d+$",
        r"^-?\d{1,3}(,\d{3})+(\.\d+)?$",
        r"^-?\d{1,3}(\.\d{3})+(,\d+)?$",
        r"^-?\d+(\.\d+)?%$",
    ]

    return any(re.match(p, text) for p in patterns)


def looks_like_code_or_id(text):
    """
    Không dịch các mã thuần kỹ thuật.
    """
    if not text:
        return False

    text = str(text).strip()

    if len(text) <= 2:
        return False

    if re.fullmatch(r"[A-Z]{2,10}[-_/]?\d{2,}", text):
        return True

    if re.fullmatch(r"[A-Z0-9_-]{5,}", text):
        return True

    return False


def should_translate_cell(value):
    """
    Chỉ dịch text thực sự.
    Không dịch: công thức, số, ô rỗng, mã kỹ thuật rõ ràng.
    """
    if value is None:
        return False

    if isinstance(value, bool):
        return False

    if isinstance(value, (int, float)):
        return False

    text = str(value).strip()

    if not text:
        return False

    if is_formula(text):
        return False

    if is_number(text):
        return False

    if looks_like_code_or_id(text):
        return False

    return True


# ============================================================
# GEMINI JSON CALL
# ============================================================

def extract_json(text):
    """
    Gemini đôi khi trả về markdown JSON. Hàm này cố lấy JSON ra an toàn.
    """
    if not text:
        raise ValueError("Gemini không trả về dữ liệu.")

    text = text.strip()

    try:
        return json.loads(text)
    except Exception:
        pass

    text = re.sub(r"^```json\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*```$", "", text, flags=re.IGNORECASE)

    try:
        return json.loads(text)
    except Exception:
        pass

    first_obj = text.find("{")
    last_obj = text.rfind("}")

    if first_obj >= 0 and last_obj > first_obj:
        candidate = text[first_obj:last_obj + 1]
        try:
            return json.loads(candidate)
        except Exception:
            pass

    first_list = text.find("[")
    last_list = text.rfind("]")

    if first_list >= 0 and last_list > first_list:
        candidate = text[first_list:last_list + 1]
        try:
            return json.loads(candidate)
        except Exception:
            pass

    raise ValueError("Không thể đọc JSON từ phản hồi Gemini.")


# ============================================================
# GEMINI TRANSLATION
# ============================================================

def gemini_generate(prompt, model, image=None, mime_type=None):
    """
    Gọi Gemini với retry.
    """
    if client is None:
        raise RuntimeError("Chưa cấu hình GEMINI_API_KEY.")

    last_error = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            if image is None:
                response = client.models.generate_content(
                    model=model,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        temperature=0,
                        response_mime_type="application/json",
                    ),
                )
            else:
                image_part = types.Part.from_bytes(
                    data=image,
                    mime_type=mime_type,
                )
                response = client.models.generate_content(
                    model=model,
                    contents=[prompt, image_part],
                    config=types.GenerateContentConfig(
                        temperature=0,
                        response_mime_type="application/json",
                    ),
                )

            if not response or not response.text:
                raise RuntimeError("Gemini trả về phản hồi rỗng.")

            return response.text, model

        except Exception as exc:
            last_error = exc
            error_text = str(exc).lower()

            temporary_error = any(
                keyword in error_text
                for keyword in [
                    "503",
                    "unavailable",
                    "429",
                    "resource exhausted",
                    "timeout",
                    "deadline",
                    "internal",
                    "overloaded",
                ]
            )

            if attempt < MAX_RETRIES and temporary_error:
                time.sleep(2 * attempt)
                continue

            raise last_error


def translate_batch_with_gemini(items):
    """
    Dịch danh sách các đoạn text theo batch qua Gemini.
    """
    if not items:
        return {}

    source_name = "Chinese" if direction == "Trung → Việt" else "Vietnamese"
    target_name = "Vietnamese" if direction == "Trung → Việt" else "Chinese"

    payload = json.dumps(items, ensure_ascii=False)

    prompt = f"""
You are a professional Chinese-Vietnamese translator.
 
Translation direction:
{source_name} -> {target_name}
 
Translate ONLY the text values.
 
Important rules:
1. Preserve product names, technical terminology and manufacturing terminology accurately.
2. Do not add explanations.
3. Do not summarize.
4. Do not change numbers.
5. Do not change units.
6. Do not change model numbers.
7. Keep the same item IDs.
8. Return valid JSON only.
9. Every input ID must have exactly one translated value.
 
Input JSON:
{payload}
 
Output JSON format:
{{
 "translations": [
    {{
      "id": 1,
     "translation": "..."
    }}
  ]
}}
"""

    # PRIMARY MODEL
    try:
        raw, used_model = gemini_generate(
            prompt=prompt,
            model=PRIMARY_MODEL,
        )

        data = extract_json(raw)
        translations = data.get("translations", [])

        result = {}
        for item in translations:
            item_id = item.get("id")
            translated = item.get("translation")

            if item_id is not None:
                result[int(item_id)] = str(translated) if translated is not None else ""

        if len(result) < len(items):
            raise ValueError("Gemini trả thiếu bản dịch.")

        return result, used_model

    except Exception as primary_error:
        # FALLBACK MODEL
        try:
            raw, used_model = gemini_generate(
                prompt=prompt,
                model=FALLBACK_MODEL,
            )

            data = extract_json(raw)
            translations = data.get("translations", [])

            result = {}
            for item in translations:
                item_id = item.get("id")
                translated = item.get("translation")

                if item_id is not None:
                    result[int(item_id)] = str(translated) if translated is not None else ""

            if len(result) < len(items):
                raise ValueError("Fallback Gemini trả thiếu bản dịch.")

            return result, used_model

        except Exception as fallback_error:
            raise RuntimeError(
                "Gemini chính và Gemini fallback đều thất bại.\n\n"
                f"Primary: {primary_error}\n"
                f"Fallback: {fallback_error}"
            )


# ============================================================
# EXCEL STYLE HELPERS
# ============================================================

def copy_cell_style(src, dst):
    if src.has_style:
        dst._style = copy.copy(src._style)

    if src.number_format:
        dst.number_format = src.number_format

    if src.protection:
        dst.protection = copy.copy(src.protection)

    if src.alignment:
        dst.alignment = copy.copy(src.alignment)

    if src.font:
        dst.font = copy.copy(src.font)

    if src.fill:
        dst.fill = copy.copy(src.fill)

    if src.border:
        dst.border = copy.copy(src.border)


def copy_row_dimension(ws, source_row, target_row):
    src = ws.row_dimensions[source_row]
    dst = ws.row_dimensions[target_row]

    if src.height is not None:
        dst.height = src.height

    dst.hidden = src.hidden
    dst.outlineLevel = src.outlineLevel
    dst.collapsed = src.collapsed


# ============================================================
# EXCEL PROCESSING
# ============================================================

def collect_excel_translation_items(wb):
    items = []
    counter = 1

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue

                value = cell.value

                if not should_translate_cell(value):
                    continue

                items.append(
                    {
                        "id": counter,
                        "sheet": ws.title,
                        "row": cell.row,
                        "column": cell.column,
                        "text": str(value),
                    }
                )

                counter += 1

    return items


def translate_excel_texts(wb, progress_callback=None):
    all_items = collect_excel_translation_items(wb)

    if not all_items:
        return {}, PRIMARY_MODEL, 0

    translations = {}
    total = len(all_items)
    used_models = []

    batches = [
        all_items[i:i + MAX_TEXT_BATCH]
        for i in range(0, total, MAX_TEXT_BATCH)
    ]

    for batch_index, batch in enumerate(batches):
        batch_input = [
            {
                "id": item["id"],
                "text": item["text"],
            }
            for item in batch
        ]

        batch_result, used_model = translate_batch_with_gemini(batch_input)

        translations.update(batch_result)
        used_models.append(used_model)

        if progress_callback:
            progress_callback(
                min(
                    1.0,
                    (batch_index + 1) / len(batches),
                )
            )

    used_model_final = (
        max(
            set(used_models),
            key=used_models.count,
        )
        if used_models
        else PRIMARY_MODEL
    )

    return translations, used_model_final, total


def snapshot_workbook_structure(wb):
    snapshot = {}

    for ws in wb.worksheets:
        snapshot[ws.title] = {
            "max_row": ws.max_row,
            "max_column": ws.max_column,
            "column_widths": {
                key: ws.column_dimensions[key].width
                for key in ws.column_dimensions
            },
            "row_heights": {
                r: ws.row_dimensions[r].height
                for r in range(1, ws.max_row + 1)
                if ws.row_dimensions[r].height is not None
            },
            "row_hidden": {
                r: ws.row_dimensions[r].hidden
                for r in range(1, ws.max_row + 1)
                if ws.row_dimensions[r].hidden
            },
            "merged_ranges": [
                str(x) for x in ws.merged_cells.ranges
            ],
            "freeze_panes": ws.freeze_panes,
        }

    return snapshot


def build_translation_rows(wb, translation_map, original_items):
    rows_by_sheet = {}

    for item in original_items:
        item_id = item["id"]
        if item_id not in translation_map:
            continue
        sheet = item["sheet"]
        row = item["row"]
        rows_by_sheet.setdefault(sheet, set()).add(row)

    for ws in wb.worksheets:
        target_rows = rows_by_sheet.get(ws.title, set())
        if not target_rows:
            continue

        original_max_col = ws.max_column

        for row in sorted(target_rows, reverse=True):
            merged_before = []
            for merged in list(ws.merged_cells.ranges):
                min_col, min_row, max_col, max_row = range_boundaries(str(merged))
                if min_row <= row <= max_row:
                    merged_before.append((min_col, min_row, max_col, max_row))

            for merged in list(ws.merged_cells.ranges):
                min_col, min_row, max_col, max_row = range_boundaries(str(merged))
                if min_row <= row <= max_row:
                    ws.unmerge_cells(str(merged))

            ws.insert_rows(row + 1, amount=1)
            copy_row_dimension(ws, row, row + 1)

            for col in range(1, original_max_col + 1):
                src = ws.cell(row, col)
                dst = ws.cell(row + 1, col)

                copy_cell_style(src, dst)

                if src.hyperlink:
                    dst._hyperlink = copy.copy(src.hyperlink)
                if src.comment:
                    dst.comment = copy.copy(src.comment)

                dst.value = src.value

            for item in original_items:
                if item["sheet"] != ws.title:
                    continue
                if item["row"] != row:
                    continue

                item_id = item["id"]
                if item_id not in translation_map:
                    continue

                col = item["column"]
                src = ws.cell(row, col)
                dst = ws.cell(row + 1, col)

                original_text = item["text"]
                translated_text = translation_map[item_id]

                if direction == "Trung → Việt":
                    src.value = original_text
                    dst.value = translated_text
                else:
                    src.value = translated_text
                    dst.value = original_text

            for (min_col, min_row, max_col, max_row) in merged_before:
                if min_row == row and max_row == row:
                    range_original = f"{get_column_letter(min_col)}{row}:{get_column_letter(max_col)}{row}"
                    range_translation = f"{get_column_letter(min_col)}{row + 1}:{get_column_letter(max_col)}{row + 1}"

                    try:
                        ws.merge_cells(range_original)
                    except Exception:
                        pass
                    try:
                        ws.merge_cells(range_translation)
                    except Exception:
                        pass
                else:
                    shifted_min_row = min_row + (1 if min_row > row else 0)
                    shifted_max_row = max_row + 1 if max_row >= row else max_row

                    new_range = f"{get_column_letter(min_col)}{shifted_min_row}:{get_column_letter(max_col)}{shifted_max_row}"

                    try:
                        ws.merge_cells(new_range)
                    except Exception:
                        pass

            if ws.row_dimensions[row].hidden:
                ws.row_dimensions[row + 1].hidden = True


def apply_translations_to_excel(wb, translation_map, original_items):
    build_translation_rows(
        wb=wb,
        translation_map=translation_map,
        original_items=original_items,
    )
    return wb


# ============================================================
# IMAGE OCR + TRANSLATION
# ============================================================

def image_font(size, bold=False):
    candidates = []

    if os.name == "nt":
        candidates.extend(
            [
                r"C:\Windows\Fonts\msyh.ttc",
                r"C:\Windows\Fonts\simhei.ttf",
                r"C:\Windows\Fonts\arial.ttf",
            ]
        )

    candidates.extend(
        [
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Bold.ttc",
        ]
    )

    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size=size)
            except Exception:
                pass

    return ImageFont.load_default()


def ocr_image_with_gemini(image_bytes, mime_type):
    prompt = """
Analyze the provided image as an OCR engine.
 
Detect all visible Chinese or Vietnamese text.
 
For each text region return:
- id
- original_text
- x1
- y1
- x2
- y2
 
Coordinates must be normalized from 0 to 1000.
 
Important:
1. Do not invent text.
2. Preserve the exact original text.
3. Ignore purely decorative elements.
4. Include text from tables, labels, forms and documents.
5. Keep each logical text line as a separate item.
 
Return JSON only:
 
{
  "items": [
    {
      "id": 1,
     "original_text": "...",
      "x1": 100,
      "y1": 100,
      "x2": 400,
      "y2": 150
    }
  ]
}
"""

    raw, used_model = gemini_generate(
        prompt=prompt,
        model=PRIMARY_MODEL,
        image=image_bytes,
        mime_type=mime_type,
    )

    data = extract_json(raw)
    items = data.get("items", [])

    return items, used_model


def translate_image_items(items):
    valid_items = []

    for item in items:
        text = clean_text(item.get("original_text", ""))

        if not text:
            continue

        if not should_translate_cell(text):
            continue

        valid_items.append(
            {
                "id": int(item.get("id", len(valid_items) + 1)),
                "text": text,
            }
        )

    if not valid_items:
        return {}, PRIMARY_MODEL

    result, used_model = translate_batch_with_gemini(valid_items)

    return result, used_model


def draw_text_with_background(draw, xy, text, font, fill, background, padding=4):
    x, y = xy
    bbox = draw.textbbox((x, y), text, font=font)

    draw.rectangle(
        (
            bbox[0] - padding,
            bbox[1] - padding,
            bbox[2] + padding,
            bbox[3] + padding,
        ),
        fill=background,
    )

    draw.text((x, y), text, font=font, fill=fill)


def process_image(image_bytes, mime_type):
    image = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    width, height = image.size

    ocr_items, ocr_model = ocr_image_with_gemini(image_bytes, mime_type)
    translations, translation_model = translate_image_items(ocr_items)

    output = image.copy()
    draw = ImageDraw.Draw(output)

    default_font = image_font(max(16, int(width / 80)))

    for item in ocr_items:
        item_id = int(item.get("id", 0))
        original_text = clean_text(item.get("original_text", ""))

        if not original_text:
            continue

        translated = translations.get(item_id, "")

        if not translated:
            continue

        x1 = float(item.get("x1", 0)) / 1000 * width
        y1 = float(item.get("y1", 0)) / 1000 * height
        x2 = float(item.get("x2", 0)) / 1000 * width
        y2 = float(item.get("y2", 0)) / 1000 * height

        box_height = max(1, y2 - y1)
        font_size = max(14, int(box_height * 0.8))
        font = image_font(font_size)

        if direction == "Trung → Việt":
            tx = x1
            ty = y2 + 4

            if ty + font_size + 10 > height:
                ty = max(0, y1 - font_size - 8)

            draw_text_with_background(
                draw=draw,
                xy=(tx, ty),
                text=translated,
                font=font,
                fill=(0, 80, 0),
                background=(255, 255, 255),
            )
        else:
            tx = x1
            ty = y1 - font_size - 8

            if ty < 0:
                ty = y2 + 4

            draw_text_with_background(
                draw=draw,
                xy=(tx, ty),
                text=translated,
                font=font,
                fill=(0, 0, 150),
                background=(255, 255, 255),
            )

    output_buffer = io.BytesIO()
    output.save(output_buffer, format="PNG")
    output_buffer.seek(0)

    return (
        output_buffer.getvalue(),
        ocr_model,
        translation_model,
        len(ocr_items),
    )


# ============================================================
# FILE UPLOAD
# ============================================================

uploaded_file = st.file_uploader(
    "📁 Tải file lên",
    type=SUPPORTED_EXTENSIONS,
    help="Hỗ trợ XLSX, XLSM, PNG, JPG, JPEG.",
)


# ============================================================
# MAIN PROCESS
# ============================================================

if uploaded_file is not None:
    extension = uploaded_file.name.rsplit(".", 1)[-1].lower()

    st.info(f"📄 File: **{uploaded_file.name}**")

    if not API_KEY:
        st.error("❌ Chưa có GEMINI_API_KEY.")
        st.markdown(
            """
### Cách cấu hình

**Cách 1 – Nhập API key trực tiếp ở Sidebar**

**Cách 2 – Cấu hình trên Streamlit Cloud (Secrets):**
```toml
GEMINI_API_KEY = "AIza...""""
)
        st.stop()
    # ========================================================
    # EXCEL
    # ========================================================
    if extension in EXCEL_EXTENSIONS:
        st.subheader("📊 Xử lý Excel")

        keep_vba = True if extension == "xlsm" else False

        process_button = st.button(
            "🚀 BẮT ĐẦU DỊCH EXCEL",
            type="primary",
            use_container_width=True,
        )

        if process_button:
            try:
                file_bytes = uploaded_file.getvalue()
                input_stream = io.BytesIO(file_bytes)

                wb = load_workbook(
                    input_stream,
                    data_only=False,
                    keep_vba=keep_vba,
                )

                st.write(f"📑 Số sheet: **{len(wb.sheetnames)}**")
                for name in wb.sheetnames:
                    st.write(f"• {name}")

                structure_snapshot = snapshot_workbook_structure(wb)
                original_items = collect_excel_translation_items(wb)
                total_items = len(original_items)

                st.write(f"🔤 Số ô text cần dịch: **{total_items}**")

                if total_items == 0:
                    st.warning("Không tìm thấy nội dung văn bản cần dịch.")
                    st.stop()

                progress = st.progress(0)
                status = st.empty()
                status.info("🤖 Đang gửi nội dung cho Gemini...")

                def update_progress(value):
                    progress.progress(value)

                translation_map, used_model, translated_count = translate_excel_texts(
                    wb, progress_callback=update_progress
                )

                status.success(f"✅ Gemini hoàn tất: {translated_count} mục.")
                status.info("🛠️ Đang xây dựng lại Excel...")

                wb = apply_translations_to_excel(
                    wb=wb,
                    translation_map=translation_map,
                    original_items=original_items,
                )

                status.info("💾 Đang tạo file Excel...")
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)

                progress.progress(1.0)
                status.success("🎉 Hoàn tất!")
                st.success(f"Đã dịch bằng **{used_model}**.")

                st.download_button(
                    label="📥 DOWNLOAD EXCEL SONG NGỮ",
                    data=output.getvalue(),
                    file_name=f"dich_song_ngu_{Path(uploaded_file.name).stem}.{extension}",
                    mime=(
                        "application/vnd.ms-excel.sheet.macroEnabled.12"
                        if extension == "xlsm"
                        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ),
                    use_container_width=True,
                    type="primary",
                )

                st.markdown("---")
                st.subheader("ℹ️ Thông tin xử lý")

                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Sheets", len(wb.sheetnames))
                with col2:
                    st.metric("Ô đã dịch", translated_count)
                with col3:
                    st.metric("Model", used_model)

                st.caption(
                    "Lưu ý: Excel phức tạp có VBA, PivotTable, SmartArt, slicer hoặc "
                    "drawing đặc biệt có thể cần kiểm tra lại bằng Excel sau khi xuất."
                )

            except Exception as e:
                st.error("❌ Xử lý Excel thất bại.")
                st.exception(e)

    # ========================================================
    # IMAGE
    # ========================================================
    elif extension in IMAGE_EXTENSIONS:
        st.subheader("🖼️ OCR + Dịch hình ảnh bằng Gemini")

        image_bytes = uploaded_file.getvalue()
        mime_type = {
            "png": "image/png",
            "jpg": "image/jpeg",
            "jpeg": "image/jpeg",
        }[extension]

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### Ảnh gốc")
            st.image(image_bytes, use_container_width=True)

        process_image_button = st.button(
            "🚀 OCR + DỊCH HÌNH ẢNH",
            type="primary",
            use_container_width=True,
        )

        if process_image_button:
            try:
                progress = st.progress(0)
                status = st.empty()
                status.info("👁️ Gemini đang nhận dạng chữ...")
                progress.progress(20)

                translated_image, ocr_model, translation_model, ocr_count = process_image(
                    image_bytes=image_bytes,
                    mime_type=mime_type,
                )

                progress.progress(100)
                status.success("🎉 OCR + dịch hoàn tất!")

                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("### Ảnh gốc")
                    st.image(image_bytes, use_container_width=True)
                with col2:
                    st.markdown("### Ảnh song ngữ")
                    st.image(translated_image, use_container_width=True)

                st.success(
                    f"OCR: **{ocr_model}** | "
                    f"Dịch: **{translation_model}** | "
                    f"Vùng text: **{ocr_count}**"
                )

                output_name = f"dich_song_ngu_{Path(uploaded_file.name).stem}.png"

                st.download_button(
                    label="📥 DOWNLOAD ẢNH SONG NGỮ",
                    data=translated_image,
                    file_name=output_name,
                    mime="image/png",
                    use_container_width=True,
                    type="primary",
                )

            except Exception as e:
                st.error("❌ OCR/Dịch hình ảnh thất bại.")
                st.exception(e)
