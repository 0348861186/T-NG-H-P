from pathlib import Path

code = r'''import io
import json
import mimetypes
import os
import re
import tempfile
from copy import copy

import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import RowDimension

from google import genai
from google.genai import types


# ============================================================
# STREAMLIT APP - GEMINI TRANSLATOR
# Excel + Image
#
# Quy tac:
# 1. Excel: giu nguyen workbook, sheet, merge, font, border,
#    fill, alignment, number format, column width... va chen
#    dong dich ngay ben duoi dong goc.
# 2. Image: giu nguyen anh goc, mo rong canvas xuong duoi,
#    viet them dong dich ben duoi tung dong OCR.
# 3. Hien thi theo quy tac CUOI CUNG:
#       TIENG TRUNG
#       TIENG VIET
#    Tuc la tieng Viet luon nam ngay ben duoi tieng Trung.
# 4. Chi dung Gemini de OCR/nhan dien va dich.
# ============================================================


APP_TITLE = "Gemini AI - Dịch Trung ↔ Việt"
DEFAULT_MODEL = "gemini-3.7-flash"

IMAGE_EXTENSIONS = ["png", "jpg", "jpeg", "webp", "bmp"]
EXCEL_EXTENSIONS = ["xlsx", "xlsm"]
ALL_EXTENSIONS = IMAGE_EXTENSIONS + EXCEL_EXTENSIONS


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🌏",
    layout="wide",
)


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
    <style>
    .main-title {
        font-size: 34px;
        font-weight: 800;
        margin-bottom: 4px;
    }

    .sub-title {
        color: #666;
        font-size: 15px;
        margin-bottom: 20px;
    }

    .info-box {
        padding: 14px 18px;
        border-radius: 10px;
        border: 1px solid #ddd;
        background: #fafafa;
        margin-bottom: 15px;
    }

    .rule-box {
        padding: 15px 18px;
        border-radius: 10px;
        border: 1px solid #d9e2f3;
        background: #f5f8ff;
        margin-top: 10px;
        margin-bottom: 20px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# HELPERS
# ============================================================

def get_api_key():
    """
    Lay GEMINI_API_KEY tu Streamlit secrets hoac environment.
    """
    try:
        key = st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        key = ""

    if not key:
        key = os.environ.get("GEMINI_API_KEY", "")

    return key.strip()


def get_client():
    api_key = get_api_key()

    if not api_key:
        raise RuntimeError(
            "Chua co GEMINI_API_KEY. Hay them GEMINI_API_KEY vao "
            "Streamlit Secrets hoac bien moi truong."
        )

    return genai.Client(api_key=api_key)


def clean_json_text(text):
    """
    Gemini doi khi tra ve JSON trong markdown code fence.
    """
    if not text:
        return ""

    text = text.strip()

    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.I)
        text = re.sub(r"\s*```$", "", text)

    return text.strip()


def parse_json_response(text):
    text = clean_json_text(text)

    try:
        return json.loads(text)
    except Exception:
        # Tim object/array JSON dau tien.
        start_candidates = [
            p for p in [text.find("{"), text.find("[")] if p >= 0
        ]

        if not start_candidates:
            raise ValueError(
                "Gemini khong tra ve JSON hop le:\n" + text
            )

        start = min(start_candidates)

        # Thu cat den dau cuoi hop ly.
        for end in range(len(text), start, -1):
            candidate = text[start:end].strip()
            try:
                return json.loads(candidate)
            except Exception:
                continue

        raise ValueError(
            "Khong the parse JSON tu Gemini:\n" + text
        )


def is_formula(value):
    return isinstance(value, str) and value.startswith("=")


def is_empty(value):
    return value is None or str(value).strip() == ""


def safe_text(value):
    if value is None:
        return ""
    return str(value)


def split_lines_for_translation(text):
    """
    Tach theo dong nhung van giu nguyen cac dong rong.
    """
    return safe_text(text).splitlines()


# ============================================================
# GEMINI TRANSLATION
# ============================================================

def translate_texts_gemini(
    client,
    texts,
    direction,
    model_name=DEFAULT_MODEL,
):
    """
    Dich danh sach text bang Gemini.

    direction:
        zh_vi = Trung -> Viet
        vi_zh = Viet -> Trung

    Ket qua:
        dict[index] = translated text
    """

    if not texts:
        return {}

    source_name = "Chinese" if direction == "zh_vi" else "Vietnamese"
    target_name = "Vietnamese" if direction == "zh_vi" else "Chinese"

    # Loc text co noi dung.
    indexed = []
    for i, value in enumerate(texts):
        value = safe_text(value)
        if value.strip():
            indexed.append((i, value))

    if not indexed:
        return {}

    # Chia batch de tranh prompt qua lon.
    batch_size = 80
    result = {}

    progress = st.progress(0, text="Dang dich...")

    total_batches = (len(indexed) + batch_size - 1) // batch_size

    for batch_no, start in enumerate(range(0, len(indexed), batch_size), start=1):
        batch = indexed[start:start + batch_size]

        payload = [
            {
                "id": idx,
                "text": text,
            }
            for idx, text in batch
        ]

        prompt = f"""
You are a professional Chinese-Vietnamese technical translator.

Translate the following list from {source_name} to {target_name}.

IMPORTANT:
- Preserve IDs exactly.
- Return ONLY valid JSON.
- JSON must be an array of objects.
- Each object must have exactly: "id" and "translation".
- Do not add explanations.
- Do not merge items.
- Do not omit any item.
- Preserve numbers, model names, part numbers, dimensions,
  units, punctuation, and technical terminology whenever appropriate.
- Do not translate formulas or programming expressions as natural language.
- If the input is already in the target language, return it unchanged.
- If a cell contains multiple lines, preserve line breaks where possible.
- Do not put markdown around the JSON.

Input:
{json.dumps(payload, ensure_ascii=False)}
"""

        response = client.models.generate_content(
            model=model_name,
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.1,
                response_mime_type="application/json",
            ),
        )

        data = parse_json_response(response.text)

        if not isinstance(data, list):
            raise ValueError(
                "Gemini tra ve JSON khong dung dang danh sach."
            )

        for item in data:
            if not isinstance(item, dict):
                continue

            idx = item.get("id")
            translation = item.get("translation", "")

            try:
                idx = int(idx)
            except Exception:
                continue

            result[idx] = safe_text(translation)

        progress.progress(
            min(batch_no / total_batches, 1.0),
            text=f"Dang dich batch {batch_no}/{total_batches}...",
        )

    progress.empty()

    # Neu Gemini bo sot item nao, thu dich lai tung item.
    missing = [idx for idx, _ in indexed if idx not in result]

    if missing:
        retry_payload = [
            {
                "id": idx,
                "text": dict(indexed)[idx],
            }
            for idx in missing
        ]

        prompt = f"""
Translate from {source_name} to {target_name}.

Return ONLY JSON array:
[
  {{"id": 0, "translation": "..."}}
]

Do not omit anything.
Preserve technical terms, numbers, units and line breaks.

{json.dumps(retry_payload, ensure_ascii=False)}
"""

        response = client.models.generate_content(
            model=model_name,
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.1,
                response_mime_type="application/json",
            ),
        )

        retry_data = parse_json_response(response.text)

        if isinstance(retry_data, list):
            for item in retry_data:
                if isinstance(item, dict):
                    try:
                        idx = int(item.get("id"))
                        result[idx] = safe_text(item.get("translation", ""))
                    except Exception:
                        pass

    return result


# ============================================================
# EXCEL - STYLE COPY
# ============================================================

def copy_cell_style(source, target):
    """
    Copy style cua cell ma khong lam thay doi workbook goc.
    """
    if isinstance(source, MergedCell):
        return

    if isinstance(target, MergedCell):
        return

    if source.has_style:
        target._style = copy(source._style)

    if source.number_format:
        target.number_format = source.number_format

    if source.alignment:
        target.alignment = copy(source.alignment)

    if source.protection:
        target.protection = copy(source.protection)

    if source.font:
        target.font = copy(source.font)

    if source.fill:
        target.fill = copy(source.fill)

    if source.border:
        target.border = copy(source.border)


def copy_row_format(ws, source_row, target_row, max_col):
    """
    Copy dinh dang cua ca dong.
    """
    src_dim = ws.row_dimensions[source_row]
    dst_dim = ws.row_dimensions[target_row]

    if src_dim.height is not None:
        dst_dim.height = src_dim.height

    dst_dim.hidden = src_dim.hidden
    dst_dim.outlineLevel = src_dim.outlineLevel
    dst_dim.collapsed = src_dim.collapsed

    for col in range(1, max_col + 1):
        source = ws.cell(source_row, col)
        target = ws.cell(target_row, col)

        if not isinstance(target, MergedCell):
            copy_cell_style(source, target)


def copy_sheet_view_settings(source_ws, target_ws):
    """
    Khong bat buoc nhung giup giu giao dien worksheet.
    """
    try:
        target_ws.sheet_view.showGridLines = source_ws.sheet_view.showGridLines
    except Exception:
        pass

    try:
        target_ws.freeze_panes = source_ws.freeze_panes
    except Exception:
        pass


def get_merged_ranges_touching_row(ws, row):
    ranges = []

    for merged in list(ws.merged_cells.ranges):
        if merged.min_row <= row <= merged.max_row:
            ranges.append(merged)

    return ranges


def copy_column_dimensions(ws):
    """
    Column dimensions duoc giu san khi insert_rows.
    Ham nay chi dam bao cac thuoc tinh quan trong khong bi mat.
    """
    dims = {}

    for key, dim in ws.column_dimensions.items():
        dims[key] = {
            "width": dim.width,
            "hidden": dim.hidden,
            "bestFit": dim.bestFit,
            "outlineLevel": dim.outlineLevel,
            "collapsed": dim.collapsed,
        }

    return dims


def restore_column_dimensions(ws, dims):
    for key, values in dims.items():
        dim = ws.column_dimensions[key]

        if values["width"] is not None:
            dim.width = values["width"]

        dim.hidden = values["hidden"]
        dim.bestFit = values["bestFit"]
        dim.outlineLevel = values["outlineLevel"]
        dim.collapsed = values["collapsed"]


def build_excel_translation(
    uploaded_bytes,
    direction,
    client,
    model_name=DEFAULT_MODEL,
):
    """
    Tao workbook moi tu workbook goc.

    Quy tac hien thi:
        Dong Trung
        Dong Viet

    Neu file goc la tieng Trung:
        dong goc = Trung
        dong chen = Viet

    Neu file goc la tieng Viet:
        dong chen = Trung
        dong goc = Viet

    Nhu vay tieng Viet luon nam ngay duoi tieng Trung.
    """

    input_stream = io.BytesIO(uploaded_bytes)

    # keep_vba=True de giu macro neu la XLSM.
    # data_only=False de giu cong thuc.
    # rich_text=True co the khong co tren mot so version openpyxl,
    # nen dung try/except.
    try:
        wb = load_workbook(
            input_stream,
            keep_vba=True,
            data_only=False,
            rich_text=True,
        )
    except TypeError:
        wb = load_workbook(
            input_stream,
            keep_vba=True,
            data_only=False,
        )

    # --------------------------------------------------------
    # Thu thap text.
    # --------------------------------------------------------

    jobs = []
    all_texts = []

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue

                value = cell.value

                # Khong dich formula.
                if is_empty(value) or is_formula(value):
                    continue

                if not isinstance(value, str):
                    continue

                # Bo qua chuoi qua ngan neu chi la ky hieu.
                text = value.strip()

                if not text:
                    continue

                # Khong can dich nhung gia tri thuong la ma/so.
                if re.fullmatch(r"[\d\s\W_]+", text, flags=re.UNICODE):
                    continue

                job_id = len(all_texts)

                all_texts.append(text)
                jobs.append(
                    {
                        "job_id": job_id,
                        "sheet": ws.title,
                        "row": cell.row,
                        "col": cell.column,
                        "text": text,
                    }
                )

    if not jobs:
        return workbook_to_bytes(wb), 0

    translations = translate_texts_gemini(
        client=client,
        texts=all_texts,
        direction=direction,
        model_name=model_name,
    )

    # --------------------------------------------------------
    # Gom theo sheet + row.
    #
    # Muc dich: moi dong Excel goc se co 1 dong dich ngay ben duoi.
    # --------------------------------------------------------

    jobs_by_sheet_row = {}

    for job in jobs:
        translated = translations.get(job["job_id"], "")

        key = (job["sheet"], job["row"])

        jobs_by_sheet_row.setdefault(key, []).append(
            {
                **job,
                "translation": translated,
            }
        )

    # Phai xu ly tu duoi len de insert_rows khong lam sai row index.
    sheet_rows = {}

    for (sheet_name, row), row_jobs in jobs_by_sheet_row.items():
        sheet_rows.setdefault(sheet_name, set()).add(row)

    # --------------------------------------------------------
    # Xu ly tung worksheet.
    # --------------------------------------------------------

    for ws in wb.worksheets:
        rows_to_process = sorted(
            sheet_rows.get(ws.title, set()),
            reverse=True,
        )

        if not rows_to_process:
            continue

        max_col = max(ws.max_column, 1)

        original_column_dims = copy_column_dimensions(ws)

        # Luu merge ranges hien tai.
        original_merges = list(ws.merged_cells.ranges)

        for original_row in rows_to_process:
            key = (ws.title, original_row)
            row_jobs = jobs_by_sheet_row[key]

            # ------------------------------------------------
            # Ghi nho merge range cua dong goc.
            # ------------------------------------------------

            affected_merges = []

            for merged in list(ws.merged_cells.ranges):
                if merged.min_row == original_row and merged.max_row == original_row:
                    affected_merges.append(merged)

            # ------------------------------------------------
            # Chen 1 dong ngay ben duoi.
            # ------------------------------------------------

            ws.insert_rows(original_row + 1, 1)

            translated_row = original_row + 1

            # Copy format row.
            copy_row_format(
                ws,
                original_row,
                translated_row,
                max_col,
            )

            # Copy height.
            try:
                ws.row_dimensions[translated_row].height = (
                    ws.row_dimensions[original_row].height
                )
            except Exception:
                pass

            # ------------------------------------------------
            # Xu ly merge ngang.
            #
            # Muc tieu:
            # Original:
            # A1:C1 merged
            #
            # Sau:
            # A1:C1 merged
            # A2:C2 merged
            # ------------------------------------------------

            for merged in affected_merges:
                min_col = merged.min_col
                max_col_merged = merged.max_col

                # Insert rows co the lam thay doi merge ranges.
                # Xoa merge target neu ton tai.
                target_range = (
                    f"{get_column_letter(min_col)}{translated_row}:"
                    f"{get_column_letter(max_col_merged)}{translated_row}"
                )

                try:
                    ws.unmerge_cells(target_range)
                except Exception:
                    pass

                try:
                    ws.merge_cells(target_range)
                except Exception:
                    pass

            # ------------------------------------------------
            # Ghi translation.
            #
            # Trung luon tren, Viet luon duoi.
            # ------------------------------------------------

            for job in row_jobs:
                col = job["col"]
                source_value = job["text"]
                translated = job["translation"]

                if not translated:
                    translated = ""

                if direction == "zh_vi":
                    # Goc Trung, dich Viet.
                    # Dong goc nam tren.
                    target = ws.cell(translated_row, col)
                    target.value = translated

                else:
                    # Goc Viet.
                    # Can dua Trung len tren, Viet xuong duoi.
                    #
                    # De khong pha workbook qua manh, ta doi noi dung
                    # cua 2 dong cho cac cell co noi dung.
                    source_cell = ws.cell(original_row, col)
                    target_cell = ws.cell(translated_row, col)

                    # Dong tren = Trung.
                    source_cell.value = translated

                    # Dong duoi = Viet goc.
                    target_cell.value = source_value

            # ------------------------------------------------
            # Neu row co nhieu cell, nhung chi 1 cell co text,
            # cac cell con lai van duoc style.
            # ------------------------------------------------

            # Giu row hidden/trang thai.
            try:
                ws.row_dimensions[translated_row].hidden = (
                    ws.row_dimensions[original_row].hidden
                )
            except Exception:
                pass

        restore_column_dimensions(ws, original_column_dims)

        copy_sheet_view_settings(ws, ws)

    # --------------------------------------------------------
    # Sau khi insert row, cac merge phuc tap co the bi thay doi.
    #
    # Openpyxl xu ly merge co han che. Phan duoi day co gang
    # khoi phuc merge ngang cua cac row da dich.
    # --------------------------------------------------------

    return workbook_to_bytes(wb), len(jobs)


def workbook_to_bytes(wb):
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ============================================================
# IMAGE OCR + TRANSLATION
# ============================================================

IMAGE_SCHEMA = {
    "type": "array",
    "items": {
        "type": "object",
        "properties": {
            "text": {
                "type": "string",
            },
            "x": {
                "type": "number",
            },
            "y": {
                "type": "number",
            },
            "width": {
                "type": "number",
            },
            "height": {
                "type": "number",
            },
        },
        "required": [
            "text",
            "x",
            "y",
            "width",
            "height",
        ],
    },
}


def image_ocr_gemini(
    client,
    image_bytes,
    mime_type,
    model_name=DEFAULT_MODEL,
):
    """
    Gemini nhan dien text + toa do bounding box.
    Toa do duoc yeu cau theo pixel cua anh.
    """

    image = Image.open(io.BytesIO(image_bytes)).convert("RGB")

    prompt = """
Analyze this image as an OCR/document-layout task.

Find ALL visible human-readable text that should be translated.

Return ONLY a JSON array.

Each object:
{
  "text": "exact text",
  "x": number,
  "y": number,
  "width": number,
  "height": number
}

Coordinates must be in pixels relative to the original image:
- x = left
- y = top
- width = text bounding box width
- height = text bounding box height

Important:
- Preserve the original reading order.
- Do not invent text.
- Do not translate.
- Include Chinese and Vietnamese text.
- Include text inside tables/forms when readable.
- Exclude decorative noise when it is not actual text.
"""

    response = client.models.generate_content(
        model=model_name,
        contents=[
            prompt,
            image,
        ],
        config=types.GenerateContentConfig(
            temperature=0.0,
            response_mime_type="application/json",
        ),
    )

    data = parse_json_response(response.text)

    if not isinstance(data, list):
        raise ValueError("OCR Gemini khong tra ve danh sach.")

    cleaned = []

    img_w, img_h = image.size

    for item in data:
        if not isinstance(item, dict):
            continue

        text = safe_text(item.get("text", "")).strip()

        if not text:
            continue

        try:
            x = float(item.get("x", 0))
            y = float(item.get("y", 0))
            width = float(item.get("width", 0))
            height = float(item.get("height", 0))
        except Exception:
            continue

        # Clamp toa do.
        x = max(0, min(x, img_w - 1))
        y = max(0, min(y, img_h - 1))
        width = max(1, min(width, img_w - x))
        height = max(1, min(height, img_h - y))

        cleaned.append(
            {
                "text": text,
                "x": x,
                "y": y,
                "width": width,
                "height": height,
            }
        )

    return cleaned


def choose_font(size, bold=False):
    """
    Tim font Unicode pho bien tren Windows/Linux.
    """
    candidates = []

    if os.name == "nt":
        candidates.extend(
            [
                r"C:\Windows\Fonts\arial.ttf",
                r"C:\Windows\Fonts\arialuni.ttf",
                r"C:\Windows\Fonts\msyh.ttc",
                r"C:\Windows\Fonts\simhei.ttf",
                r"C:\Windows\Fonts\tahoma.ttf",
            ]
        )
    else:
        candidates.extend(
            [
                "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
                "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
                "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            ]
        )

    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size=size)
            except Exception:
                pass

    return ImageFont.load_default()


def wrap_text(draw, text, font, max_width):
    """
    Word wrap theo pixel.
    """
    if not text:
        return [""]

    # Neu text la tieng Trung, khong co space -> wrap theo ky tu.
    if " " not in text:
        chunks = []
        current = ""

        for char in text:
            candidate = current + char

            bbox = draw.textbbox(
                (0, 0),
                candidate,
                font=font,
            )

            if bbox[2] - bbox[0] <= max_width:
                current = candidate
            else:
                if current:
                    chunks.append(current)
                current = char

        if current:
            chunks.append(current)

        return chunks or [""]

    words = text.split()
    lines = []
    current = ""

    for word in words:
        candidate = word if not current else current + " " + word

        bbox = draw.textbbox(
            (0, 0),
            candidate,
            font=font,
        )

        if bbox[2] - bbox[0] <= max_width:
            current = candidate
        else:
            if current:
                lines.append(current)
            current = word

    if current:
        lines.append(current)

    return lines or [""]


def create_translated_image(
    client,
    image_bytes,
    direction,
    model_name=DEFAULT_MODEL,
):
    """
    Tao anh dich.

    De giu anh goc toi da:
    - Anh goc duoc giu nguyen.
    - Khong xoa text goc.
    - Mo rong canvas xuong phia duoi.
    - Chen cap:
          Trung
          Viet

    voi quy tac tieng Viet luon nam duoi tieng Trung.
    """

    image = Image.open(io.BytesIO(image_bytes)).convert("RGB")

    # OCR.
    with st.spinner("Gemini dang nhan dien noi dung trong anh..."):
        boxes = image_ocr_gemini(
            client=client,
            image_bytes=image_bytes,
            mime_type="image/" + (
                "jpeg"
                if image.format == "JPEG"
                else "png"
            ),
            model_name=model_name,
        )

    if not boxes:
        raise ValueError("Khong tim thay noi dung text trong anh.")

    texts = [item["text"] for item in boxes]

    with st.spinner("Gemini dang dich noi dung trong anh..."):
        translations = translate_texts_gemini(
            client=client,
            texts=texts,
            direction=direction,
            model_name=model_name,
        )

    for i, item in enumerate(boxes):
        item["translation"] = translations.get(i, "")

    # --------------------------------------------------------
    # Tinh khu vuc them vao.
    # --------------------------------------------------------

    original_w, original_h = image.size

    # Khoang cach giua cac cap text.
    gap = 10

    # Chi can them mot khu vuc duoi anh.
    # Moi cap se duoc dat ben duoi bounding box.
    extra_height = 0

    prepared = []

    for item in boxes:
        text = item["text"]
        translation = item.get("translation", "")

        x = int(item["x"])
        y = int(item["y"])
        width = max(30, int(item["width"]))
        height = max(12, int(item["height"]))

        font_size = max(
            12,
            min(
                48,
                int(height * 0.85),
            ),
        )

        font = choose_font(font_size)

        # Kich thuoc dong dich.
        translated_lines = wrap_text(
            ImageDraw.Draw(image),
            translation,
            font,
            max(100, width),
        )

        line_height = max(
            height,
            int(font_size * 1.25),
        )

        translation_height = (
            len(translated_lines) * line_height
            + gap * 2
        )

        # Vi tri cap text dat ben duoi text goc.
        top = y + height + gap

        needed_bottom = top + translation_height

        prepared.append(
            {
                **item,
                "x": x,
                "y": y,
                "width": width,
                "height": height,
                "font": font,
                "lines": translated_lines,
                "translation_top": top,
                "translation_height": translation_height,
            }
        )

        extra_height = max(
            extra_height,
            needed_bottom - original_h,
        )

    # Them padding.
    canvas_extra = int(max(80, extra_height + 30))

    canvas = Image.new(
        "RGB",
        (original_w, original_h + canvas_extra),
        "white",
    )

    # Giữ nguyên ảnh gốc 100%.
    canvas.paste(image, (0, 0))

    draw = ImageDraw.Draw(canvas)

    # Font fallback cho text.
    separator_font = choose_font(16, bold=True)

    # --------------------------------------------------------
    # Ve text.
    # --------------------------------------------------------

    for item in prepared:
        original_text = item["text"]
        translation = item.get("translation", "")

        x = item["x"]
        y = item["y"]
        width = item["width"]
        height = item["height"]
        font = item["font"]

        # Neu input la Viet -> Trung:
        # Trung la ban dich, Viet la goc.
        #
        # Yeu cau cuoi cung:
        # Trung tren
        # Viet duoi
        #
        # Ta khong ghi de anh goc. Anh goc van giu nguyen.
        # Vung moi o duoi chi hien thi cap Trung/Viet.
        #
        # Neu input la Trung:
        # Trung goc tren
        # Viet dich duoi.
        #
        # De tranh xoa noi dung goc, voi ca 2 truong hop
        # ta chi them cap dich o phan mo rong.
        #
        # Tuy nhien, voi vi->zh, cap trong vung moi la:
        # Trung (dich)
        # Viet (goc)

        if direction == "zh_vi":
            chinese_text = original_text
            vietnamese_text = translation
        else:
            chinese_text = translation
            vietnamese_text = original_text

        # Vi tri cap o duoi text goc.
        top = item["translation_top"]

        # Khung moi.
        padding_x = 6
        padding_y = 4

        # Do rong cap.
        box_width = min(
            original_w - x - 4,
            max(
                width + padding_x * 2,
                160,
            ),
        )

        # Ve nen trang de doc de hon.
        draw.rounded_rectangle(
            [
                x,
                top,
                x + box_width,
                top + item["translation_height"],
            ],
            radius=5,
            fill="white",
            outline="black",
            width=1,
        )

        # Dong Trung.
        chinese_font = item["font"]
        chinese_lines = wrap_text(
            draw,
            chinese_text,
            chinese_font,
            box_width - padding_x * 2,
        )

        current_y = top + padding_y

        for line in chinese_lines:
            draw.text(
                (x + padding_x, current_y),
                line,
                fill="black",
                font=chinese_font,
            )

            bbox = draw.textbbox(
                (0, 0),
                line,
                font=chinese_font,
            )

            current_y += max(
                12,
                bbox[3] - bbox[1] + 2,
            )

        # Dong Viet ngay ben duoi.
        viet_font = choose_font(
            max(
                12,
                int(chinese_font.size * 0.90)
                if hasattr(chinese_font, "size")
                else 14,
            )
        )

        vietnamese_lines = wrap_text(
            draw,
            vietnamese_text,
            viet_font,
            box_width - padding_x * 2,
        )

        current_y += 4

        for line in vietnamese_lines:
            draw.text(
                (x + padding_x, current_y),
                line,
                fill="black",
                font=viet_font,
            )

            bbox = draw.textbbox(
                (0, 0),
                line,
                font=viet_font,
            )

            current_y += max(
                12,
                bbox[3] - bbox[1] + 2,
            )

    output = io.BytesIO()

    # PNG de khong mat chat luong.
    canvas.save(
        output,
        format="PNG",
        optimize=False,
    )

    output.seek(0)

    return output.getvalue(), len(boxes)


# ============================================================
# DOWNLOAD NAME
# ============================================================

def make_output_name(filename, kind):
    base, ext = os.path.splitext(filename)

    if kind == "excel":
        # XLSX output, XLSM giu XLSM.
        if ext.lower() == ".xlsm":
            return base + "_dich.xlsm"
        return base + "_dich.xlsx"

    return base + "_dich.png"


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.header("⚙️ Cấu hình")

    api_key_present = bool(get_api_key())

    if api_key_present:
        st.success("Gemini API: đã cấu hình")
    else:
        st.error("Gemini API: chưa cấu hình")

    model_name = st.text_input(
        "Gemini model",
        value=DEFAULT_MODEL,
        help="Có thể đổi sang model Gemini khác nếu tài khoản của bạn hỗ trợ.",
    )

    direction_label = st.radio(
        "Chế độ dịch",
        [
            "🇨🇳 Trung → Việt",
            "🇻🇳 Việt → Trung",
        ],
    )

    direction = (
        "zh_vi"
        if direction_label.startswith("🇨🇳")
        else "vi_zh"
    )

    st.markdown(
        """
        <div class="rule-box">
        <b>Quy tắc hiển thị:</b><br><br>
        🇨🇳 Tiếng Trung<br>
        🇻🇳 Tiếng Việt<br><br>
        Trong cả hai chế độ, tiếng Việt luôn nằm ngay bên dưới
        tiếng Trung.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.caption(
        "Lưu ý: Gemini dịch nội dung; Python/OpenPyXL chịu trách nhiệm "
        "giữ cấu trúc và định dạng Excel."
    )


# ============================================================
# MAIN UI
# ============================================================

st.markdown(
    '<div class="main-title">🌏 Gemini AI - Dịch Trung ↔ Việt</div>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="sub-title">'
    "Tải Excel hoặc hình ảnh lên → dịch → xuất file."
    "</div>",
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="info-box">
    <b>Excel:</b> giữ sheet, công thức, merge, font, màu, border,
    alignment, độ rộng cột và các thuộc tính workbook quan trọng;
    nội dung dịch được chèn thành dòng ngay bên dưới dòng gốc.<br><br>

    <b>Hình ảnh:</b> ảnh gốc được giữ nguyên, không ghi đè nội dung;
    app mở rộng phần dưới ảnh để đặt bản dịch.
    </div>
    """,
    unsafe_allow_html=True,
)


uploaded_file = st.file_uploader(
    "📁 Chọn file cần dịch",
    type=ALL_EXTENSIONS,
    help="Hỗ trợ PNG, JPG, JPEG, WEBP, BMP, XLSX, XLSM.",
)


if uploaded_file is None:
    st.info("Hãy tải file lên để bắt đầu.")
    st.stop()


file_name = uploaded_file.name
file_ext = os.path.splitext(file_name)[1].lower()

st.write(f"**File:** `{file_name}`")

if file_ext in [".xlsx", ".xlsm"]:
    file_type = "excel"
else:
    file_type = "image"


# ============================================================
# PREVIEW
# ============================================================

if file_type == "image":
    try:
        preview_image = Image.open(uploaded_file)
        st.image(
            preview_image,
            caption="Ảnh gốc",
            use_container_width=True,
        )
    except Exception as e:
        st.error(f"Không thể đọc hình ảnh: {e}")
        st.stop()

else:
    st.success(
        "Đã nhận file Excel. Khi dịch, app sẽ tạo một bản sao "
        "và chèn dòng dịch bên dưới nội dung gốc."
    )


# ============================================================
# TRANSLATE BUTTON
# ============================================================

translate_clicked = st.button(
    "🚀 BẮT ĐẦU DỊCH",
    type="primary",
    use_container_width=True,
)


if translate_clicked:
    if not get_api_key():
        st.error(
            "Bạn chưa cấu hình GEMINI_API_KEY. "
            "Hãy thêm API key vào Streamlit Secrets."
        )
        st.stop()

    try:
        client = get_client()

        uploaded_file.seek(0)
        input_bytes = uploaded_file.read()

        if not input_bytes:
            st.error("File rỗng.")
            st.stop()

        # ----------------------------------------------------
        # Excel
        # ----------------------------------------------------

        if file_type == "excel":
            with st.spinner(
                "Gemini đang dịch Excel. Vui lòng chờ..."
            ):
                output_bytes, count = build_excel_translation(
                    uploaded_bytes=input_bytes,
                    direction=direction,
                    client=client,
                    model_name=model_name,
                )

            output_name = make_output_name(
                file_name,
                "excel",
            )

            st.session_state["translated_bytes"] = output_bytes
            st.session_state["translated_name"] = output_name
            st.session_state["translated_type"] = "excel"

            st.success(
                f"Đã dịch xong. Đã xử lý khoảng {count} ô có nội dung."
            )

        # ----------------------------------------------------
        # Image
        # ----------------------------------------------------

        else:
            output_bytes, count = create_translated_image(
                client=client,
                image_bytes=input_bytes,
                direction=direction,
                model_name=model_name,
            )

            output_name = make_output_name(
                file_name,
                "image",
            )

            st.session_state["translated_bytes"] = output_bytes
            st.session_state["translated_name"] = output_name
            st.session_state["translated_type"] = "image"

            st.success(
                f"Đã dịch xong. Gemini nhận diện {count} vùng chữ."
            )

    except Exception as e:
        st.error("Có lỗi trong quá trình dịch.")
        st.exception(e)


# ============================================================
# RESULT + DOWNLOAD
# ============================================================

if "translated_bytes" in st.session_state:
    st.divider()

    st.subheader("✅ Kết quả")

    output_bytes = st.session_state["translated_bytes"]
    output_name = st.session_state["translated_name"]
    output_type = st.session_state["translated_type"]

    if output_type == "image":
        result_image = Image.open(
            io.BytesIO(output_bytes)
        )

        st.image(
            result_image,
            caption="Ảnh sau khi dịch",
            use_container_width=True,
        )

        st.download_button(
            label="⬇️ Download ảnh sau dịch",
            data=output_bytes,
            file_name=output_name,
            mime="image/png",
            type="primary",
            use_container_width=True,
        )

    else:
        st.success(
            "File Excel đã sẵn sàng."
        )

        # Download Excel.
        mime = (
            "application/vnd.ms-excel.sheet.macroEnabled.12"
            if output_name.lower().endswith(".xlsm")
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="⬇️ DOWNLOAD FILE EXCEL SAU KHI DỊCH",
            data=output_bytes,
            file_name=output_name,
            mime=mime,
            type="primary",
            use_container_width=True,
        )

        st.caption(
            "File được tạo từ workbook gốc; bản dịch được đặt "
            "ngay dưới dòng gốc."
        )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "Gemini AI Translator • Excel + Image • Chinese ↔ Vietnamese"
)
'''

path = Path("/mnt/data/gemini_streamlit_translator.py")
path.write_text(code, encoding="utf-8")

requirements = """streamlit>=1.45
google-genai>=1.30.0
openpyxl>=3.1.5
Pillow>=10.0.0
"""

req_path = Path("/mnt/data/requirements.txt")
req_path.write_text(requirements, encoding="utf-8")

print(f"Đã tạo: {path}")
print(f"Đã tạo: {req_path}")
