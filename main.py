from pathlib import Path

project = Path("/mnt/data/chung_viet_translator")
project.mkdir(exist_ok=True)

app_code = r'''import io
import json
import math
import re
from copy import copy
from typing import List

import streamlit as st
from PIL import Image
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension
from pydantic import BaseModel, Field
from google import genai
from google.genai import types


# ============================================================
# TRUNG <-> VIỆT EXCEL / IMAGE TRANSLATOR
# - Excel: giữ nguyên format, merge, border, màu, độ rộng cột...
# - Ảnh: OCR + nhận dạng bảng bằng Gemini rồi dựng lại thành Excel
# - Song ngữ: nguyên văn ở dòng 1, bản dịch ở dòng 2 trong cùng ô
# ============================================================

APP_TITLE = "🇨🇳 ↔ 🇻🇳 Trung ↔ Việt Excel & Image Translator"

MODEL_DEFAULT = "gemini-3.7-flash"
BATCH_SIZE = 35

HEADER_ORANGE = "F28C00"
WHITE = "FFFFFF"
BLACK = "000000"
THIN_GRAY = "808080"


class TranslationItem(BaseModel):
    id: int
    translation: str


class TranslationBatch(BaseModel):
    items: List[TranslationItem]


class ImageTable(BaseModel):
    title: str = ""
    rows: List[List[str]] = Field(default_factory=list)


def get_api_key():
    """Lấy API key từ Streamlit Secrets hoặc biến môi trường."""
    try:
        key = st.secrets.get("GEMINI_API_KEY", "")
        if key:
            return key
    except Exception:
        pass
    return ""


@st.cache_resource(show_spinner=False)
def get_client(api_key: str):
    return genai.Client(api_key=api_key)


def is_formula(value):
    return isinstance(value, str) and value.startswith("=")


def is_translatable_text(value):
    if value is None:
        return False
    if not isinstance(value, str):
        return False

    text = value.strip()
    if not text:
        return False

    if is_formula(text):
        return False

    # Ô chỉ có số, ký hiệu, ngày tháng... không cần dịch.
    if re.fullmatch(r"[\d\s.,:/\\%+\-*=()]+", text):
        return False

    return True


def already_bilingual(text):
    """Không dịch lại nếu ô đã có dấu xuống dòng rõ ràng."""
    if not isinstance(text, str):
        return False
    lines = [x.strip() for x in text.splitlines() if x.strip()]
    return len(lines) >= 2


def bilingual_text(original, translation):
    if not translation:
        return original
    if already_bilingual(original):
        return original
    return f"{original}\n{translation}"


def translate_batch(client, model, texts, source_lang, target_lang):
    """Dịch một batch bằng structured output JSON."""
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
Bạn là biên dịch viên chuyên nghiệp Trung - Việt trong môi trường nhà máy,
sản xuất, máy móc, nhân sự, bảng chấm công và biểu mẫu Excel.

Dịch từ {source} sang {target}.

YÊU CẦU:
1. Giữ nguyên ý nghĩa và thuật ngữ kỹ thuật.
2. Không dịch số, mã máy, ký hiệu, đơn vị nếu không cần.
3. Không thêm giải thích.
4. Không thêm dấu ngoặc kép.
5. Giữ thứ tự ID.
6. Chỉ trả về bản dịch cho từng ID.
7. Nếu nội dung là tên cột/bảng, dịch ngắn gọn, tự nhiên như biểu mẫu nhà máy.
8. Với tiếng Trung giản thể, ưu tiên cách dịch tiếng Việt dùng trong nhà máy tại Việt Nam.
9. Không tự ý bỏ chữ.

Danh sách:
{numbered}
"""

    response = client.models.generate_content(
        model=model,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=TranslationBatch,
            temperature=0.1,
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


def translate_texts(client, model, texts, source_lang, target_lang, progress=None):
    """Dịch toàn bộ danh sách theo batch."""
    all_results = {}

    for start in range(0, len(texts), BATCH_SIZE):
        batch = texts[start:start + BATCH_SIZE]
        result = translate_batch(
            client, model, batch, source_lang, target_lang
        )

        for local_id, translation in result.items():
            all_results[start + local_id] = translation

        if progress:
            done = min(start + BATCH_SIZE, len(texts))
            progress(done / len(texts))

    return all_results


def copy_sheet_properties(src, dst):
    """Sao chép các thuộc tính sheet phổ biến."""
    dst.sheet_view.showGridLines = src.sheet_view.showGridLines
    dst.freeze_panes = src.freeze_panes
    dst.sheet_format.defaultRowHeight = src.sheet_format.defaultRowHeight
    dst.sheet_format.defaultColWidth = src.sheet_format.defaultColWidth

    if src.sheet_properties.pageSetUpPr:
        dst.sheet_properties.pageSetUpPr = copy(src.sheet_properties.pageSetUpPr)

    dst.page_margins = copy(src.page_margins)
    dst.page_setup = copy(src.page_setup)
    dst.print_options = copy(src.print_options)

    for key, value in src.column_dimensions.items():
        dst.column_dimensions[key].width = value.width
        dst.column_dimensions[key].hidden = value.hidden

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
):
    """Dịch workbook nhưng giữ format gốc."""
    keep_vba = filename.lower().endswith(".xlsm")

    wb = load_workbook(
        io.BytesIO(uploaded_bytes),
        data_only=False,
        keep_vba=keep_vba,
    )

    # Thu thập tất cả text cần dịch.
    locations = []
    texts = []

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value

                # Không đụng vào công thức.
                if is_translatable_text(value):
                    # Nếu đang là song ngữ, bỏ qua để tránh dịch lặp.
                    if mode == "bilingual" and already_bilingual(value):
                        continue

                    locations.append((ws.title, cell.coordinate))
                    texts.append(value.strip())

    if not texts:
        return save_workbook(wb, filename)

    results = translate_texts(
        client, model, texts, source_lang, target_lang
    )

    # Ghi bản dịch vào đúng ô, không tạo sheet mới.
    for idx, (sheet_name, coordinate) in enumerate(locations):
        ws = wb[sheet_name]
        cell = ws[coordinate]
        translation = results.get(idx, "")

        if mode == "bilingual":
            cell.value = bilingual_text(cell.value, translation)

            # Giữ style gốc nhưng bật wrap text để hiển thị 2 dòng.
            old_alignment = copy(cell.alignment)
            cell.alignment = Alignment(
                horizontal=old_alignment.horizontal,
                vertical=old_alignment.vertical or "center",
                textRotation=old_alignment.textRotation,
                wrap_text=True,
                shrink_to_fit=old_alignment.shrink_to_fit,
                indent=old_alignment.indent,
            )

            # Tăng chiều cao nếu hàng chưa có chiều cao cố định.
            row_dim = ws.row_dimensions[cell.row]
            if row_dim.height is None:
                row_dim.height = 32

        else:
            cell.value = translation

    return save_workbook(wb, filename)


def save_workbook(wb, original_filename):
    out = io.BytesIO()

    if original_filename.lower().endswith(".xlsm"):
        wb.save(out)
        return out.getvalue()

    wb.save(out)
    return out.getvalue()


def clean_json_text(text):
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?", "", text.strip(), flags=re.I)
        text = re.sub(r"```$", "", text.strip())
    return text.strip()


def extract_table_from_image(client, model, image_bytes, mime_type):
    """
    Gemini nhìn ảnh trực tiếp và trả về ma trận bảng.
    Không dùng OCR engine cài thêm nên phù hợp hơn với Streamlit Cloud.
    """
    prompt = """
Bạn là chuyên gia OCR biểu mẫu Trung - Việt.

Hãy đọc ảnh bảng được cung cấp và chuyển chính xác thành JSON.

MỤC TIÊU:
- Nhận dạng tiêu đề bảng nếu có.
- Nhận dạng từng hàng và từng cột.
- Giữ nguyên thứ tự hàng/cột.
- Giữ nguyên số, ngày, mã và nội dung trống.
- Không tự sáng tạo dữ liệu.
- Nếu ô trống thì trả về chuỗi "".
- Nếu một ô có nhiều dòng, nối bằng ký tự "\\n".
- Không đưa các chữ trang trí bên ngoài bảng vào dữ liệu.
- Nếu bảng có ô merge tiêu đề, đưa tiêu đề vào trường "title".
- "rows" phải là ma trận 2 chiều.

Ví dụ cấu trúc:
{
  "title": "2026 年08月26日员工上班",
  "rows": [
    ["STT", "部门", "开机台机", "正式工", "临时工", "备注"],
    ["1", "连机", "5", "3", "2", ""]
  ]
}

Chỉ trả JSON theo schema, không giải thích.
"""

    image = Image.open(io.BytesIO(image_bytes))

    response = client.models.generate_content(
        model=model,
        contents=[image, prompt],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=ImageTable,
            temperature=0.0,
        ),
    )

    data = response.parsed
    if data is not None:
        return data

    return ImageTable.model_validate_json(clean_json_text(response.text))


def auto_col_width(ws, min_width=10, max_width=35):
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        max_len = 0

        for row in ws.iter_rows(min_col=col_idx, max_col=col_idx):
            cell = row[0]
            if cell.value is None:
                continue

            value = str(cell.value)
            longest = max(
                [len(line) for line in value.splitlines()] or [0]
            )
            max_len = max(max_len, longest)

        ws.column_dimensions[letter].width = min(
            max(min_width, max_len + 3),
            max_width,
        )


def set_row_heights(ws):
    for row in range(1, ws.max_row + 1):
        max_lines = 1
        max_chars = 0

        for col in range(1, ws.max_column + 1):
            value = ws.cell(row, col).value
            if value is None:
                continue

            text = str(value)
            max_lines = max(max_lines, text.count("\n") + 1)
            max_chars = max(
                max_chars,
                max([len(x) for x in text.splitlines()] or [0]),
            )

        height = 20 * max_lines
        if max_chars > 45:
            height += 10

        ws.row_dimensions[row].height = max(22, min(height, 90))


def build_excel_from_image(
    image_bytes,
    original_filename,
    client,
    model,
    source_lang,
    target_lang,
    mode="bilingual",
):
    """OCR bảng từ ảnh -> dịch -> dựng Excel giống kiểu ảnh mẫu."""
    mime_type = "image/png"
    lower = original_filename.lower()

    if lower.endswith(".jpg") or lower.endswith(".jpeg"):
        mime_type = "image/jpeg"
    elif lower.endswith(".webp"):
        mime_type = "image/webp"

    table = extract_table_from_image(
        client, model, image_bytes, mime_type
    )

    rows = table.rows
    if not rows:
        raise ValueError("Không nhận dạng được bảng trong ảnh.")

    max_cols = max(len(row) for row in rows)
    normalized_rows = [
        row + [""] * (max_cols - len(row))
        for row in rows
    ]

    # Thu thập text cần dịch.
    texts = []
    locations = []

    for r, row in enumerate(normalized_rows):
        for c, value in enumerate(row):
            if is_translatable_text(value) and not already_bilingual(value):
                texts.append(value.strip())
                locations.append((r, c))

    title = table.title.strip()

    if title and is_translatable_text(title):
        texts.insert(0, title)
        title_index = 0
        location_offset = 1
    else:
        title_index = None
        location_offset = 0

    translations = translate_texts(
        client,
        model,
        texts,
        source_lang,
        target_lang,
    )

    title_translation = ""
    if title_index is not None:
        title_translation = translations.get(0, "")

    # Các translation cho cell bắt đầu từ index location_offset.
    cell_translation = {}
    for i, location in enumerate(locations):
        result_index = i + location_offset
        cell_translation[location] = translations.get(result_index, "")

    wb = Workbook()
    ws = wb.active
    ws.title = "Translated"

    # Tiêu đề giống ảnh mẫu: căn giữa, merge toàn bảng.
    title_row = 1
    header_row = 2

    if title:
        title_value = title
        if mode == "bilingual" and title_translation:
            title_value = f"{title}\n{title_translation}"
        elif mode == "translated" and title_translation:
            title_value = title_translation

        ws.merge_cells(
            start_row=title_row,
            start_column=1,
            end_row=title_row,
            end_column=max_cols,
        )

        title_cell = ws.cell(title_row, 1, title_value)
        title_cell.font = Font(
            name="Arial",
            size=16,
            bold=True,
        )
        title_cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )
        ws.row_dimensions[title_row].height = 42 if mode == "bilingual" else 30

    # Ghi bảng.
    start_row = header_row if title else 1

    thin = Side(style="thin", color=THIN_GRAY)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    orange_fill = PatternFill(
        fill_type="solid",
        fgColor=HEADER_ORANGE,
    )

    for r, row in enumerate(normalized_rows):
        excel_row = start_row + r

        for c, value in enumerate(row):
            cell = ws.cell(excel_row, c + 1)

            if (
                (r, c) in cell_translation
                and cell_translation[(r, c)]
            ):
                if mode == "bilingual":
                    cell.value = bilingual_text(
                        value,
                        cell_translation[(r, c)],
                    )
                else:
                    cell.value = cell_translation[(r, c)]
            else:
                cell.value = value

            cell.border = border
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

            # Header dòng đầu tiên.
            if r == 0:
                cell.fill = orange_fill
                cell.font = Font(
                    name="Arial",
                    size=11,
                    bold=True,
                    color=WHITE,
                )
            else:
                cell.font = Font(
                    name="Arial",
                    size=11,
                    color=BLACK,
                )

    auto_col_width(ws)
    set_row_heights(ws)

    # Định dạng giống bảng mẫu.
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = f"A{start_row + 1}"

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), table


def extension_type(filename):
    lower = filename.lower()
    if lower.endswith(".xlsx") or lower.endswith(".xlsm"):
        return "excel"
    if lower.endswith((".png", ".jpg", ".jpeg", ".webp")):
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
    "Dịch Trung ↔ Việt bằng Gemini • Excel giữ nguyên định dạng • "
    "Ảnh bảng được OCR và dựng lại thành Excel"
)

with st.sidebar:
    st.header("⚙️ Cài đặt")

    source_label = st.selectbox(
        "Ngôn ngữ nguồn",
        ["中文 — Tiếng Trung", "Tiếng Việt"],
        index=0,
    )
    source_lang = "zh" if source_label.startswith("中文") else "vi"
    target_lang = "vi" if source_lang == "zh" else "zh"

    direction = (
        "Trung → Việt"
        if source_lang == "zh"
        else "Việt → Trung"
    )
    st.info(f"Chiều dịch: **{direction}**")

    mode_label = st.radio(
        "Kiểu xuất",
        [
            "Song ngữ — nguyên văn + bản dịch",
            "Chỉ bản dịch",
        ],
        index=0,
    )
    mode = "bilingual" if mode_label.startswith("Song ngữ") else "translated"

    model = st.selectbox(
        "Gemini model",
        [
            "gemini-3.7-flash",
            "gemini-2.5-flash",
        ],
        index=0,
    )

    api_key = get_api_key()

    if not api_key:
        api_key = st.text_input(
            "Gemini API Key",
            type="password",
            help="Có thể nhập tạm thời hoặc cấu hình GEMINI_API_KEY trong Streamlit Secrets.",
        )

    st.divider()
    st.markdown(
        """
**Định dạng hỗ trợ**

- Excel: `.xlsx`, `.xlsm`
- Ảnh: `.png`, `.jpg`, `.jpeg`, `.webp`

**Excel sẽ giữ:** merge, border, màu nền, font,
độ rộng cột, chiều cao dòng, freeze panes và các thuộc tính sheet chính.
        """
    )

uploaded = st.file_uploader(
    "📤 Kéo thả file vào đây",
    type=["xlsx", "xlsm", "png", "jpg", "jpeg", "webp"],
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
**Cách xử lý ảnh**

1. Gemini đọc cấu trúc bảng.
2. Nhận dạng tiêu đề, hàng, cột và ô trống.
3. Dịch các ô có chữ.
4. Giữ số liệu nguyên bản.
5. Tạo Excel với header màu cam, border,
   căn giữa và nội dung song ngữ giống bố cục ảnh mẫu.
                """
            )

        if st.button(
            "🚀 OCR + DỊCH + XUẤT EXCEL",
            type="primary",
            use_container_width=True,
        ):
            if not api_key:
                st.error(
                    "Chưa có GEMINI_API_KEY. Hãy nhập API key ở sidebar "
                    "hoặc thêm vào Streamlit Secrets."
                )
                st.stop()

            try:
                client = get_client(api_key)
                progress = st.progress(0)

                with st.spinner("Gemini đang đọc bảng và dịch..."):
                    output_bytes, table = build_excel_from_image(
                        image_bytes,
                        uploaded.name,
                        client,
                        model,
                        source_lang,
                        target_lang,
                        mode,
                    )

                progress.progress(1.0)

                st.success(
                    f"Hoàn tất. Nhận dạng {len(table.rows)} hàng × "
                    f"{max(len(r) for r in table.rows)} cột."
                )

                st.dataframe(
                    table.rows,
                    use_container_width=True,
                    hide_index=True,
                )

                out_name = (
                    Path(uploaded.name).stem
                    + "_"
                    + ("song_ngu" if mode == "bilingual" else "da_dich")
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
                st.error(f"Lỗi xử lý ảnh: {exc}")
                st.exception(exc)

    elif file_type == "excel":
        st.subheader("📊 Excel nguồn")
        excel_bytes = uploaded.getvalue()

        st.info(
            "Excel sẽ được sửa trực tiếp trên bản sao: "
            "giữ nguyên cấu trúc sheet và định dạng; "
            "chỉ thay nội dung chữ bằng bản dịch."
        )

        if st.button(
            "🚀 DỊCH EXCEL + GIỮ ĐỊNH DẠNG",
            type="primary",
            use_container_width=True,
        ):
            if not api_key:
                st.error(
                    "Chưa có GEMINI_API_KEY. Hãy nhập API key ở sidebar "
                    "hoặc thêm vào Streamlit Secrets."
                )
                st.stop()

            try:
                client = get_client(api_key)
                progress = st.progress(0)

                with st.spinner("Đang đọc Excel và dịch..."):
                    output_bytes = translate_excel(
                        excel_bytes,
                        uploaded.name,
                        client,
                        model,
                        source_lang,
                        target_lang,
                        mode,
                    )

                progress.progress(1.0)

                st.success("Dịch Excel hoàn tất.")

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

                st.download_button(
                    "⬇️ TẢI EXCEL SAU KHI DỊCH",
                    data=output_bytes,
                    file_name=out_name,
                    mime=(
                        "application/vnd.ms-excel.sheet.macroEnabled.12"
                        if original_ext == ".xlsm"
                        else "application/vnd.openxmlformats-officedocument."
                             "spreadsheetml.sheet"
                    ),
                    type="primary",
                    use_container_width=True,
                )

            except Exception as exc:
                st.error(f"Lỗi xử lý Excel: {exc}")
                st.exception(exc)

else:
    st.markdown(
        """
### 👇 Bắt đầu

Chọn **Trung → Việt** hoặc **Việt → Trung**, sau đó tải lên:

- **Excel**: chương trình giữ định dạng gốc và chèn bản dịch.
- **Ảnh bảng**: chương trình OCR nội dung và tạo Excel mới theo bố cục
  bảng trong ảnh.

> Với ảnh mẫu của bạn, chương trình ưu tiên kiểu trình bày:
> **tiếng gốc ở dòng trên + tiếng Việt ở dòng dưới**, header màu cam,
> đường viền bảng, căn giữa và giữ nguyên số liệu.
        """
    )
'''
(project / "app.py").write_text(app_code, encoding="utf-8")

requirements = """streamlit>=1.45
openpyxl>=3.1.5
Pillow>=10.4.0
google-genai>=1.40.0
pydantic>=2.8.0
"""
(project / "requirements.txt").write_text(requirements, encoding="utf-8")

readme = """# Trung ↔ Việt Excel & Image Translator

Ứng dụng Streamlit dịch song ngữ Trung ↔ Việt.

## Chức năng

- Upload `.xlsx` / `.xlsm`
- Upload `.png` / `.jpg` / `.jpeg` / `.webp`
- Trung → Việt hoặc Việt → Trung
- Chế độ song ngữ: nguyên văn + bản dịch trong cùng ô
- Excel: giữ merge, border, màu, font, độ rộng cột, chiều cao dòng, freeze panes và thuộc tính sheet chính
- Ảnh bảng: Gemini đọc bảng, OCR và dựng lại thành Excel
- Không cần cài Tesseract/PaddleOCR trên Streamlit Cloud

## Cài local

```bash
pip install -r requirements.txt
streamlit run app.py
