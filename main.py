from pathlib import Path
import textwrap, zipfile, os

project = Path("/mnt/data/trung_viet_excel_translator")
project.mkdir(exist_ok=True)

app_py = r'''import io
import re
from copy import copy

import cv2
import numpy as np
import openpyxl
import pandas as pd
import pytesseract
import streamlit as st
from PIL import Image
from deep_translator import GoogleTranslator
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


st.set_page_config(
    page_title="Trung ↔ Việt | Excel & Ảnh",
    page_icon="🌏",
    layout="wide",
)

st.title("🌏 Dịch Trung ↔ Việt — Excel & Ảnh")
st.caption("Upload Excel hoặc ảnh bảng biểu → OCR/đọc dữ liệu → dịch → xuất Excel.")


# =========================
# CẤU HÌNH
# =========================
LANG = {
    "Trung → Việt": ("zh-CN", "vi"),
    "Việt → Trung": ("vi", "zh-CN"),
}

# Từ khóa không nên dịch nếu là số, mã, công thức, email, URL...
URL_RE = re.compile(r"^(https?://|www\.)", re.I)
FORMULA_RE = re.compile(r"^=")


def looks_translatable(value: str) -> bool:
    """Chỉ dịch ô có chữ; giữ nguyên số, mã, URL, công thức."""
    if not isinstance(value, str):
        return False
    s = value.strip()
    if not s or URL_RE.match(s) or FORMULA_RE.match(s):
        return False
    # Có ít nhất một ký tự chữ Hán / Latin có chữ.
    return bool(re.search(r"[A-Za-zÀ-ỹĂăÂâÊêÔôƠơƯưĐđ一-龥]", s))


@st.cache_resource(show_spinner=False)
def get_translator(source: str, target: str):
    return GoogleTranslator(source=source, target=target)


def translate_texts(texts, source: str, target: str):
    """Dịch theo batch, có fallback từng câu."""
    if not texts:
        return {}

    translator = get_translator(source, target)
    result = {}

    # Google Translator thường ổn với batch vừa phải.
    batch_size = 40
    for start in range(0, len(texts), batch_size):
        batch = texts[start:start + batch_size]
        try:
            translated = translator.translate_batch(batch)
            if isinstance(translated, str):
                translated = [translated]
            for src, dst in zip(batch, translated):
                result[src] = dst if dst else src
        except Exception:
            for src in batch:
                try:
                    result[src] = translator.translate(src)
                except Exception:
                    result[src] = src
    return result


def translate_workbook(uploaded_bytes: bytes, source: str, target: str):
    """Dịch nội dung workbook nhưng giữ style, merge, độ rộng cột, chiều cao dòng..."""
    wb = openpyxl.load_workbook(
        io.BytesIO(uploaded_bytes),
        data_only=False,
        keep_links=True,
    )

    texts = []
    seen = set()

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and looks_translatable(value):
                    if value not in seen:
                        texts.append(value)
                        seen.add(value)

    translations = translate_texts(texts, source, target)

    changed = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value in translations:
                    new_value = translations[value]
                    if new_value != value:
                        cell.value = new_value
                        changed += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), changed, len(texts)


# =========================
# OCR ẢNH
# =========================
def preprocess_for_ocr(image: Image.Image):
    img = np.array(image.convert("RGB"))
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)

    # Phóng to giúp Tesseract đọc bảng nhỏ tốt hơn.
    scale = 2.0 if max(gray.shape) < 2200 else 1.5
    gray = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)

    # Khử nhiễu nhẹ + threshold.
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    bw = cv2.adaptiveThreshold(
        gray, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        31, 11
    )
    return bw


def ocr_words(image: Image.Image):
    """OCR có bounding box để gom chữ thành các hàng/cột."""
    bw = preprocess_for_ocr(image)

    # psm 6 phù hợp bảng đơn giản.
    config = "--oem 3 --psm 6"
    data = pytesseract.image_to_data(
        bw,
        lang="chi_sim+vie+eng",
        config=config,
        output_type=pytesseract.Output.DICT,
    )

    words = []
    n = len(data["text"])
    for i in range(n):
        text = (data["text"][i] or "").strip()
        conf_raw = data["conf"][i]
        try:
            conf = float(conf_raw)
        except Exception:
            conf = -1

        if not text or conf < 15:
            continue

        x = int(data["left"][i])
        y = int(data["top"][i])
        w = int(data["width"][i])
        h = int(data["height"][i])
        words.append({
            "text": text,
            "x": x, "y": y, "w": w, "h": h,
            "cx": x + w / 2,
            "cy": y + h / 2,
            "conf": conf,
        })
    return words, bw


def group_words_to_rows(words):
    if not words:
        return []

    words = sorted(words, key=lambda z: (z["cy"], z["x"]))
    rows = []

    for word in words:
        placed = False
        for row in rows:
            avg_y = sum(w["cy"] for w in row) / len(row)
            avg_h = sum(w["h"] for w in row) / len(row)
            if abs(word["cy"] - avg_y) <= max(12, avg_h * 0.65):
                row.append(word)
                placed = True
                break
        if not placed:
            rows.append([word])

    rows.sort(key=lambda r: min(w["y"] for w in r))
    for row in rows:
        row.sort(key=lambda w: w["x"])
    return rows


def rows_to_matrix(rows):
    """
    Gom OCR thành ma trận gần giống bảng:
    - Các từ trên cùng hàng được ghép theo khoảng cách.
    - Số cột lấy theo vị trí x tương đối.
    """
    if not rows:
        return []

    # Ước lượng các tâm cột bằng clustering 1D đơn giản.
    all_x = sorted([w["cx"] for r in rows for w in r])
    if not all_x:
        return []

    # Khoảng cách lớn giữa các từ -> khả năng là cột khác.
    gaps = [all_x[i+1] - all_x[i] for i in range(len(all_x)-1)]
    median_gap = np.median(gaps) if gaps else 30
    threshold = max(45, median_gap * 2.4)

    centers = []
    for x in all_x:
        if not centers or abs(x - centers[-1]) > threshold:
            centers.append(x)
        else:
            centers[-1] = (centers[-1] + x) / 2

    # Nếu quá nhiều cột giả do câu dài, giới hạn ở 12 cột.
    if len(centers) > 12:
        centers = centers[:12]

    matrix = []
    for row in rows:
        cells = [""] * len(centers)
        for word in row:
            idx = min(range(len(centers)), key=lambda j: abs(word["cx"] - centers[j]))
            if cells[idx]:
                cells[idx] += " " + word["text"]
            else:
                cells[idx] = word["text"]
        # Loại các cột rỗng ở cuối.
        while cells and not cells[-1]:
            cells.pop()
        matrix.append(cells)

    max_cols = max((len(r) for r in matrix), default=0)
    matrix = [r + [""] * (max_cols - len(r)) for r in matrix]
    return matrix


def make_image_workbook(image: Image.Image, source: str, target: str):
    words, processed = ocr_words(image)
    rows = group_words_to_rows(words)
    matrix = rows_to_matrix(rows)

    if not matrix:
        raise ValueError("Không nhận diện được chữ trong ảnh.")

    # Lấy các text cần dịch.
    texts = []
    seen = set()
    for row in matrix:
        for value in row:
            if looks_translatable(value) and value not in seen:
                texts.append(value)
                seen.add(value)

    translations = translate_texts(texts, source, target)

    wb = Workbook()
    ws = wb.active
    ws.title = "Dịch từ ảnh"

    # Style gần với bảng mẫu: header cam, viền đen, căn giữa.
    orange = "E87800"
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Tách dòng tiêu đề nếu OCR nhận diện được một dòng đầu dài.
    start_row = 1
    if matrix and len(matrix[0]) <= 2 and any(matrix[0]):
        title = " ".join([x for x in matrix[0] if x])
        title = translations.get(title, title)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(1, max(len(r) for r in matrix)))
        c = ws.cell(1, 1, title)
        c.font = Font(name="Arial", size=14, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        start_row = 2

    for r_idx, row in enumerate(matrix[start_row - 1:], start=start_row):
        for c_idx, value in enumerate(row, start=1):
            new_value = translations.get(value, value)
            cell = ws.cell(r_idx, c_idx, new_value)
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            if r_idx == start_row:
                cell.font = Font(name="Arial", size=10, bold=True)
                cell.fill = PatternFill("solid", fgColor=orange)
            else:
                cell.font = Font(name="Arial", size=10)

    # Tự động độ rộng cột.
    max_cols = max((len(r) for r in matrix), default=1)
    for c in range(1, max_cols + 1):
        values = [str(ws.cell(r, c).value or "") for r in range(1, ws.max_row + 1)]
        width = min(35, max(8, max(len(v) for v in values) * 1.15 + 2))
        ws.column_dimensions[get_column_letter(c)].width = width

    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = 30 if r > 1 else 40

    ws.freeze_panes = "A2"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), matrix, translations, processed


# =========================
# UI
# =========================
with st.sidebar:
    st.header("⚙️ Cấu hình")
    direction = st.radio("Hướng dịch", list(LANG.keys()))
    source, target = LANG[direction]

    st.info(
        "Dịch dùng Google Translate qua thư viện deep-translator. "
        "Không cần API key."
    )

tab_excel, tab_image = st.tabs(["📊 Excel → Excel", "🖼️ Ảnh → Excel"])

with tab_excel:
    st.subheader("1. Dịch file Excel và giữ định dạng")
    excel_file = st.file_uploader(
        "Chọn file .xlsx / .xlsm",
        type=["xlsx", "xlsm"],
        key="excel",
    )

    if excel_file:
        st.success(f"Đã chọn: {excel_file.name}")

        if st.button("🚀 Dịch Excel", type="primary", key="translate_excel"):
            try:
                with st.spinner("Đang đọc Excel và dịch..."):
                    output, changed, total = translate_workbook(
                        excel_file.getvalue(), source, target
                    )

                st.success(f"Hoàn tất: dịch {changed}/{total} chuỗi.")
                base = re.sub(r"\.(xlsx|xlsm)$", "", excel_file.name, flags=re.I)
                ext = ".xlsx"  # xuất chuẩn xlsx
                out_name = f"{base}_{'VI' if target == 'vi' else 'ZH'}.xlsx"

                st.download_button(
                    "⬇️ Tải Excel đã dịch",
                    data=output,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error("Không thể dịch file Excel.")
                st.exception(e)

with tab_image:
    st.subheader("2. Ảnh bảng biểu → OCR → dịch → Excel")
    image_file = st.file_uploader(
        "Chọn file .png / .jpg / .jpeg / .webp",
        type=["png", "jpg", "jpeg", "webp"],
        key="image",
    )

    if image_file:
        image = Image.open(image_file)
        st.image(image, caption="Ảnh gốc", use_container_width=True)

        if st.button("🔎 OCR + dịch + tạo Excel", type="primary", key="translate_image"):
            try:
                with st.spinner("Đang OCR ảnh, nhận diện bảng và dịch..."):
                    output, matrix, translations, processed = make_image_workbook(
                        image, source, target
                    )

                st.success("Đã tạo Excel từ ảnh.")

                # Preview dữ liệu OCR sau dịch
                preview = []
                for row in matrix:
                    preview.append([translations.get(v, v) for v in row])
                max_cols = max((len(r) for r in preview), default=1)
                preview = [r + [""] * (max_cols - len(r)) for r in preview]
                st.dataframe(pd.DataFrame(preview), use_container_width=True, hide_index=True)

                base = re.sub(r"\.[^.]+$", "", image_file.name)
                out_name = f"{base}_translated.xlsx"
                st.download_button(
                    "⬇️ Tải Excel đã dịch",
                    data=output,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(
                    "OCR ảnh thất bại. Hãy dùng ảnh rõ, thẳng, độ phân giải cao "
                    "và chữ Trung/Việt không bị mờ."
                )
                st.exception(e)

st.divider()
st.caption(
    "Lưu ý: Excel đầu vào được giữ nguyên merge cell, font, màu, border, "
    "độ rộng cột, chiều cao dòng và sheet ở mức openpyxl hỗ trợ. "
    "Ảnh không chứa cấu trúc Excel thật nên Excel tạo từ ảnh là bản tái dựng."
)
'''

requirements = r'''streamlit>=1.38,<2
openpyxl>=3.1.5
pandas>=2.2
numpy>=1.26
opencv-python-headless>=4.10
Pillow>=10.4
pytesseract>=0.3.13
deep-translator>=1.11.4
'''

packages = r'''tesseract-ocr
tesseract-ocr-chi-sim
tesseract-ocr-vie
tesseract-ocr-eng
'''

readme = """# Trung ↔ Việt Excel & Image Translator 
Ứng dụng Streamlit dịch hai chiều:
- Trung → Việt
- Việt → Trung
- Excel .xlsx / .xlsm → Excel đã dịch
- Ảnh .png / .jpg / .jpeg / .webp → OCR → dịch → Excel

Excel đầu vào được giữ lại các thành phần định dạng chính:
- Sheet
- Merge cell
- Font
- Màu nền
- Border
- Alignment
- Độ rộng cột
- Chiều cao dòng
- Freeze panes

## Chạy local 

```bash
pip install -r requirements.txt
streamlit run app.py
