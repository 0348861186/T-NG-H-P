code = r'''import io
import os
import re
import copy
import hashlib
from pathlib import Path

import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import numpy as np
import openpyxl
from openpyxl.cell.cell import MergedCell
from deep_translator import GoogleTranslator
import easyocr


# ============================================================
# CẤU HÌNH
# ============================================================
st.set_page_config(
    page_title="Dịch Song Ngữ Trung - Việt",
    page_icon="🌐",
    layout="wide",
)

SUPPORTED_IMAGE_TYPES = ["png", "jpg", "jpeg", "webp", "bmp", "tif", "tiff"]
SUPPORTED_EXCEL_TYPES = ["xlsx", "xlsm"]


# ============================================================
# CACHE OCR
# ============================================================
@st.cache_resource
def load_ocr_reader():
    # EasyOCR không cần "vi" để nhận diện mọi trường hợp; giữ ch/en để
    # nhận diện Trung + Latin tốt hơn. Nếu môi trường EasyOCR hỗ trợ vi,
    # có thể thêm "vi".
    return easyocr.Reader(["ch_sim", "en"], gpu=False)


# ============================================================
# DỊCH
# ============================================================
@st.cache_data(show_spinner=False)
def cached_translate(text: str, direction: str) -> str:
    text = str(text).strip()
    if not text:
        return ""

    # Không dịch các ô chỉ chứa số / ký hiệu.
    if re.fullmatch(r"[\d\s.,%+\-*/=()（）【】\[\]{}:;，。！？!?_/\\]+", text):
        return text

    source, target = (
        ("zh-CN", "vi") if direction == "Trung - Việt"
        else ("vi", "zh-CN")
    )

    try:
        return GoogleTranslator(source=source, target=target).translate(text) or ""
    except Exception as exc:
        # Không phá file vì lỗi dịch; giữ nội dung gốc và ghi cảnh báo.
        return f"[LỖI DỊCH: {exc}]"


def translate_text(text, direction):
    if text is None:
        return None
    return cached_translate(str(text), direction)


# ============================================================
# HÀM HỖ TRỢ EXCEL
# ============================================================
def is_formula(value):
    return isinstance(value, str) and value.startswith("=")


def contains_translatable_text(value):
    if value is None:
        return False
    text = str(value).strip()
    if not text:
        return False
    if is_formula(text):
        return False
    # Không dịch ô chỉ có số/ký hiệu.
    return not bool(re.fullmatch(r"[\d\s.,%+\-*/=()（）【】\[\]{}:;，。！？!?_/\\]+", text))


def make_bilingual(original, translated, direction):
    """
    Bất kể chiều dịch:
        Trung
        Việt

    Nếu Trung -> Việt: original là Trung.
    Nếu Việt -> Trung: translated là Trung.
    """
    if direction == "Trung - Việt":
        chinese = str(original)
        vietnamese = str(translated)
    else:
        chinese = str(translated)
        vietnamese = str(original)

    return f"{chinese}\n{vietnamese}"


def get_line_count(text):
    if text is None:
        return 1
    return max(1, str(text).count("\n") + 1)


def adjust_row_height(ws, row_idx, original_height, extra_lines=1):
    """
    Tăng chiều cao dòng một cách an toàn khi cell chuyển thành 2 dòng.
    Không thay đổi nếu người dùng đang dùng chiều cao tự động/không đặt.
    """
    if original_height is None:
        # Excel thường tự điều chỉnh khi mở file nếu wrap_text=True.
        return

    base = float(original_height)
    # Với 2 dòng, khoảng 1.9 lần thường đủ để không cắt chữ.
    factor = max(1.0, 1.0 + 0.9 * extra_lines)
    new_height = max(base, base * factor)

    # Giới hạn để tránh row khổng lồ.
    ws.row_dimensions[row_idx].height = min(new_height, 409.0)


def load_excel_preserving_features(file_obj, extension):
    """
    Mở trực tiếp workbook gốc thay vì tạo workbook mới.
    Điều này giúp giữ tối đa các thành phần Excel mà openpyxl hỗ trợ.
    """
    file_obj.seek(0)

    keep_vba = extension.lower() == "xlsm"

    return openpyxl.load_workbook(
        file_obj,
        read_only=False,
        data_only=False,
        keep_vba=keep_vba,
    )


def process_excel(file_obj, direction):
    extension = Path(file_obj.name).suffix.lower().lstrip(".")
    if extension not in SUPPORTED_EXCEL_TYPES:
        raise ValueError(
            "Phiên bản này hỗ trợ .xlsx và .xlsm. "
            "File .xls cũ cần chuyển sang .xlsx trước khi dịch."
        )

    wb = load_excel_preserving_features(file_obj, extension)

    translated_count = 0
    skipped_count = 0
    warnings = []

    for ws in wb.worksheets:
        # Lưu chiều cao dòng ban đầu.
        original_row_heights = {
            idx: ws.row_dimensions[idx].height
            for idx in range(1, ws.max_row + 1)
        }

        # Xử lý snapshot ô trước khi sửa, tránh ảnh hưởng khi merged cell.
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue

                value = cell.value

                if not contains_translatable_text(value):
                    skipped_count += 1
                    continue

                original = str(value)

                try:
                    translated = translate_text(original, direction)
                    if not translated:
                        skipped_count += 1
                        continue

                    # Nếu translator trả lỗi, không thay nội dung để tránh
                    # làm hỏng file.
                    if str(translated).startswith("[LỖI DỊCH:"):
                        warnings.append(
                            f"{ws.title}!{cell.coordinate}: {translated}"
                        )
                        continue

                    # Nếu đã là song ngữ do chạy lại app, không dịch lại.
                    if "\n" in original:
                        lines = original.splitlines()
                        if len(lines) >= 2:
                            # Trong cả hai chiều, dòng 1 phải là Trung.
                            # Không cố đoán lại nếu file đã song ngữ.
                            skipped_count += 1
                            continue

                    cell.value = make_bilingual(original, translated, direction)

                    # Chỉ chỉnh wrap_text, giữ lại toàn bộ thuộc tính alignment còn lại.
                    old_alignment = copy.copy(cell.alignment)
                    old_alignment.wrap_text = True
                    cell.alignment = old_alignment

                    # Không thay font/fill/border/number_format:
                    # cell vẫn là cell gốc của workbook.
                    row_height = original_row_heights.get(cell.row)

                    if row_height is not None:
                        adjust_row_height(
                            ws,
                            cell.row,
                            row_height,
                            extra_lines=1,
                        )

                    translated_count += 1

                except Exception as exc:
                    warnings.append(
                        f"{ws.title}!{cell.coordinate}: {type(exc).__name__}: {exc}"
                    )

    output = io.BytesIO()
    output_format = "xlsm" if extension == "xlsm" else "xlsx"

    # Giữ VBA khi file nguồn là XLSM.
    wb.save(output)
    output.seek(0)

    return output, translated_count, skipped_count, warnings, output_format


# ============================================================
# HỖ TRỢ FONT CHO HÌNH ẢNH
# ============================================================
def find_font():
    candidates = [
        # Windows
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\msyh.ttf",
        r"C:\Windows\Fonts\simhei.ttf",
        r"C:\Windows\Fonts\simsun.ttc",
        r"C:\Windows\Fonts\arialuni.ttf",
        r"C:\Windows\Fonts\arial.ttf",

        # Linux
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.otf",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.otf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",

        # macOS
        "/System/Library/Fonts/PingFang.ttc",
        "/System/Library/Fonts/Hiragino Sans GB.ttc",
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    ]

    for path in candidates:
        if os.path.exists(path):
            return path

    return None


def load_font(font_path, size):
    if font_path:
        try:
            return ImageFont.truetype(font_path, max(8, int(size)))
        except Exception:
            pass
    return ImageFont.load_default()


def bbox_to_rect(bbox):
    xs = [p[0] for p in bbox]
    ys = [p[1] for p in bbox]
    return (
        int(min(xs)),
        int(min(ys)),
        int(max(xs)),
        int(max(ys)),
    )


def group_ocr_lines(results, y_tolerance_ratio=0.55):
    """
    Gom các box OCR nằm trên cùng một dòng.
    Trả về danh sách dòng, mỗi dòng gồm:
        text, x1, y1, x2, y2
    """
    items = []

    for bbox, text, prob in results:
        if prob < 0.20 or not str(text).strip():
            continue

        x1, y1, x2, y2 = bbox_to_rect(bbox)
        h = max(1, y2 - y1)
        center_y = (y1 + y2) / 2

        items.append({
            "text": str(text).strip(),
            "x1": x1,
            "y1": y1,
            "x2": x2,
            "y2": y2,
            "h": h,
            "cy": center_y,
        })

    items.sort(key=lambda x: (x["cy"], x["x1"]))

    lines = []

    for item in items:
        placed = False

        for line in lines:
            avg_cy = sum(i["cy"] for i in line) / len(line)
            avg_h = sum(i["h"] for i in line) / len(line)

            if abs(item["cy"] - avg_cy) <= max(4, avg_h * y_tolerance_ratio):
                line.append(item)
                placed = True
                break

        if not placed:
            lines.append([item])

    output = []

    for line in lines:
        line.sort(key=lambda x: x["x1"])

        text = " ".join(i["text"] for i in line)

        output.append({
            "text": text,
            "x1": min(i["x1"] for i in line),
            "y1": min(i["y1"] for i in line),
            "x2": max(i["x2"] for i in line),
            "y2": max(i["y2"] for i in line),
            "height": max(i["h"] for i in line),
        })

    output.sort(key=lambda x: (x["y1"], x["x1"]))
    return output


def text_size(draw, text, font):
    try:
        box = draw.textbbox((0, 0), text, font=font)
        return box[2] - box[0], box[3] - box[1]
    except Exception:
        return draw.textsize(text, font=font)


def fit_font_to_width(draw, text, font_path, initial_size, max_width):
    size = max(8, int(initial_size))

    while size > 8:
        font = load_font(font_path, size)
        width, _ = text_size(draw, text, font)
        if width <= max_width:
            return font
        size -= 1

    return load_font(font_path, 8)


def choose_background_color(image, box, sample_padding=3):
    """
    Lấy màu trung bình quanh vùng bên dưới OCR để tạo nền cho bản dịch.
    Không sửa vùng chữ gốc.
    """
    x1, y1, x2, y2 = box

    crop = image.crop((
        max(0, x1 - sample_padding),
        max(0, y1 - sample_padding),
        min(image.width, x2 + sample_padding),
        min(image.height, y2 + sample_padding),
    ))

    arr = np.asarray(crop).astype(np.float32)

    if arr.size == 0:
        return (255, 255, 255)

    mean = arr.reshape(-1, arr.shape[-1]).mean(axis=0)

    return tuple(int(max(0, min(255, x))) for x in mean[:3])


def contrast_text_color(bg):
    r, g, b = bg
    luminance = 0.299 * r + 0.587 * g + 0.114 * b
    return (20, 20, 20) if luminance > 150 else (245, 245, 245)


def draw_translation_below(
    base_image,
    line,
    translated,
    font_path,
    direction,
):
    """
    Chỉ vẽ phần bổ sung phía dưới dòng OCR.
    Không vẽ lại chữ gốc.

    Nếu phía dưới không đủ chỗ:
    - mở rộng canvas theo chiều dọc.
    """
    draw = ImageDraw.Draw(base_image)

    x1, y1, x2, y2 = (
        line["x1"],
        line["y1"],
        line["x2"],
        line["y2"],
    )

    original_h = max(10, y2 - y1)

    # Font dịch gần tương đương font gốc nhưng nhỏ hơn một chút.
    font_size = max(12, int(original_h * 0.72))

    # Không để bản dịch dài quá vùng OCR quá nhiều.
    max_width = max(80, min(base_image.width - x1 - 5, max(120, x2 - x1) * 1.35))

    font = fit_font_to_width(
        draw,
        translated,
        font_path,
        font_size,
        max_width,
    )

    _, text_h = text_size(draw, translated, font)

    gap = max(3, int(original_h * 0.18))
    top = y2 + gap
    bottom = top + text_h + gap

    # Mở rộng ảnh nếu cần để bản dịch không bị cắt.
    if bottom > base_image.height:
        extra = bottom - base_image.height + gap
        expanded = Image.new(
            "RGB",
            (base_image.width, base_image.height + extra),
            (255, 255, 255),
        )
        expanded.paste(base_image, (0, 0))
        base_image = expanded
        draw = ImageDraw.Draw(base_image)

    # Nền nhỏ ngay phía sau bản dịch để dễ đọc, nhưng không che chữ gốc.
    bg = choose_background_color(base_image, (x1, y1, x2, y2))
    text_color = contrast_text_color(bg)

    pad_x = 3
    pad_y = 1

    # Chỉ tạo nền trong vùng bản dịch.
    bbox = draw.textbbox(
        (x1, top),
        translated,
        font=font,
    )

    rect = (
        max(0, bbox[0] - pad_x),
        max(0, bbox[1] - pad_y),
        min(base_image.width, bbox[2] + pad_x),
        min(base_image.height, bbox[3] + pad_y),
    )

    draw.rectangle(rect, fill=bg)

    draw.text(
        (x1, top),
        translated,
        font=font,
        fill=text_color,
    )

    return base_image


# ============================================================
# XỬ LÝ HÌNH ẢNH
# ============================================================
def process_image(image_file, direction):
    reader = load_ocr_reader()

    image_file.seek(0)
    original = Image.open(image_file)

    # Giữ RGB/RGBA thành RGB để output ổn định.
    image = original.convert("RGB")

    # OCR trên ảnh gốc, không resize để giữ tọa độ thực.
    img_np = np.array(image)
    results = reader.readtext(img_np)

    lines = group_ocr_lines(results)

    output = image.copy()
    font_path = find_font()

    if not lines:
        return output, 0, [
            "Không nhận diện được chữ trong ảnh. Hãy thử ảnh có độ phân giải cao hơn."
        ]

    # Xử lý từ trên xuống dưới.
    translated_count = 0
    warnings = []

    for line in lines:
        original_text = line["text"]

        if not contains_translatable_text(original_text):
            continue

        try:
            translated = translate_text(original_text, direction)

            if not translated:
                continue

            if str(translated).startswith("[LỖI DỊCH:"):
                warnings.append(
                    f'OCR "{original_text}": {translated}'
                )
                continue

            # Đảm bảo tiếng Việt nằm dưới tiếng Trung:
            # trên ảnh gốc đã có dòng gốc. Chỉ vẽ phần còn thiếu.
            vietnamese = (
                translated if direction == "Trung - Việt"
                else original_text
            )
            chinese = (
                original_text if direction == "Trung - Việt"
                else translated
            )

            # Dòng gốc trong ảnh phải là tiếng Trung ở phía trên.
            # Với Việt -> Trung, bản gốc Việt đang ở trên ảnh.
            # Để đáp ứng yêu cầu, ta phải che dòng Việt cũ rồi vẽ Trung
            # ở trên và Việt ở dưới. Cách an toàn nhất là redraw vùng OCR.
            if direction == "Việt - Trung":
                x1, y1, x2, y2 = (
                    line["x1"], line["y1"], line["x2"], line["y2"]
                )

                pad = max(4, int(line["height"] * 0.25))

                # Lấy màu nền từ vùng quanh box.
                bg = choose_background_color(
                    output,
                    (x1, y1, x2, y2),
                    sample_padding=pad,
                )

                draw = ImageDraw.Draw(output)
                draw.rectangle(
                    (
                        max(0, x1 - pad),
                        max(0, y1 - pad),
                        min(output.width, x2 + pad),
                        min(output.height, y2 + pad),
                    ),
                    fill=bg,
                )

                # Vẽ tiếng Trung tại vị trí chữ gốc.
                font = fit_font_to_width(
                    draw,
                    chinese,
                    font_path,
                    max(12, int(line["height"] * 0.82)),
                    max(80, x2 - x1),
                )

                draw.text(
                    (x1, y1),
                    chinese,
                    font=font,
                    fill=contrast_text_color(bg),
                )

                # Sau đó vẽ tiếng Việt bên dưới.
                output = draw_translation_below(
                    output,
                    {
                        **line,
                        "y2": y1 + line["height"],
                    },
                    vietnamese,
                    font_path,
                    direction,
                )

            else:
                output = draw_translation_below(
                    output,
                    line,
                    vietnamese,
                    font_path,
                    direction,
                )

            translated_count += 1

        except Exception as exc:
            warnings.append(
                f'OCR "{original_text}": {type(exc).__name__}: {exc}'
            )

    return output, translated_count, warnings


# ============================================================
# GIAO DIỆN STREAMLIT
# ============================================================
st.title("🌐 Ứng dụng dịch song ngữ Trung ↔ Việt")
st.caption(
    "Giữ nội dung gốc và đặt tiếng Việt ngay bên dưới tiếng Trung."
)

with st.sidebar:
    st.header("⚙️ Cấu hình")

    direction = st.selectbox(
        "Chiều dịch",
        ["Trung - Việt", "Việt - Trung"],
        index=0,
    )

    st.divider()

    st.markdown("### Quy tắc kết quả")
    st.markdown(
        """
        **Luôn hiển thị:**

        🇨🇳 Tiếng Trung  
        🇻🇳 Tiếng Việt

        Với Excel, bản dịch nằm trong **cùng ô** và xuống dòng.
        """
    )

uploaded_file = st.file_uploader(
    "📤 Tải file lên",
    type=SUPPORTED_EXCEL_TYPES + SUPPORTED_IMAGE_TYPES,
    help="Excel: .xlsx/.xlsm | Hình ảnh: .png/.jpg/.jpeg/.webp/.bmp/.tif/.tiff",
)

if uploaded_file is None:
    st.info("Hãy tải file Excel hoặc hình ảnh để bắt đầu.")
    st.stop()

extension = Path(uploaded_file.name).suffix.lower().lstrip(".")

st.write(f"**File:** `{uploaded_file.name}`")

# ============================================================
# EXCEL
# ============================================================
if extension in SUPPORTED_EXCEL_TYPES:
    st.info(
        "📊 Excel: app chỉnh trực tiếp workbook gốc trong bộ nhớ, "
        "thay vì tạo workbook mới, nhằm bảo toàn tối đa cấu trúc/định dạng."
    )

    col1, col2 = st.columns([1, 3])

    with col1:
        start_excel = st.button(
            "🚀 Bắt đầu dịch Excel",
            type="primary",
            use_container_width=True,
        )

    if start_excel:
        with st.spinner(
            "Đang dịch Excel. Với file lớn có thể mất thời gian..."
        ):
            try:
                (
                    output_excel,
                    translated_count,
                    skipped_count,
                    warnings,
                    output_format,
                ) = process_excel(
                    uploaded_file,
                    direction,
                )

                st.success(
                    f"✅ Hoàn tất: đã dịch {translated_count} ô."
                )

                if skipped_count:
                    st.caption(
                        f"Đã bỏ qua {skipped_count} ô không cần dịch "
                        f"(ô trống, số/ký hiệu, công thức hoặc nội dung đã song ngữ)."
                    )

                if warnings:
                    with st.expander(
                        f"⚠️ Có {len(warnings)} cảnh báo"
                    ):
                        for warning in warnings[:100]:
                            st.warning(warning)

                original_stem = Path(uploaded_file.name).stem
                output_name = (
                    f"{original_stem}_song_ngu.{output_format}"
                )

                mime = (
                    "application/vnd.ms-excel.sheet.macroEnabled.12"
                    if output_format == "xlsm"
                    else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.download_button(
                    "📥 Download Excel đã dịch",
                    data=output_excel.getvalue(),
                    file_name=output_name,
                    mime=mime,
                    type="primary",
                    use_container_width=True,
                )

            except Exception as exc:
                st.error(
                    "❌ Không thể xử lý Excel."
                )
                st.exception(exc)


# ============================================================
# IMAGE
# ============================================================
elif extension in SUPPORTED_IMAGE_TYPES:
    st.info(
        "🖼️ Hình ảnh: OCR từng dòng và đặt bản dịch ngay phía dưới. "
        "Ảnh có thể được mở rộng chiều cao nếu phía dưới không đủ chỗ."
    )

    start_image = st.button(
        "🚀 Bắt đầu dịch hình ảnh",
        type="primary",
        use_container_width=True,
    )

    if start_image:
        with st.spinner(
            "Đang OCR và dịch hình ảnh..."
        ):
            try:
                result_img, translated_count, warnings = process_image(
                    uploaded_file,
                    direction,
                )

                st.success(
                    f"✅ Hoàn tất: đã xử lý {translated_count} dòng chữ."
                )

                if warnings:
                    with st.expander(
                        f"⚠️ Có {len(warnings)} cảnh báo"
                    ):
                        for warning in warnings[:100]:
                            st.warning(warning)

                st.image(
                    result_img,
                    caption="Kết quả: tiếng Trung ở trên, tiếng Việt ở dưới",
                    use_container_width=True,
                )

                buf = io.BytesIO()
                result_img.save(
                    buf,
                    format="PNG",
                    optimize=False,
                )
                buf.seek(0)

                output_name = (
                    f"{Path(uploaded_file.name).stem}_song_ngu.png"
                )

                st.download_button(
                    "📥 Download hình ảnh đã dịch",
                    data=buf.getvalue(),
                    file_name=output_name,
                    mime="image/png",
                    type="primary",
                    use_container_width=True,
                )

            except Exception as exc:
                st.error(
                    "❌ Không thể xử lý hình ảnh."
                )
                st.exception(exc)
else:
    st.error(
        "Định dạng file chưa được hỗ trợ."
    )
'''
path = "/mnt/data/app_dich_song_ngu.py"
with open(path, "w", encoding="utf-8") as f:
    f.write(code)

req = """streamlit
openpyxl
pillow
numpy
pandas
deep-translator
easyocr
torch
torchvision
"""
with open("/mnt/data/requirements.txt", "w", encoding="utf-8") as f:
    f.write(req)

print(path)
print("/mnt/data/requirements.txt")
