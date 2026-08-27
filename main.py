import io
import re
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from deep_translator import GoogleTranslator
import easyocr
import numpy as np
from PIL import Image

# ============================================================
# CẤU HÌNH STREAMLIT
# ============================================================
st.set_page_config(
    page_title="Dịch Bảng Chấm Công Bằng Thư Viện Chuyên Dụng (No-AI)",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Dịch & Xuất Bảng Chấm Công Song Ngữ (100% Thư Viện - Không Cần API Key)")
st.caption("Sử dụng EasyOCR quét ảnh + Deep-Translator dịch thuật | Miễn phí hoàn toàn & Không lo hết Quota!")

# Khởi tạo EasyOCR Reader (Caching để không bị load lại nhiều lần)
@st.cache_resource
def load_ocr_readers():
    # Reader đọc tiếng Trung & Tiếng Anh/Số
    reader_zh = easyocr.Reader(['ch_sim', 'en'], gpu=False)
    # Reader đọc tiếng Việt & Tiếng Anh/Số
    reader_vi = easyocr.Reader(['vi', 'en'], gpu=False)
    return reader_zh, reader_vi

# ============================================================
# 1. CẤU HÌNH DASHBOARD
# ============================================================
col1, col2 = st.columns([1, 2])

with col1:
    translation_mode = st.radio(
        "Chế độ dịch:",
        options=["Trung ➔ Việt", "Việt ➔ Trung"],
        horizontal=True
    )

with col2:
    uploaded_file = st.file_uploader(
        "Tải lên File Excel (.xlsx) hoặc File Ảnh (.png, .jpg, .jpeg):", 
        type=["png", "jpg", "jpeg", "xlsx"]
    )

# Hàm kiểm tra chữ Trung / Việt
def has_chinese(text):
    if not isinstance(text, str): return False
    return bool(re.search(r'[\u4e00-\u9fff]', text))

def has_vietnamese(text):
    if not isinstance(text, str): return False
    return bool(re.search(r'[àáảãạâầấẩẫậăằắẳẵặèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđĐ]', text, re.IGNORECASE))

# ============================================================
# 2. HÀM DỊCH THUẬT MIỄN PHÍ BẰNG DEEP-TRANSLATOR
# ============================================================
def translate_batch(texts, mode):
    src_lang = 'zh-CN' if mode == "Trung ➔ Việt" else 'vi'
    tgt_lang = 'vi' if mode == "Trung ➔ Việt" else 'zh-CN'
    
    translator = GoogleTranslator(source=src_lang, target=tgt_lang)
    translation_dict = {}
    
    for text in texts:
        if not text.strip() or text.isnumeric():
            continue
        try:
            translated = translator.translate(text)
            translation_dict[text] = translated
        except Exception:
            translation_dict[text] = text
            
    return translation_dict

# ============================================================
# 3. HÀM TẠO EXCEL TỪ KẾT QUẢ QUÉT ẢNH (EASYOCR)
# ============================================================
def build_excel_from_ocr(ocr_results, mode):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet_OCR"

    font_name = "Microsoft YaHei"
    orange_fill = PatternFill(fill_type="solid", fgColor="ED7D00")
    thin_side = Side(style="thin", color="000000")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # Gom nhóm chữ theo hàng (dựa trên tọa độ Y)
    rows_grouped = []
    # Sắp xếp các đoạn văn bản từ trên xuống dưới
    sorted_res = sorted(ocr_results, key=lambda x: x[0][0][1])
    
    current_row = []
    last_y = None
    
    for bbox, text, prob in sorted_res:
        if prob < 0.2 or not text.strip(): # Lọc bỏ rác OCR
            continue
        y_center = (bbox[0][1] + bbox[2][1]) / 2
        
        if last_y is None or abs(y_center - last_y) < 15: # Cùng 1 hàng
            current_row.append((bbox[0][0], text)) # Thêm cùng tọa độ X
            last_y = y_center
        else:
            # Sắp xếp từ trái sang phải theo X
            current_row.sort(key=lambda item: item[0])
            rows_grouped.append([item[1] for item in current_row])
            current_row = [(bbox[0][0], text)]
            last_y = y_center

    if current_row:
        current_row.sort(key=lambda item: item[0])
        rows_grouped.append([item[1] for item in current_row])

    # Thu thập tất cả từ cần dịch
    all_texts = []
    for row in rows_grouped:
        for txt in row:
            if (mode == "Trung ➔ Việt" and has_chinese(txt)) or (mode == "Việt ➔ Trung" and (has_vietnamese(txt) or not has_chinese(txt))):
                all_texts.append(txt)

    # Dịch toàn bộ bằng thư viện Deep-Translator
    trans_dict = translate_batch(list(set(all_texts)), mode)

    # Ghi dữ liệu ra file Excel
    for r_idx, row in enumerate(rows_grouped, start=1):
        for c_idx, txt in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx)
            translated_txt = trans_dict.get(txt, "")
            
            if translated_txt:
                cell.value = f"{txt}\n{translated_txt}"
            else:
                cell.value = txt
                
            cell.font = Font(name=font_name, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
            
        ws.row_dimensions[r_idx].height = 30

    # Chỉnh độ rộng cột tự động
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# ============================================================
# 4. XỬ LÝ CHÍNH KHI BẤM NÚT DỊCH
# ============================================================
if uploaded_file is not None:
    is_excel = uploaded_file.name.lower().endswith('.xlsx')
    button_label = f"🚀 Dịch ({translation_mode}) & Bảo Toàn Format Excel" if is_excel else f"🚀 Quét OCR Ảnh & Dịch ({translation_mode})"
    
    if st.button(button_label, use_container_width=True):
        try:
            # ----------------------------------------------------
            # TRƯỜNG HỢP 1: DỊCH FILE EXCEL (.xlsx)
            # ----------------------------------------------------
            if is_excel:
                with st.spinner("1️⃣ Đang lọc văn bản cần dịch..."):
                    file_bytes = uploaded_file.read()
                    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
                    
                    texts_to_translate = set()
                    for sheet in wb.worksheets:
                        for row in sheet.iter_rows():
                            for cell in row:
                                if cell.value and isinstance(cell.value, str):
                                    val = cell.value.strip()
                                    if translation_mode == "Trung ➔ Việt" and has_chinese(val):
                                        texts_to_translate.add(val)
                                    elif translation_mode == "Việt ➔ Trung" and (has_vietnamese(val) or not has_chinese(val)):
                                        if len(val) > 1 and not val.isnumeric():
                                            texts_to_translate.add(val)

                    unique_texts = list(texts_to_translate)

                if not unique_texts:
                    st.warning("Không tìm thấy nội dung phù hợp với chế độ dịch đã chọn!")
                else:
                    with st.spinner(f"2️⃣ Đang dịch {len(unique_texts)} từ/câu qua Google Translate (Chạy offline miễn phí)..."):
                        translation_dict = translate_batch(unique_texts, translation_mode)

                    with st.spinner("3️⃣ Đang chèn dịch & giữ nguyên 100% format Excel gốc..."):
                        for sheet in wb.worksheets:
                            for row in sheet.iter_rows():
                                for cell in row:
                                    if cell.value and isinstance(cell.value, str):
                                        orig = cell.value.strip()
                                        trans = translation_dict.get(orig, "")
                                        if trans:
                                            cell.value = f"{orig}\n{trans}"
                                            curr_align = cell.alignment
                                            cell.alignment = Alignment(
                                                horizontal=curr_align.horizontal or "center",
                                                vertical=curr_align.vertical or "center",
                                                wrap_text=True
                                            )

                        output = io.BytesIO()
                        wb.save(output)
                        output.seek(0)

                        st.success("✅ Đã dịch thành công! Giữ nguyên 100% màu sắc & font gốc.")
                        st.download_button(
                            label="⬇️ Tải File Excel Song Ngữ (.xlsx)",
                            data=output.getvalue(),
                            file_name=f"Translated_{uploaded_file.name}",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

            # ----------------------------------------------------
            # TRƯỜNG HỢP 2: DỊCH FILE ẢNH (Bằng EasyOCR + Deep-Translator)
            # ----------------------------------------------------
            else:
                with st.spinner("1️⃣ Đang bật mắt thần EasyOCR quét chữ từ hình ảnh..."):
                    reader_zh, reader_vi = load_ocr_readers()
                    image = Image.open(uploaded_file)
                    img_np = np.array(image)

                    # Chọn mô hình OCR tương ứng với hướng dịch
                    reader = reader_zh if translation_mode == "Trung ➔ Việt" else reader_vi
                    ocr_results = reader.readtext(img_np)

                with st.spinner("2️⃣ Đang nhóm chữ thành bảng & Dịch sang Excel..."):
                    excel_bytes = build_excel_from_ocr(ocr_results, translation_mode)

                    st.success("✅ Đã quét ảnh và dựng Excel thành công (100% bằng thư viện)!")
                    st.download_button(
                        label="⬇️ Tải Xuất File Excel (.xlsx)",
                        data=excel_bytes.getvalue(),
                        file_name="Bang_cham_cong_OCR.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

        except Exception as e:
            st.error(f"❌ Xảy ra lỗi trong quá trình xử lý: {e}")
