import io
import re
import datetime
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from deep_translator import GoogleTranslator
from PIL import Image
import easyocr
import numpy as np

# ============================================================
# CẤU HÌNH STREAMLIT
# ============================================================
st.set_page_config(
    page_title="Hệ Thống Dịch Bảng Chấm Công (Không tốn Quota)",
    page_icon="🤖",
    layout="wide"
)

st.title("🤖 Dịch & Xuất Bảng Chấm Công Song Ngữ (Chuẩn Cấu Trúc Bảng)")

# ============================================================
# TẠO 2 READER RIÊNG BIỆT CHO TRUNG VÀ VIỆT
# ============================================================
@st.cache_resource
def load_ocr_reader_zh():
    return easyocr.Reader(['ch_sim', 'en'], gpu=False)

@st.cache_resource
def load_ocr_reader_vi():
    return easyocr.Reader(['vi', 'en'], gpu=False)

# ============================================================
# TỪ ĐIỂN CHUẨN HÓA CHUYÊN NGÀNH & SỬA LỖI OCR BÓC SAI CHỮ
# ============================================================
SPECIALIZED_DICT = {
    "部分": "Bộ phận",
    "部门": "Bộ phận",
    "宫分": "Bộ phận",  # Sửa lỗi EasyOCR quét sai 部门 thành 宫分
    "开几台机": "Số máy mở",
    "正式工": "Chính thức",
    "临时工": "Thời vụ",
    "备注": "Ghi chú",
    "连机": "Máy liên hợp",
    "制袋机": "Máy làm túi",
    "连机吹膜": "Máy thổi màng liên hợp",
    "制袋机吹膜": "Máy thổi màng làm túi",
    "吹膜": "Thổi màng",
    "制袋": "Làm túi",
    "一共": "Tổng cộng",
    "总计": "Tổng cộng"
}

# ============================================================
# BỘ LỌC HƯỚNG DỊCH & TẢI FILE
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
        "Tải lên Ảnh, PDF hoặc File Excel:", 
        type=["png", "jpg", "jpeg", "pdf", "xlsx"]
    )

def has_chinese(text):
    if not isinstance(text, str):
        return False
    return bool(re.search(r'[\u4e00-\u9fff]', text))

def has_vietnamese(text):
    if not isinstance(text, str):
        return False
    vietnamese_pattern = r'[àáảãạâầấẩẫậăằắẳẵặèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđĐ]'
    return bool(re.search(vietnamese_pattern, text, re.IGNORECASE))

def translate_single_text(text, mode):
    """
    Dịch 1 câu: Ưu tiên Tra Từ Điển Chuyên Ngành -> Sau đó mới dùng Google Translate.
    """
    if not text or not str(text).strip():
        return ""
    
    clean_txt = str(text).strip()
    
    # Tra từ điển sửa lỗi trước
    if clean_txt in SPECIALIZED_DICT:
        return SPECIALIZED_DICT[clean_txt]
    
    # Kiểm tra xem có cần dịch không
    should_trans = (mode == "Trung ➔ Việt" and has_chinese(clean_txt)) or \
                   (mode == "Việt ➔ Trung" and (has_vietnamese(clean_txt) or not has_chinese(clean_txt)))
                   
    if not should_trans or clean_txt.replace('.','',1).isdigit():
        return clean_txt
        
    try:
        src_lang = 'zh-CN' if mode == "Trung ➔ Việt" else 'vi'
        tgt_lang = 'vi' if mode == "Trung ➔ Việt" else 'zh-CN'
        translator = GoogleTranslator(source=src_lang, target=tgt_lang)
        res = translator.translate(clean_txt)
        return res if res else clean_txt
    except:
        return clean_txt

def translate_text_list(text_list, mode):
    translation_dict = {}
    for text in text_list:
        if not text:
            continue
        translation_dict[text] = translate_single_text(text, mode)
    return translation_dict

# ============================================================
# HÀM PHÂN TÍCH VÀ GOM NHÓM TỌA ĐỘ ẢNH (TABLE STRUCTURE FROM OCR)
# ============================================================
def process_ocr_image_to_table(reader, img_np, mode):
    """
    Phân tích vị trí Bounding Box (X, Y) của từng chữ để nhóm thành Dòng & Cột chuẩn.
    """
    # Đọc chi tiết bao gồm tọa độ
    results = reader.readtext(img_np, detail=1)
    if not results:
        return []
    
    # Lọc lấy danh sách (y_center, x_center, text)
    items = []
    for bbox, text, prob in results:
        text_str = text.strip()
        if not text_str:
            continue
        # bbox = [[x1, y1], [x2, y2], [x3, y3], [x4, y4]]
        y_center = (bbox[0][1] + bbox[2][1]) / 2.0
        x_center = (bbox[0][0] + bbox[2][0]) / 2.0
        items.append({'y': y_center, 'x': x_center, 'text': text_str})
    
    # Sắp xếp theo trục Y (từ trên xuống dưới)
    items = sorted(items, key=lambda i: i['y'])
    
    # Gom các chữ có cùng độ cao Y (chênh lệch nhau < 15px) vào cùng 1 dòng
    rows = []
    current_row = []
    last_y = None
    
    for item in items:
        if last_y is None or abs(item['y'] - last_y) < 18:
            current_row.append(item)
            last_y = item['y'] if last_y is None else (last_y + item['y']) / 2.0
        else:
            # Sắp xếp các ô trong dòng từ trái sang phải theo trục X
            current_row = sorted(current_row, key=lambda i: i['x'])
            rows.append(current_row)
            current_row = [item]
            last_y = item['y']
            
    if current_row:
        current_row = sorted(current_row, key=lambda i: i['x'])
        rows.append(current_row)
        
    return rows

# ============================================================
# DỰNG FILE EXCEL CHUẨN ĐẸP TỪ BẢNG DỮ LIỆU CÓ CẤU TRÚC
# ============================================================
def build_excel_from_ocr_rows(grouped_rows, mode):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    font_name = "Microsoft YaHei"
    orange_fill = PatternFill(fill_type="solid", fgColor="ED7D00")
    thin_side = Side(style="thin", color="000000")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # Dòng 1: Tiêu đề Bảng
    headers_vi = ["STT", "Bộ phận", "Số máy mở", "Chính thức", "Thời vụ", "Ghi chú"]
    headers_zh = ["STT", "部门", "开几台机", "正式工", "临时工", "备注"]
    
    top_headers = headers_zh if mode == "Trung ➔ Việt" else headers_vi
    bot_headers = headers_vi if mode == "Trung ➔ Việt" else headers_zh

    # Tạo Header Chuẩn 2 Ngôn Ngữ
    for col_idx in range(1, 7):
        cell = ws.cell(row=1, column=col_idx)
        th = top_headers[col_idx-1]
        bh = bot_headers[col_idx-1]
        cell.value = f"{th}\n{bh}" if th != bh else th
        cell.font = Font(name=font_name, size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = orange_fill
        cell.border = border
    ws.row_dimensions[1].height = 36

    current_row_idx = 2
    
    for r in grouped_rows:
        row_texts = [item['text'] for item in r]
        # Bỏ qua dòng nếu là dòng Header quét từ ảnh
        first_txt = "".join(row_texts)
        if "STT" in first_txt or "部门" in first_txt or "部分" in first_txt or "宫分" in first_txt:
            continue
            
        # Ghi dữ liệu vào từng ô tương ứng
        for col_idx, item in enumerate(r, start=1):
            if col_idx > 6:
                break
            orig = item['text']
            trans = translate_single_text(orig, mode)
            
            cell = ws.cell(row=current_row_idx, column=col_idx)
            cell.value = f"{orig}\n{trans}" if trans and trans != orig else orig
            cell.font = Font(name=font_name, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
            
        # Các cột trống còn lại kẻ border đầy đủ
        for c_idx in range(len(r) + 1, 7):
            cell = ws.cell(row=current_row_idx, column=c_idx)
            cell.border = border
            
        ws.row_dimensions[current_row_idx].height = 32
        current_row_idx += 1

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 15
    ws.column_dimensions["F"].width = 18

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# ============================================================
# XỬ LÝ CHÍNH
# ============================================================
if uploaded_file is not None:
    is_excel = uploaded_file.name.lower().endswith('.xlsx')
    button_label = f"🚀 Dịch ({translation_mode}) & Bảo Toàn Format Excel" if is_excel else f"🚀 OCR Quét Bảng & Dịch ({translation_mode})"
    
    if st.button(button_label, use_container_width=True):
        try:
            # ----------------------------------------------------
            # TRƯỜNG HỢP 1: FILE EXCEL (.xlsx)
            # ----------------------------------------------------
            if is_excel:
                with st.spinner(f"1️⃣ Đang quét ô cần dịch [{translation_mode}]..."):
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
                    st.warning("Không tìm thấy nội dung phù hợp!")
                else:
                    with st.spinner(f"2️⃣ Đang dịch {len(unique_texts)} văn bản..."):
                        translation_dict = translate_text_list(unique_texts, translation_mode)

                    with st.spinner("3️⃣ Đang ghi nhận nội dung dịch..."):
                        for sheet in wb.worksheets:
                            rows_modified = set()
                            for row in sheet.iter_rows():
                                for cell in row:
                                    if cell.value and isinstance(cell.value, str):
                                        orig = cell.value.strip()
                                        trans = translation_dict.get(orig, "")
                                        if trans and trans.strip() != orig:
                                            cell.value = f"{orig}\n{trans.strip()}"
                                            if cell.alignment:
                                                cell.alignment = Alignment(
                                                    horizontal=cell.alignment.horizontal,
                                                    vertical=cell.alignment.vertical,
                                                    wrap_text=True
                                                )
                                            else:
                                                cell.alignment = Alignment(wrap_text=True)
                                            rows_modified.add(cell.row)

                            for r_idx in rows_modified:
                                current_h = sheet.row_dimensions[r_idx].height
                                sheet.row_dimensions[r_idx].height = max(current_h * 1.8, 28) if current_h else 30

                        output = io.BytesIO()
                        wb.save(output)
                        output.seek(0)

                        st.success(f"✅ Đã dịch xong Excel!")
                        st.download_button(
                            label="⬇️ Tải File Excel Song Ngữ (.xlsx)",
                            data=output.getvalue(),
                            file_name=f"Translated_{uploaded_file.name}",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

            # ----------------------------------------------------
            # TRƯỜNG HỢP 2: FILE ẢNH / PDF (SỬA TRIỆT ĐỂ LỖI DỒN CỘT)
            # ----------------------------------------------------
            else:
                with st.spinner("1️⃣ Đang quét cấu trúc Bảng và vị trí chữ bằng EasyOCR..."):
                    file_bytes = uploaded_file.read()
                    
                    if uploaded_file.type == "application/pdf":
                        from pdf2image import convert_from_bytes
                        images = convert_from_bytes(file_bytes)
                        img = images[0]
                    else:
                        img = Image.open(io.BytesIO(file_bytes))
                    
                    reader = load_ocr_reader_zh() if translation_mode == "Trung ➔ Việt" else load_ocr_reader_vi()
                    img_np = np.array(img)
                    
                    # Gọi hàm nhóm tọa độ theo Hàng và Cột
                    grouped_rows = process_ocr_image_to_table(reader, img_np, translation_mode)

                if not grouped_rows:
                    st.warning("Không quét được bảng hoặc văn bản nào!")
                else:
                    with st.spinner("2️⃣ Đang dịch & Đưa chữ vào đúng ô/cột tương ứng..."):
                        excel_bytes = build_excel_from_ocr_rows(grouped_rows, translation_mode)

                        st.success("✅ Đã dựng lại cấu trúc Bảng Excel chuẩn xác!")
                        st.download_button(
                            label="⬇️ Tải File Excel Song Ngữ (.xlsx)",
                            data=excel_bytes.getvalue(),
                            file_name=f"Bang_cham_cong_OCR_{datetime.date.today()}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

        except Exception as e:
            st.error(f"❌ Lỗi xử lý: {e}")
