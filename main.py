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

st.title("🤖 Dịch & Xuất Bảng Chấm Công Song Ngữ (Bảo Toàn Định Dạng Gốc)")
st.caption("Hỗ trợ chọn chế độ Trung ➔ Việt hoặc Việt ➔ Trung | Giữ nguyên 100% format Excel gốc.")

# ============================================================
# TẠO 2 READER RIÊNG BIỆT CHO TRUNG VÀ VIỆT (TRÁNH LỖI VALUEERROR)
# ============================================================
@st.cache_resource
def load_ocr_reader_zh():
    """Model OCR chuyên cho Tiếng Trung + Tiếng Anh"""
    return easyocr.Reader(['ch_sim', 'en'], gpu=False)

@st.cache_resource
def load_ocr_reader_vi():
    """Model OCR chuyên cho Tiếng Việt + Tiếng Anh"""
    return easyocr.Reader(['vi', 'en'], gpu=False)

# ============================================================
# 1. BỘ LỌC HƯỚNG DỊCH & TẢI FILE
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

# Hàm kiểm tra chuỗi có chứa chữ Hán / Tiếng Trung không
def has_chinese(text):
    if not isinstance(text, str):
        return False
    return bool(re.search(r'[\u4e00-\u9fff]', text))

# Hàm kiểm tra chuỗi có chứa tiếng Việt không
def has_vietnamese(text):
    if not isinstance(text, str):
        return False
    vietnamese_pattern = r'[àáảãạâầấẩẫậăằắẳẵặèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđĐ]'
    return bool(re.search(vietnamese_pattern, text, re.IGNORECASE))

# ============================================================
# HÀM DỊCH SỬ DỤNG THƯ VIỆN DEEP-TRANSLATOR (TỐI ƯU CÂU)
# ============================================================
def translate_text_list(text_list, mode):
    """
    Dịch danh sách các văn bản, tách theo từng câu/dòng để đảm bảo không bỏ sót ô nào.
    """
    src_lang = 'zh-CN' if mode == "Trung ➔ Việt" else 'vi'
    tgt_lang = 'vi' if mode == "Trung ➔ Việt" else 'zh-CN'
    
    translator = GoogleTranslator(source=src_lang, target=tgt_lang)
    translation_dict = {}
    
    for text in text_list:
        if not text or not str(text).strip():
            continue
        try:
            # Nếu chuỗi chứa nhiều dòng (\n), dịch từng dòng nhỏ rồi ghép lại
            sub_lines = str(text).split('\n')
            translated_sub_lines = []
            for line in sub_lines:
                line_str = line.strip()
                if not line_str:
                    translated_sub_lines.append("")
                    continue
                
                # Kiểm tra điều kiện có cần dịch dòng này không
                should_trans = (mode == "Trung ➔ Việt" and has_chinese(line_str)) or \
                               (mode == "Việt ➔ Trung" and (has_vietnamese(line_str) or not has_chinese(line_str)))
                
                if should_trans and not line_str.replace('.','',1).isdigit():
                    t_line = translator.translate(line_str)
                    translated_sub_lines.append(t_line if t_line else line_str)
                else:
                    translated_sub_lines.append(line_str)
            
            translation_dict[text] = "\n".join(translated_sub_lines)
        except Exception as e:
            translation_dict[text] = text  # Nếu lỗi kết nối thì giữ nguyên

    return translation_dict

# Hàm tự tạo file Excel chuẩn đẹp khi đọc dữ liệu từ Ảnh / PDF
def build_excel_from_json(data, mode):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    font_name = "Microsoft YaHei"
    orange_fill = PatternFill(fill_type="solid", fgColor="ED7D00")
    thin_side = Side(style="thin", color="000000")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    t_src = data.get("title_src", "")
    t_tgt = data.get("title_tgt", "")
    dt_str = data.get("date_str", "")
    rows = data.get("rows", [])

    top_title = t_src if mode == "Trung ➔ Việt" else t_tgt
    bot_title = t_tgt if mode == "Trung ➔ Việt" else t_src

    full_title = f"{dt_str} {top_title}\n{bot_title} ngày {dt_str}".strip()
    ws.merge_cells("A1:F1")
    ws["A1"] = full_title
    ws["A1"].font = Font(name=font_name, size=13, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 42

    if mode == "Trung ➔ Việt":
        headers = [("STT", "STT"), ("部门", "Bộ phận"), ("开几台机", "Số máy mở"), ("正式工", "Chính thức"), ("临时工", "Thời vụ"), ("备注", "Ghi chú")]
    else:
        headers = [("STT", "STT"), ("Bộ phận", "部门"), ("Số máy mở", "开几台机"), ("Chính thức", "正式工"), ("Thời vụ", "临时工"), ("Ghi chú", "备注")]

    for col_idx, (top_h, bot_h) in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col_idx)
        cell.value = f"{top_h}\n{bot_h}" if top_h != bot_h else top_h
        cell.font = Font(name=font_name, size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = orange_fill
        cell.border = border
    ws.row_dimensions[2].height = 38

    current_row = 3
    total_workers = 0

    for row in rows:
        stt = row.get("stt", "")
        d_src = str(row.get("dept_src", "")) if row.get("dept_src") else ""
        d_tgt = str(row.get("dept_tgt", "")) if row.get("dept_tgt") else ""
        mac = row.get("machines", "") or ""
        fml = row.get("formal", "") or ""
        tmp = row.get("temp", "") or ""
        rmk = str(row.get("remark", "")) if row.get("remark") else ""

        try:
            if fml: total_workers += float(fml)
            if tmp: total_workers += float(tmp)
        except:
            pass

        ws.cell(row=current_row, column=1, value=stt)
        ws.cell(row=current_row, column=2, value=f"{d_src}\n{d_tgt}".strip())
        ws.cell(row=current_row, column=3, value=mac)
        ws.cell(row=current_row, column=4, value=fml)
        ws.cell(row=current_row, column=5, value=tmp)
        ws.cell(row=current_row, column=6, value=rmk)

        for col in range(1, 7):
            c = ws.cell(row=current_row, column=col)
            c.font = Font(name=font_name, size=10)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = border

        ws.row_dimensions[current_row].height = 32
        current_row += 1

    total_row = current_row
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=2)
    ws.cell(row=total_row, column=1, value="一共\nTổng cộng" if mode == "Trung ➔ Việt" else "Tổng cộng\n一共")
    ws.merge_cells(start_row=total_row, start_column=3, end_row=total_row, end_column=5)
    ws.cell(row=total_row, column=3, value=int(total_workers) if isinstance(total_workers, float) and total_workers.is_integer() else total_workers)

    for col in range(1, 7):
        c = ws.cell(row=total_row, column=col)
        c.font = Font(name=font_name, size=11, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    ws.row_dimensions[total_row].height = 36

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 17
    ws.column_dimensions["E"].width = 17
    ws.column_dimensions["F"].width = 18

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# ============================================================
# 2. XỬ LÝ DỊCH CHÍNH (BẢO TOÀN FORMAT EXCEL GỐC)
# ============================================================
if uploaded_file is not None:
    is_excel = uploaded_file.name.lower().endswith('.xlsx')
    
    button_label = f"🚀 Dịch ({translation_mode}) & Bảo Toàn Format Excel" if is_excel else f"🚀 Thư viện OCR Quét Ảnh/PDF & Dịch ({translation_mode})"
    
    if st.button(button_label, use_container_width=True):
        try:
            # ----------------------------------------------------
            # TRƯỜNG HỢP 1: FILE EXCEL (.xlsx)
            # ----------------------------------------------------
            if is_excel:
                with st.spinner(f"1️⃣ Đang quét danh sách các ô cần dịch [{translation_mode}]..."):
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
                    with st.spinner(f"2️⃣ Đang dịch {len(unique_texts)} văn bản bằng Google Translate..."):
                        translation_dict = translate_text_list(unique_texts, translation_mode)

                    with st.spinner("3️⃣ Đang ghi nhận nội dung dịch & giữ nguyên 100% định dạng gốc..."):
                        for sheet in wb.worksheets:
                            rows_modified = set()
                            
                            for row in sheet.iter_rows():
                                for cell in row:
                                    if cell.value and isinstance(cell.value, str):
                                        orig = cell.value.strip()
                                        trans = translation_dict.get(orig, "")
                                        
                                        # Nếu có kết quả dịch và khác văn bản gốc
                                        if trans and trans.strip() != orig:
                                            cell.value = f"{orig}\n{trans.strip()}"
                                            
                                            # Bật wrap_text để hiển thị xuống dòng, GIỮ NGUYÊN căn lề gốc của ô
                                            if cell.alignment:
                                                cell.alignment = Alignment(
                                                    horizontal=cell.alignment.horizontal,
                                                    vertical=cell.alignment.vertical,
                                                    wrap_text=True
                                                )
                                            else:
                                                cell.alignment = Alignment(wrap_text=True)
                                                
                                            rows_modified.add(cell.row)

                            # Tăng chiều cao của các dòng có chứa ô đã dịch để hiển thị đủ 2 dòng chữ
                            for r_idx in rows_modified:
                                current_h = sheet.row_dimensions[r_idx].height
                                if current_h is not None:
                                    sheet.row_dimensions[r_idx].height = max(current_h * 1.8, 28)
                                else:
                                    sheet.row_dimensions[r_idx].height = 30  # Mặc định chiều cao vừa đẹp cho 2 dòng

                        output = io.BytesIO()
                        wb.save(output)
                        output.seek(0)

                        st.success(f"✅ Đã dịch xong! File Excel giữ nguyên 100% định dạng, font chữ và màu sắc gốc.")
                        st.download_button(
                            label="⬇️ Tải File Excel Song Ngữ (.xlsx)",
                            data=output.getvalue(),
                            file_name=f"Translated_{uploaded_file.name}",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

            # ----------------------------------------------------
            # TRƯỜNG HỢP 2: FILE ẢNH / PDF (EASYOCR)
            # ----------------------------------------------------
            else:
                with st.spinner("1️⃣ Đang quét chữ bằng EasyOCR..."):
                    file_bytes = uploaded_file.read()
                    
                    if uploaded_file.type == "application/pdf":
                        from pdf2image import convert_from_bytes
                        images = convert_from_bytes(file_bytes)
                        img = images[0]
                    else:
                        img = Image.open(io.BytesIO(file_bytes))
                    
                    if translation_mode == "Trung ➔ Việt":
                        current_reader = load_ocr_reader_zh()
                    else:
                        current_reader = load_ocr_reader_vi()

                    img_np = np.array(img)
                    ocr_results = current_reader.readtext(img_np, detail=0)
                    extracted_texts = [text.strip() for text in ocr_results if text.strip()]

                if not extracted_texts:
                    st.warning("Không tìm thấy văn bản nào trong ảnh/PDF!")
                else:
                    with st.spinner("2️⃣ Đang dịch danh sách văn bản..."):
                        translation_dict = translate_text_list(extracted_texts, translation_mode)
                        
                        rows_data = []
                        for idx, src_txt in enumerate(extracted_texts, 1):
                            rows_data.append({
                                "stt": idx,
                                "dept_src": src_txt,
                                "dept_tgt": translation_dict.get(src_txt, ""),
                                "machines": "",
                                "formal": "",
                                "temp": "",
                                "remark": ""
                            })

                        parsed_data = {
                            "title_src": "BẢNG CHẤM CÔNG",
                            "title_tgt": "考勤表",
                            "date_str": datetime.date.today().strftime("%Y-%m-%d"),
                            "rows": rows_data
                        }

                    with st.spinner("3️⃣ Đang dựng bảng Excel..."):
                        excel_bytes = build_excel_from_json(parsed_data, translation_mode)

                        st.success(f"✅ Đã quét chữ và tạo Excel ({translation_mode}) thành công!")
                        st.download_button(
                            label="⬇️ Tải File Excel (.xlsx)",
                            data=excel_bytes.getvalue(),
                            file_name=f"Bang_cham_cong_OCR_{parsed_data.get('date_str', 'export')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

        except Exception as e:
            st.error(f"❌ Xảy ra lỗi trong quá trình xử lý: {e}")
