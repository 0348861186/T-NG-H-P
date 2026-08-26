import os
import re
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins

# Thư viện xử lý OCR và Dịch thuật
import easyocr
from deep_translator import GoogleTranslator

class AutoBilingualExcelGenerator:
    def __init__(self, input_file_path, output_file_path=None):
        self.input_path = input_file_path
        self.output_path = output_file_path or self._generate_output_name()

    def _generate_output_name(self):
        base, _ = os.path.splitext(self.input_path)
        return f"{base}_Bilingual_Output.xlsx"

    def _get_ocr_reader(self, is_chinese=True):
        """
        Khởi tạo Reader phù hợp với quy định của EasyOCR:
        - Tiếng Trung bắt buộc đi kèm Tiếng Anh: ['ch_sim', 'en']
        - Tiếng Việt đi kèm Tiếng Anh: ['vi', 'en']
        """
        if is_chinese:
            return easyocr.Reader(['ch_sim', 'en'], gpu=False)
        else:
            return easyocr.Reader(['vi', 'en'], gpu=False)

    def _read_image_ocr(self, image_path):
        """Quét ảnh thông minh để đảm bảo độ chính xác (Yêu cầu 2)"""
        # Bắt đầu quét bằng mô hình Trung - Anh
        reader_cn = self._get_ocr_reader(is_chinese=True)
        results = reader_cn.readtext(image_path, detail=0)
        
        full_text = "".join(str(r) for r in results)
        
        # Nếu phát hiện ký tự Trung Quốc -> Giữ nguyên kết quả
        if re.search(r'[\u4e00-\u9fff]', full_text):
            return results
        
        # Nếu không có chữ Trung, quét lại bằng mô hình Việt - Anh để chuẩn dấu
        reader_vi = self._get_ocr_reader(is_chinese=False)
        return reader_vi.readtext(image_path, detail=0)

    def _detect_primary_language(self, text_list):
        """
        Nhận diện hướng dịch tự động (Yêu cầu 3):
        - Nếu có tiếng Trung -> Dịch "Trung - Việt" (zh-CN -> vi)
        - Nếu là tiếng Việt -> Dịch "Việt - Trung" (vi -> zh-CN)
        """
        full_text = "".join(str(t) for t in text_list)
        chinese_chars = re.findall(r'[\u4e00-\u9fff]', full_text)
        
        if len(chinese_chars) / max(len(full_text), 1) > 0.15:
            st.info("🌐 Phát hiện ngôn ngữ chính: **Tiếng Trung** (Chế độ dịch: **Trung → Việt**)")
            return 'zh-CN', 'vi'
        else:
            st.info("🌐 Phát hiện ngôn ngữ chính: **Tiếng Việt** (Chế độ dịch: **Việt → Trung**)")
            return 'vi', 'zh-CN'

    def _translate_text(self, text, src_lang, tgt_lang):
        """Dịch nghĩa chính xác theo hướng ngôn ngữ"""
        if not text or str(text).strip() == "" or str(text).replace('.', '', 1).isdigit():
            return str(text), ""
        
        try:
            translated = GoogleTranslator(source=src_lang, target=tgt_lang).translate(str(text))
            return str(text).strip(), translated.strip()
        except Exception:
            return str(text).strip(), ""

    def process_file(self):
        ext = os.path.splitext(self.input_path)[-1].lower()
        extracted_data = []

        # 1. Tùy biến đọc file theo định dạng load lên (Yêu cầu 1)
        if ext in ['.jpg', '.jpeg', '.png', '.bmp']:
            results = self._read_image_ocr(self.input_path)
            extracted_data = [[res] for res in results]
            
        elif ext in ['.xlsx', '.xls']:
            df = pd.read_excel(self.input_path, header=None)
            extracted_data = df.fillna("").values.tolist()
        else:
            raise ValueError("Định dạng file không được hỗ trợ! Chỉ nhận file Ảnh (.png, .jpg) hoặc Excel (.xlsx, .xls)")

        # 2. Tự động nhận diện hướng dịch (Yêu cầu 3)
        flat_text = [cell for row in extracted_data for cell in row if cell]
        src_lang, tgt_lang = self._detect_primary_language(flat_text)

        # 3. Tạo file Excel theo đúng Style và Logic của Code gốc
        wb = Workbook()
        ws = wb.active
        ws.title = "Bảng song ngữ"

        thin = Side(style="thin", color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for r_idx, row in enumerate(extracted_data, start=1):
            ws.row_dimensions[r_idx].height = 32
            
            for c_idx, cell_value in enumerate(row, start=1):
                cell = ws.cell(row=r_idx, column=c_idx)
                
                # Thực hiện dịch nội dung
                orig, trans = self._translate_text(cell_value, src_lang, tgt_lang)
                
                if trans and orig != trans:
                    cell.value = f"{orig}\n{trans}"
                else:
                    cell.value = orig

                # Giữ nguyên Định dạng/Style theo Code gốc
                cell.font = Font(name="Microsoft YaHei", size=10)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = border

                # Header dòng 1 phủ màu cam ED7D00
                if r_idx == 1:
                    cell.font = Font(name="Microsoft YaHei", size=11, bold=True)
                    cell.fill = PatternFill("solid", fgColor="ED7D00")

        # Cài đặt thuộc tính trang in như code gốc
        ws.sheet_view.showGridLines = True
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins = PageMargins(left=0.2, right=0.2, top=0.3, bottom=0.3, header=0.1, footer=0.1)

        wb.save(self.output_path)
        return self.output_path

# ==========================================
# GIAO DIỆN STREAMLIT APP
# ==========================================
def main():
    st.set_page_config(page_title="Chuyển Đổi Excel Song Ngữ", page_icon="📊", layout="centered")
    st.title("📊 Tự Động Chuyển Đổi Bảng Song Ngữ (Trung - Việt)")

    uploaded_file = st.file_uploader(
        "Tải lên File Ảnh (PNG, JPG) hoặc File Excel (XLSX, XLS)", 
        type=["png", "jpg", "jpeg", "xlsx", "xls"]
    )

    if uploaded_file is not None:
        temp_input_path = f"temp_{uploaded_file.name}"
        with open(temp_input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        if st.button("🚀 Bắt đầu Xử lý & Tạo File Excel"):
            with st.spinner("Đang xử lý OCR, nhận diện ngôn ngữ và dịch thuật..."):
                try:
                    processor = AutoBilingualExcelGenerator(temp_input_path)
                    output_file_path = processor.process_file()

                    st.success("🎉 Xử lý thành công!")
                    
                    with open(output_file_path, "rb") as fp:
                        st.download_button(
                            label="📥 Tải xuống File Excel Song Ngữ",
                            data=fp,
                            file_name=os.path.basename(output_file_path),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {e}")
                finally:
                    if os.path.exists(temp_input_path):
                        os.remove(temp_input_path)

if __name__ == "__main__":
    main()
