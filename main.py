import os
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins

# Tích hợp EasyOCR và Deep Translator
import easyocr
from deep_translator import GoogleTranslator

class AutoBilingualExcelGenerator:
    def __init__(self, input_file_path, output_file_path=None):
        self.input_path = input_file_path
        self.output_path = output_file_path or self._generate_output_name()
        
        # Khởi tạo mô hình OCR cho cả tiếng Trung (giản thể/phồn thể) và tiếng Việt
        self.ocr_reader = easyocr.Reader(['ch_sim', 'vi', 'en'], gpu=False)

    def _generate_output_name(self):
        base, _ = os.path.splitext(self.input_path)
        return f"{base}_Bilingual_Output.xlsx"

    def _detect_primary_language(self, text_list):
        """
        Nhận diện hướng dịch (Yêu cầu 3):
        - Nếu phần lớn là tiếng Trung -> Chế độ 'zh-CN' -> 'vi' (Trung - Việt)
        - Nếu phần lớn là tiếng Việt/Latinh -> Chế độ 'vi' -> 'zh-CN' (Việt - Trung)
        """
        full_text = "".join(str(t) for t in text_list)
        # Đếm số ký tự tiếng Trung trong chuỗi
        chinese_chars = re.findall(r'[\u4e00-\u9fff]', full_text)
        
        if len(chinese_chars) / max(len(full_text), 1) > 0.15:
            print("--> Phát hiện ngôn ngữ chính: Tiếng Trung (Dịch: Trung -> Việt)")
            return 'zh-CN', 'vi'
        else:
            print("--> Phát hiện ngôn ngữ chính: Tiếng Việt (Dịch: Việt -> Trung)")
            return 'vi', 'zh-CN'

    def _translate_text(self, text, src_lang, tgt_lang):
        """Thực hiện dịch thuật chính xác dựa trên hướng ngôn ngữ đã chọn"""
        if not text or str(text).strip() == "" or str(text).replace('.','',1).isdigit():
            return str(text), ""
        
        try:
            translated = GoogleTranslator(source=src_lang, target=tgt_lang).translate(str(text))
            return str(text).strip(), translated.strip()
        except Exception:
            return str(text).strip(), ""

    def process_file(self):
        ext = os.path.splitext(self.input_path)[-1].lower()
        extracted_data = []

        print(f"Đang xử lý file: {self.input_path}...")

        # 1. Nâng khả năng tùy biến theo định dạng file load lên (Yêu cầu 1 & 2)
        if ext in ['.jpg', '.jpeg', '.png', '.bmp']:
            # Đọc OCR nâng cao từ Ảnh
            results = self.ocr_reader.readtext(self.input_path, detail=0)
            extracted_data = [[res] for res in results]
            
        elif ext in ['.xlsx', '.xls']:
            # Đọc dữ liệu cấu trúc từ Excel
            df = pd.read_excel(self.input_path, header=None)
            extracted_data = df.fillna("").values.tolist()
        else:
            raise ValueError("Định dạng file không được hỗ trợ! Chỉ chấp nhận .png, .jpg, .xlsx, .xls")

        # 2. Tự động xác định hướng dịch (Yêu cầu 3)
        flat_text = [cell for row in extracted_data for cell in row if cell]
        src_lang, tgt_lang = self._detect_primary_language(flat_text)

        # 3. Tiến hành khởi tạo file Excel theo logic định dạng của code gốc
        wb = Workbook()
        ws = wb.active
        ws.title = "Bảng song ngữ"

        thin = Side(style="thin", color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for r_idx, row in enumerate(extracted_data, start=1):
            ws.row_dimensions[r_idx].height = 32
            
            for c_idx, cell_value in enumerate(row, start=1):
                cell = ws.cell(row=r_idx, column=c_idx)
                
                # Tiến hành dịch thuật nếu cell có nội dung chữ
                orig, trans = self._translate_text(cell_value, src_lang, tgt_lang)
                
                if trans and orig != trans:
                    cell.value = f"{orig}\n{trans}"
                else:
                    cell.value = orig

                # Áp dụng Logic Format từ Code Gốc
                cell.font = Font(name="Microsoft YaHei", size=10)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = border

                # Highlight dòng Header (dòng 1) theo sắc cam ED7D00 từ code gốc
                if r_idx == 1:
                    cell.font = Font(name="Microsoft YaHei", size=11, bold=True)
                    cell.fill = PatternFill("solid", fgColor="ED7D00")

        # Setup trang in và hiển thị theo logic code gốc
        ws.sheet_view.showGridLines = True
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins = PageMargins(left=0.2, right=0.2, top=0.3, bottom=0.3, header=0.1, footer=0.1)

        wb.save(output_file_path)
        print(f"-> Hoàn tất! File Excel song ngữ đã lưu tại: {output_file_path}")

# ==========================================
# THIẾT LẬP VÀ CHẠY THỬ
# ==========================================
if __name__ == "__main__":
    # Thay đổi đường dẫn file đầu vào của bạn tại đây (có thể là .png, .jpg hoặc .xlsx)
    input_file = "/mnt/data/Bang_cham_cong_2026-08-26_Trung_Viet.xlsx"
    output_file = "/mnt/data/Bang_cham_cong_SongNgu_Standard.xlsx"

    # Cài đặt thư viện bổ trợ nếu chưa có:
    # pip install easyocr pandas openpyxl deep-translator pillow

    processor = AutoBilingualExcelGenerator(input_file, output_file)
    processor.process_file()
