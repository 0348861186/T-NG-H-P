import io
import json
import re
import datetime
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins
from google import genai
from google.genai import types

# ============================================================
# CẤU HÌNH STREAMLIT
# ============================================================
st.set_page_config(
    page_title="Hệ Thống Dịch Bảng Chấm Công Đa Định Dạng",
    page_icon="🤖",
    layout="wide"
)

st.title("🤖 Dịch & Xuất Bảng Chấm Công Song Ngữ (Ảnh / PDF / Excel)")
st.caption("Hỗ trợ tải lên cả File Excel gốc (bảo toàn 100% format) lẫn File Ảnh/PDF (AI quét OCR & dịch tự động).")

# ============================================================
# 1. CẤU HÌNH API KEY & TẢI FILE
# ============================================================
col1, col2 = st.columns([1, 2])

with col1:
    api_key = st.text_input("Nhập GEMINI_API_KEY:", type="password")

with col2:
    uploaded_file = st.file_uploader(
        "Tải lên Ảnh, PDF hoặc File Excel:", 
        type=["png", "jpg", "jpeg", "pdf", "xlsx"]
    )

# Hàm kiểm tra chuỗi có chứa chữ Hán/Tiếng Trung không
def has_chinese(text):
    if not isinstance(text, str):
        return False
    return bool(re.search(r'[\u4e00-\u9fff]', text))

# Hàm tự tạo file Excel chuẩn đẹp nếu đọc từ Ảnh / PDF
def build_excel_from_json(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    font_name = "Microsoft YaHei"
    orange_fill = PatternFill(fill_type="solid", fgColor="ED7D00")
    thin_side = Side(style="thin", color="000000")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    t_cn = data.get("title_cn", "")
    t_vi = data.get("title_vi", "")
    dt_str = data.get("date_str", "")
    rows = data.get("rows", [])

    # Header title
    full_title = f"{dt_str} {t_cn}\n{t_vi} ngày {dt_str}".strip()
    ws.merge_cells("A1:F1")
    ws["A1"] = full_title
    ws["A1"].font = Font(name=font_name, size=13, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 42

    # Headers cột
    headers = [
        ("STT", "STT"),
        ("部门", "Bộ phận"),
        ("开几台机", "Số máy mở"),
        ("正式工", "Chính thức"),
        ("临时工", "Thời vụ"),
        ("备注", "Ghi chú"),
    ]
    for col_idx, (cn, vi) in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col_idx)
        cell.value = f"{cn}\n{vi}" if cn != vi else cn
        cell.font = Font(name=font_name, size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = orange_fill
        cell.border = border
    ws.row_dimensions[2].height = 38

    current_row = 3
    total_workers = 0

    for row in rows:
        stt = row.get("stt", "")
        d_cn = str(row.get("dept_cn", "")) if row.get("dept_cn") else ""
        d_vi = str(row.get("dept_vi", "")) if row.get("dept_vi") else ""
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
        ws.cell(row=current_row, column=2, value=f"{d_cn}\n{d_vi}".strip())
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

    # Dòng Tổng cộng
    total_row = current_row
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=2)
    ws.cell(row=total_row, column=1, value="一共\nTổng cộng")
    ws.merge_cells(start_row=total_row, start_column=3, end_row=total_row, end_column=5)
    ws.cell(row=total_row, column=3, value=int(total_workers) if isinstance(total_workers, float) and total_workers.is_integer() else total_workers)

    for col in range(1, 7):
        c = ws.cell(row=total_row, column=col)
        c.font = Font(name=font_name, size=11, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    ws.row_dimensions[total_row].height = 36

    # Widths & Margins
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
# 2. XỬ LÝ DỊCH CHÍNH (PHÂN NHÁNH ẢNH / EXCEL)
# ============================================================
if uploaded_file is not None:
    is_excel = uploaded_file.name.lower().endswith('.xlsx')
    
    button_label = "🚀 Dịch & Bảo Toàn Định Dạng Excel" if is_excel else "🚀 AI Quét Ảnh/PDF & Dịch Xuất Excel"
    
    if st.button(button_label, use_container_width=True):
        if not api_key:
            st.error("Vui lòng nhập GEMINI_API_KEY!")
        else:
            try:
                client = genai.Client(api_key=api_key)

                # ----------------------------------------------------
                # TRƯỜNG HỢP 1: TẢI UP FILE EXCEL (.xlsx)
                # ----------------------------------------------------
                if is_excel:
                    with st.spinner("1️⃣ Đang quét các ô chứa tiếng Trung trong File Excel..."):
                        file_bytes = uploaded_file.read()
                        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
                        
                        texts_to_translate = set()
                        for sheet in wb.worksheets:
                            for row in sheet.iter_rows():
                                for cell in row:
                                    if cell.value and isinstance(cell.value, str) and has_chinese(cell.value):
                                        texts_to_translate.add(cell.value.strip())

                        unique_texts = list(texts_to_translate)

                    if not unique_texts:
                        st.warning("Không tìm thấy chữ tiếng Trung nào trong file Excel!")
                    else:
                        with st.spinner(f"2️⃣ AI đang dịch {len(unique_texts)} từ/câu sang Tiếng Việt..."):
                            prompt = f"""
                            Bạn là chuyên gia dịch thuật Trung - Việt về nhân sự, xưởng sản xuất và chấm công.
                            Hãy dịch danh sách tiếng Trung sau sang tiếng Việt.
                            
                            Danh sách tiếng Trung:
                            {json.dumps(unique_texts, ensure_ascii=False, indent=2)}

                            Trả về kết quả dưới dạng MỘT JSON OBJECT duy nhất (không dùng markdown code blocks).
                            Key là văn bản tiếng Trung gốc, Value là bản dịch tiếng Việt tương ứng.
                            """

                            response = client.models.generate_content(
                                model="gemini-2.5-flash",
                                contents=prompt
                            )

                            clean_json = response.text.replace("```json", "").replace("
