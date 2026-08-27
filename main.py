import io
import json
import re
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment
from google import genai

# ============================================================
# CẤU HÌNH STREAMLIT
# ============================================================
st.set_page_config(
    page_title="Hệ Thống Dịch & Xuất Excel Giữ Nguyên Định Dạng",
    page_icon="📑",
    layout="wide"
)

st.title("📑 Dịch Bảng Chấm Công / File Excel Song Ngữ (Bảo Toàn Định Dạng Gốc)")
st.caption("Tải file Excel gốc -> AI quét tất cả văn bản tiếng Trung -> Dịch & chèn dòng tiếng Việt ngay bên dưới -> Giữ nguyên 100% màu sắc, font, ô merge gốc.")

# ============================================================
# 1. CẤU HÌNH API KEY & TẢI FILE
# ============================================================
col1, col2 = st.columns([1, 2])

with col1:
    api_key = st.text_input("Nhập GEMINI_API_KEY:", type="password")

with col2:
    uploaded_file = st.file_uploader(
        "Tải lên file Excel (.xlsx):", 
        type=["xlsx"]
    )

# Hàm kiểm tra xem một chuỗi văn bản có chứa chữ Hán/Tiếng Trung không
def has_chinese(text):
    if not isinstance(text, str):
        return False
    return bool(re.search(r'[\u4e00-\u9fff]', text))

# ============================================================
# 2. XỬ LÝ DỊCH VÀ BẢO TOÀN ĐỊNH DẠNG EXCEL
# ============================================================
if uploaded_file is not None:
    if st.button("🚀 Dịch Song Ngữ & Xuất Excel Giữ Nguyên Style", use_container_width=True):
        if not api_key:
            st.error("Vui lòng nhập GEMINI_API_KEY!")
        else:
            try:
                with st.spinner("1️⃣ Đang đọc file Excel và trích xuất các ô chứa tiếng Trung..."):
                    # Mở Workbook từ file upload (giữ nguyên style)
                    file_bytes = uploaded_file.read()
                    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
                    
                    # Tập hợp các văn bản tiếng Trung cần dịch (tránh trùng lặp để tiết kiệm quota)
                    texts_to_translate = set()
                    
                    for sheet in wb.worksheets:
                        for row in sheet.iter_rows():
                            for cell in row:
                                # Chỉ đọc ô có giá trị chuỗi văn bản và chứa ký tự tiếng Trung
                                if cell.value and isinstance(cell.value, str) and has_chinese(cell.value):
                                    texts_to_translate.add(cell.value.strip())

                    unique_texts = list(texts_to_translate)

                if not uniquetexts_to_translate:
                    st.warning("Không tìm thấy văn bản tiếng Trung nào trong file Excel!")
                else:
                    st.info(f"Đã phát hiện {len(unique_texts)} chuỗi/câu tiếng Trung cần dịch.")

                    with st.spinner("2️⃣ AI Gemini đang phân tích và dịch sang tiếng Việt..."):
                        client = genai.Client(api_key=api_key)
                        
                        prompt = f"""
                        Bạn là một chuyên gia dịch thuật Trung - Việt chuyên nghiệp trong lĩnh vực nhân sự, quản lý xưởng và chấm công.
                        Hãy dịch danh sách các đoạn văn bản/từ ngữ tiếng Trung sau đây sang tiếng Việt.
                        
                        Danh sách tiếng Trung:
                        {json.dumps(unique_texts, ensure_ascii=False, indent=2)}

                        YÊU CẦU ĐẦU RA:
                        - Trả về kết quả dưới dạng MỘT JSON OBJECT duy nhất (không dùng markdown code blocks).
                        - Key là văn bản tiếng Trung gốc (chính xác từng ký tự).
                        - Value là bản dịch tiếng Việt tương ứng (ngắn gọn, chính xác bối cảnh chấm công/nhà xưởng).
                        
                        Ví dụ định dạng trả về:
                        {{
                            "员工上班": "Nhân viên đi làm",
                            "部门": "Bộ phận",
                            "临时工": "Công nhân thời vụ"
                        }}
                        """

                        response = client.models.generate_content(
                            model="gemini-2.5-flash",
                            contents=prompt
                        )

                        # Làm sạch chuỗi JSON trả về từ AI
                        clean_json = response.text.replace("```json", "").replace("```", "").strip()
                        translation_dict = json.loads(clean_json)

                    with st.spinner("3️⃣ Đang chèn dòng dịch tiếng Việt & bảo toàn màu sắc/khung ô..."):
                        # Duyệt qua từng ô trong Excel để chèn dịch tiếng Việt ngay bên dưới
                        for sheet in wb.worksheets:
                            for row in sheet.iter_rows():
                                for cell in row:
                                    if cell.value and isinstance(cell.value, str) and has_chinese(cell.value):
                                        original_text = cell.value.strip()
                                        translated_text = translation_dict.get(original_text, "")
                                        
                                        if translated_text:
                                            # Ghép Tiếng Trung ở trên, Tiếng Việt ở dưới
                                            cell.value = f"{original_text}\n{translated_text}"
                                            
                                            # Đảm bảo bật xuống dòng (wrap_text) nhưng GIỮ NGUYÊN align/font/color/fill gốc
                                            current_alignment = cell.alignment
                                            cell.alignment = Alignment(
                                                horizontal=current_alignment.horizontal or "center",
                                                vertical=current_alignment.vertical or "center",
                                                wrap_text=True
                                            )

                        # Lưu Workbook đã chèn bản dịch ra bộ nhớ tạm
                        output = io.BytesIO()
                        wb.save(output)
                        output.seek(0)

                        st.success("✅ Đã dịch thành công! File xuất ra giữ nguyên 100% định dạng gốc.")

                        # Tạo nút Tải Xuất File
                        st.download_button(
                            label="⬇️ Tải Xuất File Excel Song Ngữ (.xlsx)",
                            data=output.getvalue(),
                            file_name=f"Translated_{uploaded_file.name}",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

            except Exception as e:
                st.error(f"❌ Xảy ra lỗi trong quá trình xử lý: {e}")
