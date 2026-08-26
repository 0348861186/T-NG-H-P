import streamlit as st
import pandas as pd
import easyocr
import numpy as np
from deep_translator import GoogleTranslator
from PIL import Image
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# 1. Cấu hình ứng dụng
st.set_page_config(page_title="Dịch Song Ngữ Trung - Việt -> Xuất Excel", layout="wide")
st.title("🌐 Dịch Song Ngữ Trung - Việt & Xuất File Excel")

# Cache bộ đọc OCR
@st.cache_resource
def get_ocr_reader(lang_tuple):
    return easyocr.Reader(list(lang_tuple), gpu=False)

# Hàm dịch an toàn chống văng lỗi
def safe_translate(text, source_lang, target_lang):
    cleaned_text = str(text).strip()
    if not cleaned_text or cleaned_text.isnumeric():
        return cleaned_text
    try:
        translated = GoogleTranslator(source=source_lang, target=target_lang).translate(cleaned_text)
        return translated if translated else cleaned_text
    except Exception:
        return cleaned_text

# Format đoạn văn bản song ngữ (Tiếng Trung bên trên, Tiếng Việt ngay bên dưới)
def format_bilingual(text, mode, src_lang, tgt_lang):
    if not str(text).strip():
        return text
    trans = safe_translate(text, src_lang, tgt_lang)
    if mode == "Trung - Việt":
        return f"{text}\n{trans}"
    else:
        return f"{trans}\n{text}"

# Hàm đóng gói thành file Excel có định dạng tự động bật xuống dòng (Wrap Text)
def export_to_styled_excel(dataframe, sheet_name="Song Ngu"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        dataframe.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Định dạng ô cho đẹp mắt
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        # Cấu hình tiêu đề
        for col_num, col_name in enumerate(dataframe.columns, 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Cấu hình dữ liệu & Tự động bật xuống dòng
        for row_num in range(2, len(dataframe) + 2):
            worksheet.row_dimensions[row_num].height = 35  # Tăng chiều cao hàng cho chữ 2 dòng
            for col_num in range(1, len(dataframe.columns) + 1):
                cell = worksheet.cell(row=row_num, column=col_num)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                cell.border = thin_border
                
        # Chỉnh độ rộng cột
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            worksheet.column_dimensions[col_letter].width = max(max_len + 5, 20)

    return output.getvalue()

# 2. Sidebar chọn file
st.sidebar.header("Tải File lên")
uploaded_file = st.sidebar.file_uploader(
    "Chọn file Ảnh (PNG, JPG) hoặc Excel (XLSX)", 
    type=["png", "jpg", "jpeg", "xlsx"]
)

if uploaded_file is not None:
    file_type = uploaded_file.name.split('.')[-1].lower()

    st.sidebar.subheader("Cấu hình Dịch")
    mode = st.sidebar.radio(
        "Chọn chế độ dịch:",
        ["Trung - Việt", "Việt - Trung"],
        help="Chọn 'Trung - Việt' nếu file gốc là Tiếng Trung, chọn 'Việt - Trung' nếu file gốc là Tiếng Việt."
    )

    src_lang = 'zh-CN' if mode == "Trung - Việt" else 'vi'
    tgt_lang = 'vi' if mode == "Trung - Việt" else 'zh-CN'

    # ------------------ 1. XỬ LÝ FILE EXCEL ------------------
    if file_type == "xlsx":
        st.subheader("📊 Dịch File Excel")
        df = pd.read_excel(uploaded_file)
        st.write("--- Preview dữ liệu gốc ---")
        st.dataframe(df.head())

        if st.button("🚀 Bắt đầu Dịch & Xuất Excel"):
            with st.spinner("Đang dịch dữ liệu Excel..."):
                df_translated = df.map(lambda x: format_bilingual(x, mode, src_lang, tgt_lang) if pd.notnull(x) else x)
            
            st.success("Dịch hoàn tất!")
            st.dataframe(df_translated)

            excel_data = export_to_styled_excel(df_translated, sheet_name="Excel_Song_Ngu")

            st.download_button(
                label="📥 Tải File Excel Song Ngữ",
                data=excel_data,
                file_name=f"translated_{uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # ------------------ 2. XỬ LÝ FILE ẢNH -> XUẤT EXCEL ------------------
    elif file_type in ["png", "jpg", "jpeg"]:
        st.subheader("🖼️ Dịch chữ trong Ảnh và Xuất ra File Excel")
        image = Image.open(uploaded_file).convert("RGB")
        img_np = np.array(image)

        st.image(image, caption="Ảnh gốc tải lên", width=500)

        if st.button("🚀 Nhận diện chữ, Dịch & Tạo File Excel"):
            with st.spinner("Đang đọc chữ từ ảnh và dịch song ngữ..."):
                # ĐÃ SỬA: Dùng 'ch_sim' thay cho 'zh_sim' chuẩn mã EasyOCR
                reader = get_ocr_reader(('ch_sim', 'en') if mode == "Trung - Việt" else ('vi', 'en'))
                results = reader.readtext(img_np)

                extracted_data = []
                for idx, (bbox, text, prob) in enumerate(results, 1):
                    if prob > 0.35 and len(str(text).strip()) > 0:
                        bilingual_text = format_bilingual(text, mode, src_lang, tgt_lang)
                        
                        # Tách riêng tiếng Trung và tiếng Việt
                        trans_text = safe_translate(text, src_lang, tgt_lang)
                        zh_text = text if mode == "Trung - Việt" else trans_text
                        vi_text = trans_text if mode == "Trung - Việt" else text

                        extracted_data.append({
                            "STT": idx,
                            "Nội Dung Song Ngữ (Trung trên / Việt dưới)": bilingual_text,
                            "Tiếng Trung (中文)": zh_text,
                            "Tiếng Việt": vi_text,
                            "Độ chính xác OCR (%)": round(prob * 100, 1)
                        })

                if extracted_data:
                    df_img = pd.DataFrame(extracted_data)
                    st.success(f"Đã trích xuất và dịch thành công {len(extracted_data)} đoạn văn bản từ ảnh!")
                    st.dataframe(df_img)

                    excel_data = export_to_styled_excel(df_img, sheet_name="Anh_Sang_Excel")

                    st.download_button(
                        label="📥 Tải File Excel Song Ngữ Từ Ảnh",
                        data=excel_data,
                        file_name=f"translated_from_image.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("Không tìm thấy văn bản đủ rõ trong ảnh.")
