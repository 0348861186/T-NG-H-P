import streamlit as st
import pandas as pd
import easyocr
import numpy as np
from deep_translator import GoogleTranslator
from PIL import Image, ImageDraw, ImageFont
import io
import time

# 1. Cấu hình trang Dashboard
st.set_page_config(page_title="Dịch Song Ngữ Trung - Việt", layout="wide")
st.title("🌐 Ứng dụng Dịch Song Ngữ Trung - Việt")

# Cache bộ đọc OCR linh hoạt theo ngôn ngữ nguồn
@st.cache_resource
def get_ocr_reader(lang_tuple):
    return easyocr.Reader(list(lang_tuple), gpu=False)

# Hàm dịch an toàn chống crash app khi gặp văn bản lỗi
def safe_translate(text, source_lang, target_lang):
    # Làm sạch văn bản
    cleaned_text = str(text).strip()
    
    # Nếu văn bản rỗng, chứa duy nhất số hoặc ký tự quá ngắn -> Không dịch
    if not cleaned_text or cleaned_text.isnumeric() or len(cleaned_text) < 1:
        return cleaned_text
        
    try:
        translated = GoogleTranslator(source=source_lang, target=target_lang).translate(cleaned_text)
        return translated if translated else cleaned_text
    except Exception:
        # Nếu Google Translate lỗi, trả lại chính văn bản gốc thay vì văng ứng dụng
        return cleaned_text

# Hàm hỗ trợ dịch ô dữ liệu cho Excel
def translate_excel_cell(text, mode):
    if pd.isna(text):
        return text
    
    src = 'zh-CN' if mode == "Trung - Việt" else 'vi'
    tgt = 'vi' if mode == "Trung - Việt" else 'zh-CN'
    
    translated = safe_translate(text, src, tgt)
    
    if mode == "Trung - Việt":
        return f"{text}\n{translated}"
    else:
        return f"{translated}\n{text}"

# 2. Sidebar chọn file
st.sidebar.header("Tải File lên")
uploaded_file = st.sidebar.file_uploader(
    "Chọn file Ảnh (PNG, JPG) hoặc Excel (XLSX)", 
    type=["png", "jpg", "jpeg", "xlsx"]
)

if uploaded_file is not None:
    file_type = uploaded_file.name.split('.')[-1].lower()

    # Nút chọn hướng dịch
    st.sidebar.subheader("Cấu hình Dịch")
    mode = st.sidebar.radio(
        "Chọn chế độ dịch:",
        ["Trung - Việt", "Việt - Trung"],
        help="Chọn 'Trung - Việt' nếu file là Tiếng Trung, chọn 'Việt - Trung' nếu file là Tiếng Việt."
    )

    # ------------------ XỬ LÝ FILE EXCEL ------------------
    if file_type == "xlsx":
        st.subheader("📊 Xử lý File Excel")
        df = pd.read_excel(uploaded_file)
        st.write("--- Preview dữ liệu gốc ---")
        st.dataframe(df.head())

        if st.button("🚀 Bắt đầu Dịch Excel"):
            with st.spinner("Đang dịch toàn bộ dữ liệu Excel..."):
                df_translated = df.map(lambda x: translate_excel_cell(x, mode))
            
            st.success("Dịch hoàn tất!")
            st.write("--- Kết quả Dịch ---")
            st.dataframe(df_translated)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_translated.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Tải Excel Song Ngữ",
                data=output.getvalue(),
                file_name=f"translated_{uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # ------------------ XỬ LÝ FILE ẢNH ------------------
    elif file_type in ["png", "jpg", "jpeg"]:
        st.subheader("🖼️ Xử lý File Ảnh")
        image = Image.open(uploaded_file).convert("RGB")
        img_np = np.array(image)

        col1, col2 = st.columns(2)
        with col1:
            st.image(image, caption="Ảnh gốc", use_container_width=True)

        if st.button("🚀 Bắt đầu Dịch Ảnh"):
            with st.spinner("Đang xử lý hình ảnh và dịch..."):
                if mode == "Trung - Việt":
                    reader = get_ocr_reader(('zh_sim', 'en'))
                    src_lang, tgt_lang = 'zh-CN', 'vi'
                else:
                    reader = get_ocr_reader(('vi', 'en'))
                    src_lang, tgt_lang = 'vi', 'zh-CN'

                results = reader.readtext(img_np)
                
                img_result = image.copy()
                draw = ImageDraw.Draw(img_result)
                font = ImageFont.load_default()

                for (bbox, text, prob) in results:
                    # Lọc bỏ nhiễu: chỉ dịch các đoạn text nhận diện rõ ràng (> 35%)
                    if prob > 0.35 and len(str(text).strip()) > 0:
                        
                        # Gọi hàm dịch an toàn
                        trans = safe_translate(text, src_lang, tgt_lang)

                        if mode == "Trung - Việt":
                            line1, line2 = text, trans
                        else:
                            line1, line2 = trans, text

                        x_min = int(bbox[0][0])
                        y_min = int(bbox[0][1])
                        x_max = int(bbox[2][0])
                        y_max = int(bbox[2][1])

                        # Che văn bản cũ
                        draw.rectangle([x_min, y_min, x_max, y_max], fill="white")

                        # Ghép dòng song ngữ
                        display_text = f"{line1}\n{line2}"
                        draw.text((x_min, y_min), display_text, fill="black", font=font)

            with col2:
                st.image(img_result, caption="Ảnh sau khi dịch song ngữ", use_container_width=True)

                buf = io.BytesIO()
                img_result.save(buf, format="PNG")
                st.download_button(
                    label="📥 Tải Ảnh Kết Quả",
                    data=buf.getvalue(),
                    file_name=f"translated_{uploaded_file.name}",
                    mime="image/png"
                )
