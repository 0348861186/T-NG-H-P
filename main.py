import streamlit as st
import pandas as pd
import easyocr
import numpy as np
from deep_translator import GoogleTranslator
from PIL import Image, ImageDraw, ImageFont
import io
import openpyxl

# 1. Cấu hình ứng dụng
st.set_page_config(page_title="Dịch Song Ngữ Trung - Việt", layout="wide")
st.title("🌐 Dịch Song Ngữ Trung - Việt")

# Cache bộ đọc OCR
@st.cache_resource
def get_ocr_reader(lang_tuple):
    return easyocr.Reader(list(lang_tuple), gpu=False)

# Hàm dịch an toàn chống văng ứng dụng
def safe_translate(text, source_lang, target_lang):
    cleaned_text = str(text).strip()
    if not cleaned_text or cleaned_text.isnumeric():
        return cleaned_text
    try:
        translated = GoogleTranslator(source=source_lang, target=target_lang).translate(cleaned_text)
        return translated if translated else cleaned_text
    except Exception:
        return cleaned_text

# Định dạng nội dung: Tiếng Trung luôn ở BÊN TRÊN, Tiếng Việt ở NGAY BÊN DƯỚI
def format_bilingual(text, mode):
    if not str(text).strip():
        return text
    
    if mode == "Trung - Việt":
        # Nguồn: Tiếng Trung, Đích: Tiếng Việt
        trans = safe_translate(text, 'zh-CN', 'vi')
        return f"{text}\n{trans}"
    else:
        # Nguồn: Tiếng Việt, Đích: Tiếng Trung
        trans = safe_translate(text, 'vi', 'zh-CN')
        return f"{trans}\n{text}"

# 2. Sidebar Tải file và Cấu hình
st.sidebar.header("Tải File")
uploaded_file = st.sidebar.file_uploader(
    "Chọn file Ảnh (PNG, JPG) hoặc Excel (XLSX)", 
    type=["png", "jpg", "jpeg", "xlsx"]
)

if uploaded_file is not None:
    file_type = uploaded_file.name.split('.')[-1].lower()

    st.sidebar.subheader("Kiểu Dịch")
    mode = st.sidebar.radio(
        "Chọn hướng dịch phù hợp với file:",
        ["Trung - Việt", "Việt - Trung"],
        help="Chọn 'Trung - Việt' nếu file gốc chứa tiếng Trung. Chọn 'Việt - Trung' nếu file gốc chứa tiếng Việt."
    )

    # ------------------ 1. XỬ LÝ FILE EXCEL ------------------
    if file_type == "xlsx":
        st.subheader("📊 Dịch File Excel (Giữ nguyên định dạng gốc)")
        df = pd.read_excel(uploaded_file)
        
        st.write("Dữ liệu gốc:")
        st.dataframe(df.head())

        if st.button("🚀 Bắt đầu Dịch"):
            with st.spinner("Đang dịch toàn bộ file Excel..."):
                # Duyệt giữ nguyên 100% cấu trúc hàng/cột của file gốc
                df_translated = df.map(lambda x: format_bilingual(x, mode) if pd.notnull(x) else x)
            
            st.success("Dịch hoàn tất!")
            st.dataframe(df_translated)

            # Xuất file Excel giữ nguyên khung cấu trúc
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_translated.to_excel(writer, index=False)
                
                # Bật tự động xuống dòng (Wrap Text) cho từng ô để hiển thị 2 dòng
                ws = writer.sheets['Sheet1']
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical='center')

            st.download_button(
                label="📥 Tải File Excel Song Ngữ",
                data=output.getvalue(),
                file_name=f"translated_{uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # ------------------ 2. XỬ LÝ FILE ẢNH ------------------
    elif file_type in ["png", "jpg", "jpeg"]:
        st.subheader("🖼️ Dịch File Ảnh (Giữ nguyên vị trí chữ trên ảnh)")
        image = Image.open(uploaded_file).convert("RGB")
        img_np = np.array(image)

        col1, col2 = st.columns(2)
        with col1:
            st.image(image, caption="Ảnh gốc", use_container_width=True)

        if st.button("🚀 Bắt đầu Dịch Ảnh"):
            with st.spinner("Đang nhận diện vị trí và dịch chữ..."):
                reader = get_ocr_reader(('ch_sim', 'en') if mode == "Trung - Việt" else ('vi', 'en'))
                results = reader.readtext(img_np)

                img_result = image.copy()
                draw = ImageDraw.Draw(img_result)
                font = ImageFont.load_default()

                for (bbox, text, prob) in results:
                    if prob > 0.35 and len(str(text).strip()) > 0:
                        bilingual_text = format_bilingual(text, mode)

                        x_min, y_min = int(bbox[0][0]), int(bbox[0][1])
                        x_max, y_max = int(bbox[2][0]), int(bbox[2][1])

                        # Che văn bản cũ
                        draw.rectangle([x_min, y_min, x_max, y_max], fill="white")
                        # Ghi chữ song ngữ tại đúng vị trí khung cũ
                        draw.text((x_min, y_min), bilingual_text, fill="black", font=font)

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
