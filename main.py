import streamlit as st
import pandas as pd
import easyocr
import cv2
import numpy as np
from deep_translator import GoogleTranslator
from PIL import Image
import io

# 1. Cấu hình trang Dashboard
st.set_page_config(page_title="Dịch Song Ngữ Trung - Việt", layout="wide")
st.title("🌐 Ứng dụng Dịch Song Ngữ Trung - Việt")

# Cache bộ đọc OCR để không bị load lại mỗi lần thao tác
@st.cache_resource
def load_ocr():
    return easyocr.Reader(['zh_sim', 'vi'], gpu=False)

reader = load_ocr()

# Hàm hỗ trợ dịch văn bản
def translate_text(text, mode):
    if not str(text).strip():
        return text
    try:
        if mode == "Trung - Việt":
            translated = GoogleTranslator(source='zh-CN', target='vi').translate(str(text))
            return f"{text}\n{translated}"  # Tiếng Trung bên trên, Tiếng Việt ngay bên dưới
        else:
            translated = GoogleTranslator(source='vi', target='zh-CN').translate(str(text))
            return f"{translated}\n{text}"  # Tiếng Trung bên trên, Tiếng Việt ngay bên dưới
    except Exception:
        return text

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
                # Áp dụng dịch trên từng ô để giữ nguyên cấu trúc khung/hàng/cột
                df_translated = df.applymap(lambda x: translate_text(x, mode) if pd.notnull(x) else x)
            
            st.success("Dịch hoàn tất!")
            st.write("--- Kết quả Dịch ---")
            st.dataframe(df_translated)

            # Xuất file Excel kết quả
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
        image = Image.open(uploaded_file)
        img_np = np.array(image)

        col1, col2 = st.columns(2)
        with col1:
            st.image(image, caption="Ảnh gốc", use_container_width=True)

        if st.button("🚀 Bắt đầu Dịch Ảnh"):
            with st.spinner("Đang nhận diện chữ và dịch..."):
                # Nhận diện vị trí và chữ trong ảnh
                results = reader.readtext(img_np)
                img_result = img_np.copy()

                for (bbox, text, prob) in results:
                    if prob > 0.3:  # Độ tin cậy
                        # Xử lý nội dung dịch song ngữ
                        if mode == "Trung - Việt":
                            trans = GoogleTranslator(source='zh-CN', target='vi').translate(text)
                            line1, line2 = text, trans
                        else:
                            trans = GoogleTranslator(source='vi', target='zh-CN').translate(text)
                            line1, line2 = trans, text

                        # Vẽ lại lên ảnh tại đúng vị trí khung chữ
                        pt1 = (int(bbox[0][0]), int(bbox[0][1]))
                        pt2 = (int(bbox[2][0]), int(bbox[2][1]))
                        
                        # Che chữ cũ bằng khung trắng
                        cv2.rectangle(img_result, pt1, pt2, (255, 255, 255), -1)
                        # Ghép chữ song ngữ dòng trên tiếng Trung, dòng dưới tiếng Việt
                        display_text = f"{line1} | {line2}"
                        cv2.putText(img_result, display_text, (pt1[0], pt1[1] + 15), 
                                    cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 0, 0), 1)

            with col2:
                st.image(img_result, caption="Ảnh sau khi dịch song ngữ", use_container_width=True)
