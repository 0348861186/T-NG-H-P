import streamlit as st
import pandas as pd
import easyocr
import numpy as np
from deep_translator import GoogleTranslator
from PIL import Image, ImageDraw, ImageFont
import io

# 1. Cấu hình trang Dashboard
st.set_page_config(page_title="Dịch Song Ngữ Trung - Việt", layout="wide")
st.title("🌐 Ứng dụng Dịch Song Ngữ Trung - Việt")

# Cache bộ đọc OCR linh hoạt theo ngôn ngữ nguồn
@st.cache_resource
def get_ocr_reader(lang_tuple):
    # Truyền tuple ngôn ngữ vào EasyOCR
    return easyocr.Reader(list(lang_tuple), gpu=False)

# Hàm hỗ trợ dịch văn bản đơn thuần (cho Excel)
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
        image = Image.open(uploaded_file).convert("RGB")
        img_np = np.array(image)

        col1, col2 = st.columns(2)
        with col1:
            st.image(image, caption="Ảnh gốc", use_container_width=True)

        if st.button("🚀 Bắt đầu Dịch Ảnh"):
            with st.spinner("Đang khởi tạo OCR và dịch..."):
                # Tải động OCR reader theo hướng dịch (CÁCH 1)
                if mode == "Trung - Việt":
                    reader = get_ocr_reader(('zh_sim', 'en'))
                else:
                    reader = get_ocr_reader(('vi', 'en'))

                # Nhận diện vị trí và chữ trong ảnh
                results = reader.readtext(img_np)
                
                # Tạo đối tượng vẽ bằng Pillow để hỗ trợ font tiếng Việt/Trung
                img_result = image.copy()
                draw = ImageDraw.Draw(img_result)
                
                # Nạp font mặc định
                font = ImageFont.load_default()

                for (bbox, text, prob) in results:
                    if prob > 0.3:  # Chỉ xử lý các vùng chữ có độ tin cậy > 30%
                        # Xử lý nội dung dịch song ngữ
                        if mode == "Trung - Việt":
                            trans = GoogleTranslator(source='zh-CN', target='vi').translate(text)
                            line1, line2 = text, trans
                        else:
                            trans = GoogleTranslator(source='vi', target='zh-CN').translate(text)
                            line1, line2 = trans, text

                        # Xác định tọa độ khung chữ
                        x_min = int(bbox[0][0])
                        y_min = int(bbox[0][1])
                        x_max = int(bbox[2][0])
                        y_max = int(bbox[2][1])

                        # 1. Vẽ hình chữ nhật nền trắng che văn bản cũ
                        draw.rectangle([x_min, y_min, x_max, y_max], fill="white")

                        # 2. Chuẩn bị văn bản 2 dòng (Trung trên, Việt dưới)
                        display_text = f"{line1}\n{line2}"

                        # 3. Vẽ văn bản song ngữ lên vị trí cũ
                        draw.text((x_min, y_min), display_text, fill="black", font=font)

            with col2:
                st.image(img_result, caption="Ảnh sau khi dịch song ngữ", use_container_width=True)

                # Nút tải ảnh kết quả
                buf = io.BytesIO()
                img_result.save(buf, format="PNG")
                st.download_button(
                    label="📥 Tải Ảnh Kết Quả",
                    data=buf.getvalue(),
                    file_name=f"translated_{uploaded_file.name}",
                    mime="image/png"
                )
