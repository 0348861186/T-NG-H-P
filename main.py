import io
import cv2
import numpy as np
import pandas as pd
import streamlit as st
from deep_translator import GoogleTranslator
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Phần mềm dịch song ngữ Trung - Việt (Excel & Ảnh)",
    layout="wide",
)

# Hàm dịch văn bản
def translate_text(text, src_lang, dest_lang):
    if not text or str(text).strip() == "":
        return ""
    try:
        translator = GoogleTranslator(source=src_lang, target=dest_lang)
        if isinstance(text, (int, float)):
            return text
        translated = translator.translate(str(text))
        return translated
    except Exception as e:
        return text

# Giao diện ứng dụng
st.title("🧮 Ứng dụng dịch song ngữ Trung - Việt cho File Excel & Ảnh")
st.markdown("Hỗ trợ tải lên file Excel hoặc Ảnh chứa bảng, dịch tự động và xuất ra file Excel giữ nguyên định dạng.")

# Sidebar cấu hình
st.sidebar.header("Cấu hình dịch")
direction = st.sidebar.selectbox(
    "Chọn chiều dịch:",
    ("Trung -> Việt", "Việt -> Trung"),
)

if direction == "Trung -> Việt":
    src = "zh-cn"
    dest = "vi"
else:
    src = "vi"
    dest = "zh-cn"

option_style = st.sidebar.radio(
    "Kiểu hiển thị sau khi dịch:",
    ("Chỉ ghi đè bản dịch", "Hiển thị song ngữ (Dòng gốc + Dòng dịch)"),
)

uploaded_file = st.file_uploader(
    "Tải lên file Excel (.xlsx) hoặc Ảnh (.png, .jpg)", type=["xlsx", "png", "jpg", "jpeg"]
)

if uploaded_file is not None:
    file_extension = uploaded_file.name.split(".")[-1].lower()

    # ================= TRƯỜNG HỢP 1: XỬ LÝ FILE EXCEL =================
    if file_extension == "xlsx":
        st.info("Đang xử lý file Excel...")

        excel_bytes = uploaded_file.read()
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
        ws = wb.active

        with st.spinner("Đang tiến hành dịch nội dung các ô..."):
            if option_style == "Chỉ ghi đè bản dịch":
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value:
                            cell.value = translate_text(cell.value, src, dest)
            else:
                max_row = ws.max_row
                max_col = ws.max_column

                for r in range(max_row, 0, -1):
                    has_data = any(ws.cell(row=r, column=c).value for c in range(1, max_col + 1))
                    if has_data:
                        ws.insert_rows(r + 1)
                        for c in range(1, max_col + 1):
                            orig_cell = ws.cell(row=r, column=c)
                            new_cell = ws.cell(row=r + 1, column=c)

                            if orig_cell.font:
                                new_cell.font = Font(
                                    name=orig_cell.font.name,
                                    size=max(8, orig_cell.font.size - 1),
                                    italic=True,
                                    color="555555"
                                )
                            if orig_cell.alignment:
                                new_cell.alignment = Alignment(
                                    horizontal=orig_cell.alignment.horizontal,
                                    vertical=orig_cell.alignment.vertical
                                )
                            if orig_cell.border:
                                new_cell.border = orig_cell.border
                            if orig_cell.fill:
                                new_cell.fill = orig_cell.fill

                            if orig_cell.value:
                                new_cell.value = translate_text(orig_cell.value, src, dest)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("Dịch thành công file Excel!")
        st.download_button(
            label="📥 Tải xuống file Excel đã dịch",
            data=output,
            file_name=f"translated_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ================= TRƯỜNG HỢP 2: XỬ LÝ FILE ẢNH =================
    elif file_extension in ["png", "jpg", "jpeg"]:
        st.info("Đang xử lý ảnh (Sử dụng OCR để bóc tách văn bản và tạo bảng Excel)...")

        image_bytes = uploaded_file.read()
        image = Image.open(io.BytesIO(image_bytes))

        st.image(image, caption="Ảnh gốc tải lên", use_container_width=True)

        with st.spinner("Đang nhận diện chữ trong ảnh (OCR)..."):
            import easyocr
            reader = easyocr.Reader(['ch_sim', 'en', 'vi'], gpu=False)
            img_np = np.array(image)
            results = reader.readtext(img_np)

        if not results:
            st.warning("Không tìm thấy văn bản nào trong ảnh. Vui lòng thử ảnh rõ nét hơn.")
        else:
            data_rows = []
            for bbox, text, prob in results:
                y_center = (bbox[0][1] + bbox[2][1]) / 2
                x_center = (bbox[0][0] + bbox[1][0]) / 2
                data_rows.append({"y": y_center, "x": x_center, "text": text})

            data_rows = sorted(data_rows, key=lambda k: k["y"])

            table_matrix = []
            if data_rows:
                row_group = [data_rows[0]]
                last_y = data_rows[0]["y"]

                for item in data_rows[1:]:
                    if abs(item["y"] - last_y) < 18:  # cùng dòng
                        row_group.append(item)
                    else:
                        row_group = sorted(row_group, key=lambda k: k["x"])
                        table_matrix.append([i["text"] for i in row_group])
                        row_group = [item]
                        last_y = item["y"]
                
                if row_group:
                    row_group = sorted(row_group, key=lambda k: k["x"])
                    table_matrix.append([i["text"] for i in row_group])

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"

            orange_fill = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
            white_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            thin_border = Border(
                left=Side(style='thin', color='000000'),
                right=Side(style='thin', color='000000'),
                top=Side(style='thin', color='000000'),
                bottom=Side(style='thin', color='000000')
            )

            for r_idx, r_data in enumerate(table_matrix, start=1):
                translated_row = []
                for c_idx, cell_text in enumerate(r_data, start=1):
                    cell = ws.cell(row=r_idx * 2 - 1, column=c_idx, value=cell_text)
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    if r_idx == 1:
                        cell.fill = orange_fill
                        cell.font = white_font

                    trans_text = translate_text(cell_text, src, dest)
                    translated_row.append(trans_text)

                for c_idx, trans_text in enumerate(translated_row, start=1):
                    cell_trans = ws.cell(row=r_idx * 2, column=c_idx, value=trans_text)
                    cell_trans.border = thin_border
                    cell_trans.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    cell_trans.font = Font(name="Calibri", size=10, italic=True, color="333333")

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)

            st.success("Đã trích xuất và dịch bảng từ ảnh thành công!")
            
            st.write("### Bảng dữ liệu trích xuất & dịch xem trước:")
            df_preview = pd.DataFrame(table_matrix)
            st.dataframe(df_preview)

            st.download_button(
                label="📥 Tải xuống file Excel kết quả",
                data=output,
                file_name="translated_from_image.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
