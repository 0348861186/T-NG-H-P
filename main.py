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
    page_title="Phần mềm dịch song ngữ chuẩn định dạng (Excel & Ảnh)",
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
st.title("🧮 Ứng dụng dịch song ngữ Trung - Việt (Giữ nguyên định dạng ô)")
st.markdown("Dịch song ngữ hiển thị gộp trong cùng một ô (trên/dưới), giữ nguyên cấu trúc bảng gốc.")

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

uploaded_file = st.file_uploader(
    "Tải lên file Excel (.xlsx) hoặc Ảnh (.png, .jpg)", type=["xlsx", "png", "jpg", "jpeg"]
)

if uploaded_file is not None:
    file_extension = uploaded_file.name.split(".")[-1].lower()

    # ================= TRƯỜNG HỢP 1: XỬ LÝ FILE EXCEL =================
    if file_extension == "xlsx":
        st.info("Đang xử lý file Excel và gộp song ngữ vào cùng một ô...")

        excel_bytes = uploaded_file.read()
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
        ws = wb.active

        with st.spinner("Đang dịch nội dung..."):
            max_row = ws.max_row
            max_col = ws.max_column

            for r in range(1, max_row + 1):
                # Tăng chiều cao dòng để hiển thị tốt 2 dòng chữ trong 1 ô
                ws.row_dimensions[r].height = 35 
                
                for c in range(1, max_col + 1):
                    cell = ws.cell(row=r, column=c)
                    if cell.value is not None and str(cell.value).strip() != "":
                        orig_val = str(cell.value)
                        # Kiểm tra nếu ô đã dịch hoặc là số thì bỏ qua dịch lại
                        trans_val = translate_text(orig_val, src, dest)
                        
                        # Gộp gốc và dịch vào chung 1 ô cách nhau bởi dấu xuống dòng (\n)
                        if src == "zh-cn":
                            cell.value = f"{orig_val}\n{trans_val}"
                        else:
                            cell.value = f"{orig_val}\n{trans_val}"
                            
                        # Đảm bảo bật tính năng ngắt dòng (wrap_text) để hiển thị đẹp
                        current_align = cell.alignment
                        h_align = current_align.horizontal if current_align and current_align.horizontal else 'center'
                        cell.alignment = Alignment(horizontal=h_align, vertical='center', wrap_text=True)
                        
                        # Tùy chỉnh font cho đẹp mắt (chữ dịch bên dưới nhỏ hơn và nghiêng nhẹ)
                        # Lưu ý: openpyxl không cho phép set font riêng từng dòng trong 1 ô đơn lẻ dễ dàng, 
                        # nên ta giữ nguyên font hoặc để size tiêu chuẩn cho cả ô.

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("Dịch và định dạng file Excel thành công!")
        st.download_button(
            label="📥 Tải xuống file Excel đã dịch",
            data=output,
            file_name=f"translated_bilingual_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ================= TRƯỜNG HỢP 2: XỬ LÝ FILE ẢNH =================
    elif file_extension in ["png", "jpg", "jpeg"]:
        st.info("Đang xử lý ảnh, bóc tách cấu trúc bảng và tái tạo file Excel chuẩn mẫu...")

        image_bytes = uploaded_file.read()
        image = Image.open(io.BytesIO(image_bytes))

        st.image(image, caption="Ảnh gốc tải lên", use_container_width=True)

        with st.spinner("Đang nhận diện chữ qua OCR..."):
            import easyocr
            reader = easyocr.Reader(['ch_sim', 'en'], gpu=False)
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

            # Sắp xếp theo chiều dọc
            data_rows = sorted(data_rows, key=lambda k: k["y"])

            table_matrix = []
            if data_rows:
                row_group = [data_rows[0]]
                last_y = data_rows[0]["y"]

                for item in data_rows[1:]:
                    if abs(item["y"] - last_y) < 18:  # Cùng một dòng bảng
                        row_group.append(item)
                    else:
                        row_group = sorted(row_group, key=lambda k: k["x"])
                        table_matrix.append([i["text"] for i in row_group])
                        row_group = [item]
                        last_y = item["y"]
                
                if row_group:
                    row_group = sorted(row_group, key=lambda k: k["x"])
                    table_matrix.append([i["text"] for i in row_group])

            # Tạo file Excel chuẩn giao diện giống ảnh mẫu
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"

            orange_header_fill = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
            header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            title_font = Font(name="Calibri", size=11, bold=True, color="000000")
            regular_font = Font(name="Calibri", size=10, color="000000")

            thin_border = Border(
                left=Side(style='thin', color='000000'),
                right=Side(style='thin', color='000000'),
                top=Side(style='thin', color='000000'),
                bottom=Side(style='thin', color='000000')
            )

            for r_idx, r_data in enumerate(table_matrix, start=1):
                ws.row_dimensions[r_idx].height = 35  # Đủ cao cho 2 dòng text trong 1 ô
                
                # Kiểm tra nếu là dòng tiêu đề chính (chứa ngày tháng năm)
                is_title = (len(r_data) == 1 and ("2026" in r_data[0] or "年" in r_data[0]))
                
                if is_title:
                    orig_title = r_data[0]
                    trans_title = translate_text(orig_title, src, dest)
                    cell = ws.cell(row=r_idx, column=1, value=f"{orig_title}\n{trans_title}")
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    cell.font = title_font
                else:
                    for c_idx, cell_text in enumerate(r_data, start=1):
                        cell = ws.cell(row=r_idx, column=c_idx)
                        cell.border = thin_border
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        
                        trans_text = translate_text(cell_text, src, dest)
                        
                        # Gộp gốc và dịch vào chung 1 ô
                        cell.value = f"{cell_text}\n{trans_text}"
                        
                        # Nếu là dòng tiêu đề bảng (header chứa STT, 部分...) thì tô màu cam
                        if r_idx == 1 or any("STT" in str(x) or "部分" in str(x) for x in r_data):
                            cell.fill = orange_header_fill
                            cell.font = header_font
                        else:
                            cell.font = regular_font

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)

            st.success("Đã tái tạo và dịch file Excel chuẩn định dạng song ngữ trong cùng 1 ô thành công!")
            
            st.write("### Xem trước dữ liệu:")
            df_preview = pd.DataFrame(table_matrix)
            st.dataframe(df_preview)

            st.download_button(
                label="📥 Tải xuống file Excel kết quả",
                data=output,
                file_name="translated_table_same_cell.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
