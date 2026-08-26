import io
import re
import google.generativeai as genai
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from PIL import Image
import streamlit as st

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Dịch Bảng Song Ngữ Trung - Việt",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Dịch Bảng Song Ngữ Trung - Việt (Giữ Nguyên Cấu Trúc Bảng)")

# Sidebar - Nhập API Key & Cấu hình
st.sidebar.header("⚙️ Cấu Hình Hệ Thống")
gemini_api_key = st.sidebar.text_input(
    "Nhập Gemini API Key (Bắt buộc cho file Ảnh):", type="password"
)

direction = st.sidebar.selectbox(
    "Chọn hướng dịch:",
    [
        "Trung sang Việt (Chinese -> Vietnamese)",
        "Việt sang Trung (Vietnamese -> Chinese)",
    ],
)

src_lang = "Tiếng Trung" if "Trung sang Việt" in direction else "Tiếng Việt"
target_lang = "Tiếng Việt" if "Trung sang Việt" in direction else "Tiếng Trung"


# --- HÀM 1: ĐỌC VÀ BỎ BẢNG TỪ ẢNH BẰNG GEMINI VISION ---
def extract_table_from_image(image, api_key, src_l, tgt_l):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")

    prompt = f"""
    Bạn là một chuyên gia xử lý bảng biểu. Hãy nhìn vào ảnh bảng biểu này và trích xuất toàn bộ dữ liệu thành dạng bảng.
    YÊU CẦU ĐẶC BIỆT:
    1. Giữ nguyên số dòng và số cột của bảng trong ảnh.
    2. Với MỖI Ô có chứa văn bản ({src_l}), hãy dịch văn bản đó sang {tgt_l}.
    3. Định dạng MỖI Ô theo quy tắc:
       Dòng 1: Văn bản gốc ({src_l})
       Dòng 2: Văn bản dịch ({tgt_l})
       (Hai dòng cách nhau bằng ký tự xuống dòng \\n).
    4. Các ô chỉ chứa số (như STT, số lượng) hoặc ô trống thì GIỮ NGUYÊN, không thêm dòng dịch.
    5. Trả về kết quả dưới dạng cấu trúc CSV (phân cách bằng dấu phẩy, các chuỗi có xuống dòng phải bọc trong dấu ngoặc kép đôi ""). KHÔNG viết thêm bất kỳ lời dẫn nào khác.
    """

    response = model.generate_content([prompt, image])
    content = response.text.strip()

    # Làm sạch chuỗi markdown nếu có
    if content.startswith("```csv"):
        content = content[6:]
    if content.startswith("```"):
        content = content[3:]
    if content.endswith("```"):
        content = content[:-3]

    df = pd.read_csv(io.StringIO(content.strip()), header=None)
    return df


# --- HÀM 2: DỊCH FILE EXCEL VÀ GHÉP DÒNG SONG NGỮ ---
def process_excel_file(df_raw, api_key, src_l, tgt_l):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")

    # Chuyển dataframe thành chuỗi JSON để gửi cho AI xử lý theo lô
    raw_json = df_raw.to_json(orient="values")

    prompt = f"""
    Dưới đây là mảng dữ liệu 2 chiều đại diện cho một bảng Excel:
    {raw_json}

    Hãy xử lý từng ô trong mảng:
    - Nếu ô chứa văn bản ({src_l}): Hãy dịch sang {tgt_l} và gộp thành chuỗi "Văn bản gốc\\nVăn bản dịch".
    - Nếu ô là số, ký hiệu hoặc ô trống: Giữ nguyên.
    - Giữ nguyên toàn bộ số dòng và số cột.
    
    Trả về kết quả dưới dạng JSON mảng 2 chiều tương tự đầu vào, không kèm lời dẫn.
    """

    response = model.generate_content(prompt)
    content = response.text.strip()

    if content.startswith("```json"):
        content = content[7:]
    if content.startswith("```"):
        content = content[3:]
    if content.endswith("```"):
        content = content[:-3]

    import json

    data = json.loads(content.strip())
    return pd.DataFrame(data)


# --- HÀM 3: TẠO FILE EXCEL CHUẨN ĐỊNH DẠNG (MÀU CAM, BẢNG VIỀN, TỰ XUỐNG DÒNG) ---
def create_styled_excel(df_bilingual):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bảng Song Ngữ"

    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    header_fill = PatternFill(
        start_color="E67E22", end_color="E67E22", fill_type="solid"
    )
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    cell_font = Font(name="Arial", size=10)
    align_center = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )

    for r_idx, row in enumerate(
        df_bilingual.itertuples(index=False), start=1
    ):
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_border
            cell.alignment = align_center

            if r_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
            else:
                cell.font = cell_font

    # Tự chỉnh độ rộng cột và độ cao dòng
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                lines = str(cell.value).split("\n")
                for line in lines:
                    if len(line) > max_len:
                        max_len = len(line)
        ws.column_dimensions[col_letter].width = max(max_len + 6, 14)

    for row in ws.iter_rows():
        ws.row_dimensions[row[0].row].height = 38

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# --- HÀM 4: HIỂN THỊ BẢNG HTML XEM TRƯỚC SỐNG ĐỘNG TRÊN STREAMLIT ---
def display_html_table(df):
    html = '<table style="width:100%; border-collapse: collapse; text-align: center; font-family: Arial;">'
    for r_idx, row in enumerate(df.values):
        html += "<tr>"
        for val in row:
            val_str = (
                str(val).replace("\n", "<br>")
                if pd.notna(val) and str(val) != "None"
                else ""
            )
            if r_idx == 0:
                html += f'<th style="border: 1px solid #333; padding: 10px; background-color: #E67E22; color: white; font-weight: bold;">{val_str}</th>'
            else:
                html += f'<td style="border: 1px solid #ccc; padding: 8px;">{val_str}</td>'
        html += "</tr>"
    html += "</table>"
    st.markdown(html, unsafe_allow_html=True)


# --- GIAO DIỆN STREAMLIT MAIN ---
uploaded_file = st.file_uploader(
    "Tải lên file Excel (.xlsx, .xls) hoặc File Ảnh (.png, .jpg, .jpeg)",
    type=["xlsx", "xls", "png", "jpg", "jpeg"],
)

if uploaded_file is not None:
    file_type = uploaded_file.name.split(".")[-1].lower()

    if st.button("🚀 Tiến Hành Dịch Bảng Song Ngữ"):
        if not gemini_api_key:
            st.error("Vui lòng nhập Gemini API Key ở thanh bên (Sidebar)!")
        else:
            with st.spinner("Đang phân tích bảng biểu và xử lý dịch..."):
                try:
                    df_result = None

                    if file_type in ["png", "jpg", "jpeg"]:
                        image = Image.open(uploaded_file)
                        st.image(
                            image,
                            caption="Ảnh tải lên",
                            use_container_width=True,
                        )
                        df_result = extract_table_from_image(
                            image, gemini_api_key, src_lang, target_lang
                        )

                    elif file_type in ["xlsx", "xls"]:
                        df_raw = pd.read_excel(uploaded_file, header=None)
                        df_result = process_excel_file(
                            df_raw, gemini_api_key, src_lang, target_lang
                        )

                    if df_result is not None:
                        st.success("✅ Đã xử lý xong!")
                        st.subheader("✨ Kết Quả Bảng Song Ngữ Xem Trước:")

                        # Hiển thị kết quả dạng bảng HTML 2 dòng chuẩn
                        display_html_table(df_result)

                        # Nút Tải File Excel
                        excel_bytes = create_styled_excel(df_result)
                        st.write("")
                        st.download_button(
                            label="📥 Tải Xuống File Excel Kết Quả",
                            data=excel_bytes,
                            file_name="Bang_Song_Ngu_Ket_Qua.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {e}")
