import io
import json
import datetime
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins
from google import genai
from google.genai import types

# ============================================================
# CẤU HÌNH STREAMLIT
# ============================================================
st.set_page_config(
    page_title="Hệ Thống Xuất Bảng Chấm Công Tự Động",
    page_icon="🤖",
    layout="wide"
)

st.title("🤖 Ứng Dụng Đọc & Xuất Bảng Chấm Công Song Ngữ (100% Dynamic)")
st.caption("Tải file/ảnh lên -> AI tự nhận diện 100% dữ liệu -> Xuất Excel chuẩn đẹp.")

# ============================================================
# 1. CẤU HÌNH API KEY & TẢI FILE
# ============================================================
col1, col2 = st.columns([1, 2])

with col1:
    api_key = st.text_input("Nhập GEMINI_API_KEY:", type="password")

with col2:
    uploaded_file = st.file_uploader("Tải lên ảnh hoặc file bảng chấm công:", type=["png", "jpg", "jpeg", "pdf"])

# Khởi tạo Session State chứa dữ liệu động hoàn toàn
if "parsed_data" not in st.session_state:
    st.session_state.parsed_data = {
        "title_cn": "员工上班",
        "title_vi": "Nhân viên đi làm",
        "date_str": datetime.date.today().strftime("%Y-%m-%d"),
        "rows": [] # Mặc định rỗng, không hardcode bất kỳ bộ phận nào!
    }

# ============================================================
# 2. XỬ LÝ AI ĐỌC DỮ LIỆU TỰ ĐỘNG TỪ FILE/ẢNH
# ============================================================
if st.button("🚀 AI Phân Tích & Tự Động Điền Dữ Liệu", use_container_width=True):
    if not api_key:
        st.error("Vui lòng nhập API Key!")
    elif uploaded_file is None:
        st.warning("Vui lòng tải lên ảnh/file để AI đọc.")
    else:
        try:
            client = genai.Client(api_key=api_key)
            file_bytes = uploaded_file.read()
            file_part = types.Part.from_bytes(data=file_bytes, mime_type=uploaded_file.type)

            prompt = """
            Hãy phân tích hình ảnh/file bảng chấm công này và trích xuất toàn bộ dữ liệu dưới dạng JSON thuần túy (không dùng markdown backticks).
            Cấu trúc JSON yêu cầu chính xác như sau:
            {
                "title_cn": "Tiêu đề tiếng Trung trích xuất được",
                "title_vi": "Tiêu đề tiếng Việt dịch tương ứng",
                "date_str": "YYYY-MM-DD (Ngày trong ảnh nếu có, không có thì để rỗng)",
                "rows": [
                    {
                        "stt": 1,
                        "dept_cn": "Tên bộ phận/công việc tiếng Trung",
                        "dept_vi": "Dịch tên bộ phận/công việc sang tiếng Việt",
                        "machines": 5,
                        "formal": 3,
                        "temp": 2,
                        "remark": "Ghi chú nếu có"
                    }
                ]
            }
            Lưu ý: 
            - Trả về đúng định dạng JSON.
            - machines, formal, temp nếu không có thì ghi 0.
            - Dịch chính xác nghĩa tiếng Việt cho các bộ phận.
            """

            with st.spinner("AI đang quét và dịch dữ liệu từ file..."):
                response = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=[file_part, prompt]
                )
                
                # Làm sạch và parse JSON
                clean_json_str = response.text.replace("```json", "").replace("```", "").strip()
                data = json.loads(clean_json_str)
                
                # Lưu vào Session State
                st.session_state.parsed_data = data
                st.success("✅ Đã quét và phân tích dữ liệu thành công!")

        except Exception as e:
            st.error(f"Lỗi khi xử lý bằng AI: {e}")

# ============================================================
# 3. HIỂN THỊ VÀ CHO PHÉP CHỈNH SỬA DỮ LIỆU ĐÃ TRÍCH XUẤT
# ============================================================
st.divider()
st.subheader("📝 Dữ liệu trích xuất (Có thể chỉnh sửa trước khi xuất Excel)")

col_a, col_b, col_c = st.columns([2, 2, 1])
with col_a:
    title_cn = st.text_input("Tiêu đề (Trung):", value=st.session_state.parsed_data.get("title_cn", ""))
with col_b:
    title_vi = st.text_input("Tiêu đề (Việt):", value=st.session_state.parsed_data.get("title_vi", ""))
with col_c:
    date_val = st.text_input("Ngày:", value=st.session_state.parsed_data.get("date_str", ""))

# Chuyển đổi dữ liệu rows sang DataFrame để hiển thị lên bảng
rows_data = st.session_state.parsed_data.get("rows", [])
if rows_data:
    df_input = pd.DataFrame(rows_data)
else:
    # Bảng trống hoàn toàn nếu chưa tải file
    df_input = pd.DataFrame(columns=["stt", "dept_cn", "dept_vi", "machines", "formal", "temp", "remark"])

edited_df = st.data_editor(
    df_input,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    column_config={
        "stt": "STT",
        "dept_cn": "Bộ phận (Trung)",
        "dept_vi": "Bộ phận (Việt)",
        "machines": "Số máy mở",
        "formal": "Chính thức",
        "temp": "Thời vụ",
        "remark": "Ghi chú"
    }
)

# ============================================================
# 4. HÀM ENGINE XUẤT EXCEL THUẦN TÚY (ZERO HARDCODE)
# ============================================================
def build_excel_file(df, t_cn, t_vi, dt_str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    font_name = "Microsoft YaHei"
    orange_fill = PatternFill(fill_type="solid", fgColor="ED7D00")
    thin_side = Side(style="thin", color="000000")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

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
        ("正式工", "Công nhân chính thức"),
        ("临时工", "Công nhân thời vụ"),
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

    # Render từng dòng từ DataFrame
    current_row = 3
    total_workers = 0

    for _, row in df.iterrows():
        stt = row.get("stt", "")
        d_cn = str(row.get("dept_cn", "")) if pd.notna(row.get("dept_cn")) else ""
        d_vi = str(row.get("dept_vi", "")) if pd.notna(row.get("dept_vi")) else ""
        mac = row.get("machines", "") if pd.notna(row.get("machines")) and row.get("machines") != 0 else ""
        fml = row.get("formal", "") if pd.notna(row.get("formal")) and row.get("formal") != 0 else ""
        tmp = row.get("temp", "") if pd.notna(row.get("temp")) and row.get("temp") != 0 else ""
        rmk = str(row.get("remark", "")) if pd.notna(row.get("remark")) else ""

        # Tính tổng
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
    ws.cell(row=total_row, column=3, value=int(total_workers) if total_workers.is_integer() else total_workers)

    for col in range(1, 7):
        c = ws.cell(row=total_row, column=col)
        c.font = Font(name=font_name, size=11, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    ws.row_dimensions[total_row].height = 36

    # Cấu hình căn chỉnh khổ in A4
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 17
    ws.column_dimensions["E"].width = 17
    ws.column_dimensions["F"].width = 18

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.2, right=0.2, top=0.3, bottom=0.3, header=0.1, footer=0.1)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# ============================================================
# 5. NÚT TẢI FILE EXCEL
# ============================================================
st.divider()
excel_file = build_excel_file(edited_df, title_cn, title_vi, date_val)

st.download_button(
    label="⬇️ Tải xuống File Excel (.xlsx)",
    data=excel_file.getvalue(),
    file_name=f"Bang_cham_cong_{date_val}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)
