import io
import os
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins
from google import genai

# ============================================================
# CẤU HÌNH STREAMLIT
# ============================================================
st.set_page_config(
    page_title="Bảng chấm công Trung - Việt & Trợ lý AI",
    page_icon="📊",
    layout="centered"
)

# ============================================================
# TIÊU ĐỀ
# ============================================================
st.title("📊 Bảng chấm công Trung - Việt")
st.caption("2026年8月26日员工上班 / Nhân viên đi làm ngày 26/08/2026")

# ============================================================
# DỮ LIỆU BẢNG
# ============================================================
title_cn = "2026 年 8 月 26 日员工上班"
title_vi = "Nhân viên đi làm ngày 26/08/2026"

headers = [
    ("STT", "STT"),
    ("部门", "Bộ phận"),
    ("开几台机", "Số máy mở"),
    ("正式工", "Công nhân chính thức"),
    ("临时工", "Công nhân thời vụ"),
    ("备注", "Ghi chú"),
]

rows = [
    (1, ("连机", "Máy liên kết"), 5, 3, 2, ""),
    (2, ("制袋机", "Máy làm túi"), 6, 3, 2, ""),
    (3, ("连机吹膜", "Thổi màng liên máy"), 5, 4, "", ""),
    (4, ("制袋机吹膜", "Thổi màng máy làm túi"), 4, 2, 1, ""),
    (5, ("巡检", "Kiểm tra tuần tra"), "", 2, "", ""),
    (6, ("打扫", "Vệ sinh"), "", 1, "", ""),
    (7, ("打箱", "Đóng thùng"), "", 2, "", ""),
    (8, ("分口", "Chia miệng"), "", 1, 1, ""),
    (9, ("仓库+材料", "Kho + nguyên liệu"), "", 2, "", ""),
    (10, ("造粒", "Tạo hạt"), "", 3, 1, ""),
    (11, ("电工", "Thợ điện"), "", 2, "", ""),
    (12, ("办公室", "Văn phòng"), "", 4, "", ""),
    (13, ("QC", "QC"), "", 2, "", ""),
    (14, ("阿秋，阿勇", "A Qiu, A Yong"), "", 2, "", ""),
]

# ============================================================
# HÀM TẠO FILE EXCEL (Xử lý hoàn toàn trong RAM)
# ============================================================
def create_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Bảng song ngữ"

    font_name = "Microsoft YaHei"

    orange_fill = PatternFill(
        fill_type="solid",
        fgColor="ED7D00"
    )

    thin_side = Side(style="thin", color="000000")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # --- Tiêu đề ---
    ws.merge_cells("A1:F1")
    ws["A1"] = f"{title_cn}\n{title_vi}"
    ws["A1"].font = Font(name=font_name, size=13, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 42

    # --- Header ---
    for col, (cn, vi) in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col)
        cell.value = cn if cn == vi else f"{cn}\n{vi}"
        cell.font = Font(name=font_name, size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = orange_fill
        cell.border = border

    ws.row_dimensions[2].height = 38

    # --- Các dòng dữ liệu ---
    current_row = 3
    for stt, dept, machines, formal, temp, remark in rows:
        ws.cell(row=current_row, column=1, value=stt)
        ws.cell(row=current_row, column=2, value=f"{dept[0]}\n{dept[1]}")
        ws.cell(row=current_row, column=3, value=machines)
        ws.cell(row=current_row, column=4, value=formal)
        ws.cell(row=current_row, column=5, value=temp)
        ws.cell(row=current_row, column=6, value=remark)

        for col in range(1, 7):
            cell = ws.cell(row=current_row, column=col)
            cell.font = Font(name=font_name, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

        ws.row_dimensions[current_row].height = 32
        current_row += 1

    # --- 2 Dòng thời vụ bổ sung ---
    row_15_old = current_row
    row_15_new = current_row + 1

    # Dòng 15 - Công nhân thời vụ cũ
    ws.cell(row=row_15_old, column=1, value=15)
    ws.cell(row=row_15_old, column=2, value="临时工\nCông nhân thời vụ")
    ws.merge_cells(start_row=row_15_old, start_column=3, end_row=row_15_old, end_column=5)
    ws.cell(row=row_15_old, column=3, value=4)

    # Dòng 15 - Công nhân thời vụ mới
    ws.cell(row=row_15_new, column=1, value=15)
    ws.cell(row=row_15_new, column=2, value="新临时工\nCông nhân thời vụ mới")
    ws.merge_cells(start_row=row_15_new, start_column=3, end_row=row_15_new, end_column=5)
    ws.cell(row=row_15_new, column=3, value=2)

    # Ghi chú 套袋 / Đóng túi (Gộp ô cột 6)
    ws.merge_cells(start_row=row_15_old, start_column=6, end_row=row_15_new, end_column=6)
    ws.cell(row=row_15_old, column=6, value="套袋\nĐóng túi")

    for row in [row_15_old, row_15_new]:
        ws.row_dimensions[row].height = 36
        for col in range(1, 7):
            cell = ws.cell(row=row, column=col)
            cell.font = Font(name=font_name, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

    # --- Dòng Tổng cộng ---
    total_row = row_15_new + 1
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=2)
    ws.cell(row=total_row, column=1, value="一共\nTổng cộng")

    ws.merge_cells(start_row=total_row, start_column=3, end_row=total_row, end_column=5)
    ws.cell(row=total_row, column=3, value=42)

    for col in range(1, 7):
        cell = ws.cell(row=total_row, column=col)
        cell.font = Font(name=font_name, size=11, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    ws.row_dimensions[total_row].height = 36

    # --- Thiết lập độ rộng cột & trang in ---
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

    ws.page_margins = PageMargins(
        left=0.2, right=0.2, top=0.3, bottom=0.3, header=0.1, footer=0.1
    )

    # Ghi file vào bộ nhớ RAM
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ============================================================
# PREVIEW DỮ LIỆU
# ============================================================
st.subheader("📋 Nội dung bảng chấm công")

preview_data = []
for stt, dept, machines, formal, temp, remark in rows:
    preview_data.append({
        "STT": stt,
        "部门 / Bộ phận": f"{dept[0]} / {dept[1]}",
        "开几台机 / Số máy": machines,
        "正式工 / Chính thức": formal,
        "临时工 / Thời vụ": temp,
    })

preview_data.extend([
    {
        "STT": 15,
        "部门 / Bộ phận": "临时工 / Công nhân thời vụ",
        "开几台机 / Số máy": "",
        "正式工 / Chính thức": "",
        "临时工 / Thời vụ": 4,
    },
    {
        "STT": 15,
        "部门 / Bộ phận": "新临时工 / Công nhân thời vụ mới",
        "开几台机 / Số máy": "",
        "正式工 / Chính thức": "",
        "临时工 / Thời vụ": 2,
    }
])

st.dataframe(preview_data, use_container_width=True, hide_index=True)

# ============================================================
# NÚT XUẤT FILE EXCEL
# ============================================================
st.divider()
st.subheader("📥 Xuất Excel")

excel_file = create_excel()
st.download_button(
    label="⬇️ 下载 Excel / Tải Excel",
    data=excel_file.getvalue(),
    file_name="Bang_cham_cong_2026-08-26_Trung_Viet.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)

st.caption("Excel được tạo trực tiếp trong bộ nhớ RAM, phù hợp với Streamlit Cloud.")

# ============================================================
# TÍCH HỢP TRỢ LÝ AI GEMINI (Google GenAI SDK Mới)
# ============================================================
st.divider()
st.subheader("🤖 Trợ lý AI Gemini")

api_key = st.text_input("Nhập GEMINI_API_KEY của bạn:", type="password")

user_prompt = st.text_area("Nhập câu hỏi hoặc yêu cầu dịch thuật cho AI:")

if st.button("Gửi tới Gemini", use_container_width=True):
    if not api_key:
        st.error("Vui lòng nhập API Key để sử dụng Gemini.")
    elif not user_prompt.strip():
        st.warning("Vui lòng nhập nội dung câu hỏi.")
    else:
        try:
            # Khởi tạo client dùng thư viện google-genai mới
            client = genai.Client(api_key=api_key)
            
            # Sử dụng mô hình gemini-2.5-flash mới nhất
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=user_prompt,
            )
            
            st.success("Phản hồi từ Gemini:")
            st.write(response.text)
        except Exception as e:
            st.error(f"Đã xảy ra lỗi khi gọi Gemini API: {e}")
