from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

out = "/mnt/data/Bang_cham_cong_2026-08-26_Trung_Viet.xlsx"

wb = Workbook()
ws = wb.active
ws.title = "Bảng song ngữ"

# --- Data translated Chinese -> Vietnamese ---
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
    (15, ("临时工", "Công nhân thời vụ"), "MERGE_4", "", "", ""),
    (15, ("新临时工", "Công nhân thời vụ mới"), "MERGE_2", "", "", ""),
]

# --- Layout ---
ws.merge_cells("A1:F1")
ws["A1"] = f"{title_cn}\n{title_vi}"
ws.row_dimensions[1].height = 42

# Header row
for col, (cn, vi) in enumerate(headers, 1):
    cell = ws.cell(row=2, column=col)
    cell.value = cn if cn == vi else f"{cn}\n{vi}"
    cell.font = Font(name="Microsoft YaHei", size=11, bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.fill = PatternFill("solid", fgColor="ED7D00")

ws.row_dimensions[2].height = 36

# Body rows
start_row = 3
for i, (stt, dept, machines, formal, temp, remark) in enumerate(rows, start_row):
    ws.cell(i, 1, stt)
    ws.cell(i, 2, f"{dept[0]}\n{dept[1]}")
    if machines != "MERGE_4" and machines != "MERGE_2":
        ws.cell(i, 3, machines)
        ws.cell(i, 4, formal)
        ws.cell(i, 5, temp)
    ws.cell(i, 6, f"{remark[0]}\n{remark[1]}" if isinstance(remark, tuple) else remark)

    for c in range(1, 7):
        ws.cell(i, c).font = Font(name="Microsoft YaHei", size=10)
        ws.cell(i, c).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[i].height = 32

# Merge the C:E cells for the two temporary-worker rows, matching the image
ws.merge_cells(start_row=17, start_column=3, end_row=17, end_column=5)
ws["C17"] = 4
ws["C17"].alignment = Alignment(horizontal="center", vertical="center")
ws.merge_cells(start_row=18, start_column=3, end_row=18, end_column=5)
ws["C18"] = 2
ws["C18"].alignment = Alignment(horizontal="center", vertical="center")

# Remark "套袋" is vertically centered across rows 15-16
ws.merge_cells("F17:F18")
ws["F17"] = "套袋\nĐóng túi"
ws["F17"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Total row
ws.merge_cells("A19:B19")
ws["A19"] = "一共\nTổng cộng"
ws.merge_cells("C19:E19")
ws["C19"] = 42
ws["A19"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws["C19"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[19].height = 34

# Borders for all cells, including merged areas
thin = Side(style="thin", color="000000")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

for row in ws.iter_rows(min_row=1, max_row=19, min_col=1, max_col=6):
    for cell in row:
        cell.border = border

# Reapply border to merged-region anchors and edges
for rng in ["A1:F1", "A19:B19", "C19:E19", "C17:E17", "C18:E18", "F17:F18"]:
    # openpyxl preserves the merged structure; anchor formatting is enough for content,
    # and the surrounding cells already carry borders.
    pass

# Column widths approximating the source image
widths = {"A": 8, "B": 24, "C": 14, "D": 16, "E": 16, "F": 18}
for col, width in widths.items():
    ws.column_dimensions[col].width = width

# Title styling
ws["A1"].font = Font(name="Microsoft YaHei", size=13, bold=True)
ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Print / page setup
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A3"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_margins = PageMargins(left=0.2, right=0.2, top=0.3, bottom=0.3, header=0.1, footer=0.1)

# Make the total visually bold
for cell_ref in ["A19", "C19"]:
    ws[cell_ref].font = Font(name="Microsoft YaHei", size=11, bold=True)

wb.save(out)

print(f"Đã tạo file Excel song ngữ: {out}")
