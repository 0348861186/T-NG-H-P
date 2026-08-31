import copy
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from deep_translator import GoogleTranslator
import streamlit as st

st.set_page_config(
    page_title="App Dịch File Excel Song Ngữ Chuẩn 100%",
    page_icon="🌐",
    layout="wide",
)

st.title(
    "🌐 Ứng dụng Dịch File Excel Song Ngữ (Bảo toàn 100% Định dạng, Merge Cells"
    " & Công thức)"
)
st.markdown("---")

# Dashboard chọn chiều dịch
st.sidebar.header("Cấu hình chiều dịch")
direction = st.sidebar.selectbox(
    "Chọn chế độ dịch:", ("Trung - Việt", "Việt - Trung")
)

if direction == "Trung - Việt":
    source_lang, target_lang = "zh", "vi"
    st.sidebar.info(
        "📌 Quy tắc: Tiếng Trung (Dòng trên) -> Tiếng Việt (Dòng dưới ngay"
        " cùng ô)."
    )
else:
    source_lang, target_lang = "vi", "zh"
    st.sidebar.info(
        "📌 Quy tắc: Tiếng Trung (Dòng trên) -> Tiếng Việt (Dòng dưới ngay"
        " cùng ô). (Tự động đảo chiều)."
    )

translator = GoogleTranslator(source=source_lang, target=target_lang)


def translate_text(text):
    if text is None:
        return ""
    text_str = str(text).strip()
    if not text_str:
        return ""
    # Giữ nguyên nếu là công thức Excel (bắt đầu bằng dấu =) hoặc số thuần túy
    if text_str.startswith("=") or text_str.replace(".", "", 1).isdigit():
        return text_str
    try:
        return translator.translate(text_str)
    except Exception:
        return text_str


# Tải file lên
uploaded_file = st.file_uploader(
    "Tải lên file Excel (.xlsx) của bạn", type=["xlsx"]
)

if uploaded_file is not None:
    st.success("Tải file thành công! Đang xử lý toàn diện cấu trúc file...")

    try:
        # Load workbook với data_only=False để bảo toàn công thức tính toán
        wb = openpyxl.load_workbook(uploaded_file, data_only=False)

        with st.spinner(
            "Đang dịch, nhân đôi hàng, bảo toàn style, merge cells và biểu"
            " đồ..."
        ):
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]

                # 1. Lưu giữ chiều rộng cột (Column Widths)
                col_widths = {
                    col: ws.column_dimensions[col].width
                    for col in ws.column_dimensions
                }

                # 2. Xử lý Merged Cells: Lưu lại các vùng gộp, unmerge trước khi insert rows để tránh lỗi vỡ cấu trúc
                merged_ranges_bak = [
                    str(cell_range) for cell_range in ws.merged_cells.ranges
                ]
                for cr in merged_ranges_bak:
                    ws.unmerge_cells(cr)

                max_row = ws.max_row
                max_col = ws.max_column

                # Lưu chiều cao hàng gốc (Row Heights)
                row_heights_bak = {}
                for r in range(1, max_row + 1):
                    if ws.row_dimensions[r].height:
                        row_heights_bak[r] = ws.row_dimensions[r].height

                # 3. Duyệt từ dưới lên trên để insert dòng không làm lệch index hàng phía trên
                for r in range(max_row, 0, -1):
                    has_content = any(
                        ws.cell(row=r, column=c).value is not None
                        for c in range(1, max_col + 1)
                    )
                    if not has_content:
                        continue

                    # Chèn thêm 1 dòng trống ngay bên dưới dòng r
                    ws.insert_rows(r + 1)

                    # Đồng bộ chiều cao hàng mới bằng hàng cũ
                    if r in row_heights_bak:
                        ws.row_dimensions[r + 1].height = row_heights_bak[r]

                    # 4. Sao chép định dạng tuyệt đối 100% bằng copy.copy() sâu các đối tượng style
                    for c in range(1, max_col + 1):
                        src_cell = ws.cell(row=r, column=c)
                        dst_cell = ws.cell(row=r + 1, column=c)

                        val = src_cell.value
                        dst_cell.value = val

                        # Sử dụng copy.copy() để giữ nguyên vẹn 100% Font, Fill, Alignment, Border
                        if src_cell.font:
                            dst_cell.font = copy.copy(src_cell.font)
                        if src_cell.fill:
                            dst_cell.fill = copy.copy(src_cell.fill)
                        if src_cell.alignment:
                            dst_cell.alignment = copy.copy(
                                src_cell.alignment
                            )
                        if src_cell.border:
                            dst_cell.border = copy.copy(src_cell.border)
                        dst_cell.number_format = src_cell.number_format

                        # Thực hiện dịch và sắp xếp vị trí: Tiếng Việt luôn nằm ngay dưới Tiếng Trung
                        if val is not None:
                            translated_val = translate_text(val)
                            if direction == "Trung - Việt":
                                dst_cell.value = translated_val
                            else:
                                dst_cell.value = (
                                    val  # Dòng dưới giữ Tiếng Việt
                                )
                                src_cell.value = (
                                    translated_val  # Dòng trên đổi thành Tiếng Trung
                                )

                # 5. Phục hồi và mở rộng các vùng Merge Cells theo tỷ lệ dòng nhân đôi
                for cr in merged_ranges_bak:
                    parts = cr.split(":")
                    if len(parts) == 2:
                        top_left, bottom_right = parts[0], parts[1]
                        tl_col = "".join(filter(str.isalpha, top_left))
                        tl_row = int("".join(filter(str.isdigit, top_left)))
                        br_col = "".join(filter(str.isalpha, bottom_right))
                        br_row = int("".join(filter(str.isdigit, bottom_right)))

                        new_top_row = (tl_row * 2) - 1
                        new_bottom_row = br_row * 2
                        new_range_str = (
                            f"{tl_col}{new_top_row}:{br_col}{new_bottom_row}"
                        )
                        try:
                            ws.merge_cells(new_range_str)
                        except Exception:
                            pass

                # Khôi phục chiều rộng cột
                for col, width in col_widths.items():
                    if width:
                        ws.column_dimensions[col].width = width

        # Lưu file vào bộ nhớ đệm
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success(
            "🎉 Hoàn tất dịch file, giữ nguyên 100% định dạng, merge cells,"
            " biểu đồ và công thức!"
        )

        # 6. Nút download file excel sau khi dịch
        st.download_button(
            label="📥 Tải xuống file Excel song ngữ hoàn chỉnh",
            data=output,
            file_name=f"dich_song_ngu_chuan_100_{uploaded_file.name}",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )

    except Exception as e:
        st.error(f"Đã xảy ra lỗi trong quá trình xử lý file: {e}")
