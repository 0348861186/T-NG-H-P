import copy
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, range_boundaries
from deep_translator import GoogleTranslator
import streamlit as st

st.set_page_config(
    page_title="App Dịch Excel Song Ngữ", page_icon="🌐", layout="wide"
)

st.title("🌐 Ứng dụng Dịch File Excel Song Ngữ (Giữ nguyên 100% định dạng)")
st.markdown("---")

st.sidebar.header("Cấu hình chiều dịch")
direction = st.sidebar.selectbox(
    "Chọn chế độ dịch:", ("Trung - Việt", "Việt - Trung")
)

# Sử dụng mã 'zh-CN' chuẩn xác theo danh sách hỗ trợ của deep_translator
if direction == "Trung - Việt":
    source_lang, target_lang = "zh-CN", "vi"
    st.sidebar.info("📌 Tiếng Trung (Dòng trên) -> Tiếng Việt (Dòng dưới).")
else:
    source_lang, target_lang = "vi", "zh-CN"
    st.sidebar.info("📌 Tiếng Việt (Dòng trên) -> Tiếng Trung (Dòng dưới).")

# Khởi tạo translator an toàn
try:
    translator = GoogleTranslator(source=source_lang, target=target_lang)
except Exception as e:
    st.sidebar.error(f"Lỗi khởi tạo bộ dịch: {e}")


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


uploaded_file = st.file_uploader(
    "Tải lên file Excel (.xlsx) của bạn", type=["xlsx"]
)

if uploaded_file is not None:
    st.success("Tải file thành công! Đang tiến hành xử lý...")

    try:
        wb = openpyxl.load_workbook(uploaded_file, data_only=False)

        with st.spinner("Đang xử lý dịch, chèn dòng và bảo toàn cấu trúc..."):
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]

                # 1. Bảo toàn chiều rộng cột
                col_widths = {
                    col: ws.column_dimensions[col].width
                    for col in ws.column_dimensions
                }

                max_row = ws.max_row
                max_col = ws.max_column

                # 2. Lưu chiều cao hàng gốc
                row_heights_bak = {}
                for r in range(1, max_row + 1):
                    if ws.row_dimensions[r].height:
                        row_heights_bak[r] = ws.row_dimensions[r].height

                # 3. Xử lý các vùng Merge Cells an toàn
                original_merged_ranges = list(ws.merged_cells.ranges)
                for cr in original_merged_ranges:
                    ws.unmerge_cells(str(cr))

                # 4. Duyệt từ dưới lên trên để chèn dòng dịch ngay phía dưới dòng gốc
                for r in range(max_row, 0, -1):
                    has_content = any(
                        ws.cell(row=r, column=c).value is not None
                        for c in range(1, max_col + 1)
                    )
                    if not has_content:
                        continue

                    # Chèn 1 dòng trống ngay bên dưới dòng r
                    ws.insert_rows(r + 1, amount=1)

                    if r in row_heights_bak:
                        ws.row_dimensions[r + 1].height = row_heights_bak[r]
                        ws.row_dimensions[r].height = row_heights_bak[r]

                    # Sao chép định dạng và dịch ô
                    for c in range(1, max_col + 1):
                        src_cell = ws.cell(row=r, column=c)
                        dst_cell = ws.cell(row=r + 1, column=c)

                        val = src_cell.value
                        dst_cell.value = val

                        # Sao chép style chi tiết
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

                        if val is not None:
                            translated_val = translate_text(val)
                            if direction == "Trung - Việt":
                                dst_cell.value = translated_val
                            else:
                                dst_cell.value = val
                                src_cell.value = translated_val

                # 5. Phục hồi và mở rộng Merge Cells theo tỷ lệ dòng mới
                for cr in original_merged_ranges:
                    min_col, min_row, max_col_merged, max_row_merged = (
                        range_boundaries(str(cr))
                    )
                    new_min_row = (min_row * 2) - 1
                    new_max_row = max_row_merged * 2
                    new_range_str = (
                        f"{get_column_letter(min_col)}{new_min_row}:"
                        f"{get_column_letter(max_col_merged)}{new_max_row}"
                    )
                    try:
                        ws.merge_cells(new_range_str)
                    except Exception:
                        pass

                # Khôi phục chiều rộng cột
                for col, width in col_widths.items():
                    if width:
                        ws.column_dimensions[col].width = width

        # Lưu file
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("🎉 Xử lý hoàn tất!")
        st.download_button(
            label="📥 Tải xuống file Excel song ngữ",
            data=output,
            file_name=f"dich_song_ngu_{uploaded_file.name}",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )

    except Exception as e:
        st.error(f"Đã xảy ra lỗi trong quá trình xử lý file: {e}")
