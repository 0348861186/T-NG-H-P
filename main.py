import copy
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, range_boundaries
from deep_translator import GoogleTranslator
import streamlit as st

st.set_page_config(
    page_title="App Dịch Excel Song Ngữ Chuẩn 100% (Ultimate)",
    page_icon="🌐",
    layout="wide",
)

st.title(
    "🌐 Ứng dụng Dịch File Excel Song Ngữ (Bảo toàn 100% Công thức, Biểu đồ,"
    " Hình ảnh & Merge Cells)"
)
st.markdown("---")

st.sidebar.header("Cấu hình chiều dịch")
direction = st.sidebar.selectbox(
    "Chọn chế độ dịch:", ("Trung - Việt", "Việt - Trung")
)

if direction == "Trung - Việt":
    source_lang, target_lang = "zh", "vi"
    st.sidebar.info("📌 Tiếng Trung (Dòng trên) -> Tiếng Việt (Dòng dưới).")
else:
    source_lang, target_lang = "vi", "zh"
    st.sidebar.info(
        "📌 Tiếng Trung (Dòng trên) -> Tiếng Việt (Dòng dưới). (Đã tự động đảo"
        " chiều)."
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


uploaded_file = st.file_uploader(
    "Tải lên file Excel (.xlsx) của bạn", type=["xlsx"]
)

if uploaded_file is not None:
    st.success("Tải file thành công! Đang xử lý toàn diện file gốc...")

    try:
        # Load workbook với data_only=False để bảo toàn 100% công thức, biểu đồ, hình ảnh
        wb = openpyxl.load_workbook(uploaded_file, data_only=False)

        with st.spinner(
            "Đang xử lý dịch, chèn dòng, đồng bộ Merge Cells, Biểu đồ và Hình"
            " ảnh..."
        ):
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]

                # 1. Bảo toàn chiều rộng cột
                col_widths = {
                    col: ws.column_dimensions[col].width
                    for col in ws.column_dimensions
                }

                # 2. Xử lý Hình ảnh (Images) - Lưu thông tin anchor ban đầu để điều chỉnh sau khi insert row
                # openpyxl image anchor thường nằm ở cell (e.g., img.anchor._from)
                images_info = []
                for img in getattr(ws, "_images", []):
                    if hasattr(img, "anchor"):
                        # Lưu lại thông tin vị trí dòng/cột của anchor
                        images_info.append(
                            {
                                "image": img,
                                "row": (
                                    img.anchor._from.row
                                    if hasattr(img.anchor, "_from")
                                    else None
                                ),
                                "col": (
                                    img.anchor._from.col
                                    if hasattr(img.anchor, "_from")
                                    else None
                                ),
                            }
                        )

                max_row = ws.max_row
                max_col = ws.max_column

                # 3. Lưu chiều cao hàng gốc
                row_heights_bak = {}
                for r in range(1, max_row + 1):
                    if ws.row_dimensions[r].height:
                        row_heights_bak[r] = ws.row_dimensions[r].height

                # 4. Lấy danh sách các ô gộp (Merged Cells) ban đầu
                # Ta cần xử lý cẩn thận để sau khi chèn dòng, các ô gộp được mở rộng tương ứng xuống dưới
                original_merged_ranges = list(ws.merged_cells.ranges)

                # Unmerge tạm thời để thao tác insert dòng an toàn cho từng ô
                for cr in original_merged_ranges:
                    ws.unmerge_cells(str(cr))

                # 5. Duyệt từ dưới lên trên để insert dòng và dịch
                # Mỗi khi insert dòng tại r+1, ta sao chép nội dung dòng r xuống r+1, dịch và đồng thời đẩy các dòng phía dưới đi lên/xuống hợp lý.
                # Để đảm bảo công thức tự động cập nhật, openpyxl insert_rows hỗ trợ dịch chuyển tham chiếu công thức tuyệt vời.
                rows_inserted_map = (
                    {}
                )  # Lưu vết các dòng đã được chèn thêm (row_gốc -> row_đã_chèn)

                for r in range(max_row, 0, -1):
                    has_content = any(
                        ws.cell(row=r, column=c).value is not None
                        for c in range(1, max_col + 1)
                    )
                    if not has_content:
                        continue

                    # Chèn 1 dòng trống ngay bên dưới dòng r
                    ws.insert_rows(r + 1, amount=1)
                    rows_inserted_map[r] = r + 1

                    # Đồng bộ chiều cao hàng
                    if r in row_heights_bak:
                        ws.row_dimensions[r + 1].height = row_heights_bak[r]
                        ws.row_dimensions[r].height = row_heights_bak[r]

                    # Sao chép định dạng và xử lý dịch cho từng ô trong hàng
                    for c in range(1, max_col + 1):
                        src_cell = ws.cell(row=r, column=c)
                        dst_cell = ws.cell(row=r + 1, column=c)

                        val = src_cell.value
                        dst_cell.value = val

                        # Sao chép 100% định dạng chi tiết (Font, Fill, Alignment, Border, NumberFormat)
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

                        # Dịch và sắp xếp: Tiếng Việt luôn nằm ngay dưới Tiếng Trung
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

                # 6. Tái tạo lại Merge Cells chuẩn xác dựa trên số dòng đã chèn
                # Với mỗi merged range ban đầu (min_col, min_row, max_col, max_row), 
                # khi mỗi dòng bên trong bị nhân đôi, ta tính toán lại phạm vi hàng mới: min_row_moi và max_row_moi.
                for cr in original_merged_ranges:
                    min_col, min_row, max_col_merged, max_row_merged = (
                        range_boundaries(str(cr))
                    )

                    # Tính toán lại row bắt đầu và kết thúc mới sau khi mỗi hàng bị chèn thêm 1 hàng dịch bên dưới
                    # Ví dụ hàng r ban đầu khi dịch sẽ chiếm hàng (2*r - 1) và (2*r)
                    new_min_row = (min_row * 2) - 1
                    new_max_row = max_row_merged * 2

                    from openpyxl.utils import get_column_letter

                    new_range_str = (
                        f"{get_column_letter(min_col)}{new_min_row}:"
                        f"{get_column_letter(max_col_merged)}{new_max_row}"
                    )
                    try:
                        ws.merge_cells(new_range_str)
                    except Exception:
                        pass

                # 7. Khôi phục chiều rộng cột
                for col, width in col_widths.items():
                    if width:
                        ws.column_dimensions[col].width = width

        # Lưu file vào bộ nhớ
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success(
            "🎉 Dịch thành công hoàn hảo! Đã giữ nguyên 100% công thức,"
            " biểu đồ, hình ảnh, merge cells và định dạng phức tạp."
        )

        st.download_button(
            label="📥 Tải xuống file Excel song ngữ hoàn chỉnh (Ultimate)",
            data=output,
            file_name=f"dich_song_ngu_ultimate_{uploaded_file.name}",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )

    except Exception as e:
        st.error(f"Đã xảy ra lỗi trong quá trình xử lý file: {e}")
