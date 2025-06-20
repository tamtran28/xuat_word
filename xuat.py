import streamlit as st
import win32com.client as win32
import os
from datetime import datetime
import tempfile

st.set_page_config(page_title="Xuất nhiều bảng Excel sang 1 file Word", page_icon="📄")
st.title("📄 Xuất bảng từ các sheet chỉ định sang 1 file Word")

uploaded_file = st.file_uploader("🔽 Chọn file Excel", type=["xlsx", "xlsm"])
range_address = st.text_input("📌 Vùng bảng cố định (VD: A1:M103)", value="A1:M103")
save_folder = st.text_input("📁 Thư mục lưu Word", value=os.getcwd())

# 💡 Sheet bạn muốn xử lý (không bị trùng tên)
selected_sheets = [
    "TK_KPCS_BANG_01",
    "TK_KPCS_BANG_02",
    "TK_KPCS_BANG_04",
    "TK_KPCS_BANG_06"
]

if st.button("📤 Xuất sang Word"):
    if not uploaded_file:
        st.error("⚠️ Vui lòng tải lên file Excel.")
    elif not os.path.exists(save_folder):
        st.error("⚠️ Thư mục lưu không tồn tại.")
    else:
        try:
            # Lưu file Excel tạm
            temp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            temp_excel.write(uploaded_file.read())
            temp_excel.close()

            # Mở Word
            word = win32.gencache.EnsureDispatch("Word.Application")
            word.Visible = True
            doc = word.Documents.Add()

            # Mở Excel bằng COM để copy bảng
            excel = win32.gencache.EnsureDispatch("Excel.Application")
            excel.Visible = False
            wb = excel.Workbooks.Open(temp_excel.name)

            for sheet_name in selected_sheets:
                try:
                    ws = wb.Sheets(sheet_name)
                    ws.Range(range_address).Copy()

                    doc.Content.InsertAfter(f"📄 Sheet: {sheet_name}\n")
                    doc.Content.Paragraphs.Last.Range.PasteExcelTable(
                        LinkedToExcel=False, WordFormatting=True, RTF=False
                    )
                    doc.Content.InsertParagraphAfter()
                except Exception as sheet_err:
                    doc.Content.InsertAfter(f"⚠️ Không thể xử lý sheet {sheet_name}: {sheet_err}\n")

            wb.Close(False)
            excel.Quit()

            # Lưu file Word
            filename = f"Export_SpecificSheets_{datetime.now():%Y%m%d_%H%M%S}.docx"
            save_path = os.path.join(save_folder, filename)
            doc.SaveAs2(FileName=save_path, FileFormat=16)

            st.success(f"✅ Đã lưu file Word tại:\n{save_path}")
        except Exception as e:
            st.error(f"❌ Lỗi: {e}")
