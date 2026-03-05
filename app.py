import streamlit as st
import pandas as pd
from datetime import datetime
import pdfplumber, re, io, zipfile
from pypdf import PdfReader, PdfWriter
from openpyxl import Workbook

# -----------------------------
# PDF 테이블 추출
# -----------------------------
def extract_pdf_tables(pdf_buffer):

    clean_tables = []

    with pdfplumber.open(pdf_buffer) as pdf:

        for page in pdf.pages:

            table = page.extract_tables()[0]
            clean_table = []

            for row in table:

                clean_row = [
                    cell.strip()
                    for cell in row
                    if cell and cell.strip()
                ]

                clean_table.append(clean_row)

            clean_tables.append(clean_table)

    return clean_tables


# -----------------------------
# 세금 데이터 추출
# -----------------------------
def parse_tax_data(clean_tables):

    results = []

    for table in clean_tables:

        address = table[4][1]

        month = table[10][0].zfill(2)
        day = table[10][1].zfill(2)

        price = int(re.sub(r"[^\d]", "", table[10][2]))
        tax = int(re.sub(r"[^\d]", "", table[10][3]))

        data = {
            "사업자주소": address,
            "회계일": f"2026-{month}-{day}",
            "공급가": price,
            "세액": tax
        }

        results.append(data)

    return results


# -----------------------------
# 엑셀 생성 (안정 버전)
# -----------------------------
def write_excel(tax_data):

    wb = Workbook()
    ws = wb.active

    ws.append(["사업자 주소", "회계일", "공급가", "세액"])

    for data in tax_data:

        ws.append([
            data["사업자주소"],
            data["회계일"],
            data["공급가"],
            data["세액"]
        ])

    for row in ws.iter_rows(min_row=2, min_col=3, max_col=4):
        for cell in row:
            cell.number_format = '#,##0'

    buffer = io.BytesIO()
    wb.save(buffer)

    buffer.seek(0)

    return buffer


# -----------------------------
# PDF 분할
# -----------------------------
def split_pdf_by_address(pdf_buffer, clean_tables):

    pdf_buffer.seek(0)

    reader = PdfReader(pdf_buffer)

    pdf_files = []

    for i, page in enumerate(reader.pages):

        writer = PdfWriter()
        writer.add_page(page)

        buffer = io.BytesIO()
        writer.write(buffer)

        file_name = f"세금계산서_{clean_tables[i][4][1]}.pdf"

        buffer.seek(0)

        pdf_files.append((file_name, buffer))

    return pdf_files


# -----------------------------
# ZIP 생성
# -----------------------------
def create_zip(pdf_files):

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w") as z:

        for file_name, buffer in pdf_files:
            z.writestr(file_name, buffer.read())

    zip_buffer.seek(0)

    return zip_buffer


# -----------------------------
# 메인
# -----------------------------
def main():

    st.title("앱 대시보드")

    menu = ["Pdf 업로드", "About"]
    choice = st.sidebar.selectbox("메뉴", menu)

    if choice == "Pdf 업로드":

        st.subheader("Pdf 파일 업로드")

        pdf_file = st.file_uploader("Pdf를 업로드 하세요.", type=["pdf"])

        if pdf_file:

            if "last_file" not in st.session_state:
                st.session_state.last_file = None

            if pdf_file.name != st.session_state.last_file:

                pdf_buffer = io.BytesIO(pdf_file.getbuffer())
                st.session_state.pdf_buffer = pdf_buffer

                pdf_buffer.seek(0)

                with st.spinner("PDF 분석 중..."):
                    st.session_state.clean_tables = extract_pdf_tables(pdf_buffer)

                st.session_state.last_file = pdf_file.name

            clean_tables = st.session_state.clean_tables
            pdf_buffer = st.session_state.pdf_buffer

            st.success("PDF 업로드 완료")

            col1, col2 = st.columns(2)

            # ---------------- PDF 분할 ----------------
            with col1:

                if st.button("PDF 페이지 분할"):

                    with st.spinner("PDF 분할 중..."):

                        pdf_files = split_pdf_by_address(pdf_buffer, clean_tables)
                        zip_buffer = create_zip(pdf_files)

                    st.download_button(
                        label="PDF 다운로드 (ZIP)",
                        data=zip_buffer,
                        file_name="세금계산서.zip",
                        mime="application/zip"
                    )

                    st.success("PDF 분할 완료")

            # ---------------- 엑셀 추출 ----------------
            with col2:

                if st.button("엑셀로 추출"):

                    with st.spinner("엑셀 추출 중..."):

                        tax_data = parse_tax_data(clean_tables)
                        buffer = write_excel(tax_data)

                    st.download_button(
                        label="엑셀 다운로드",
                        data=buffer.getvalue(),
                        file_name="세금내역.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.write(tax_data)
                    st.success("엑셀 작성 완료")

    else:

        st.subheader("이 대시보드 설명")

        st.write(
            "달릴라는 동글이, 판다, 물범, 리트리버, 치와와, 고양이, 토끼래요."
            " 그니까 조금만 이해해주세요 ㅠㅠ"
        )


if __name__ == "__main__":
    main()