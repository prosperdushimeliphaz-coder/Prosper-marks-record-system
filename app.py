import streamlit as st
import pandas as pd
from datetime import datetime
import io
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from fpdf import FPDF

# --------------------------
# PAGE CONFIG
# --------------------------
st.set_page_config(page_title="Student Marks Manager", layout="centered")
st.title("📘 Student Marks Record App")

# --------------------------
# USER INPUT SECTION
# --------------------------
st.subheader("Enter School & Class Information")

col1, col2 = st.columns(2)
with col1:
    school_name = st.text_input("School Name", "Example Secondary School")
    district = st.text_input("District", "Kigali")
with col2:
    subject = st.text_input("Subject", "Biology")
    teacher = st.text_input("Teacher Name", "Mr. Prosper")

# Dropdown of available classes
classes = ["S1A", "S1B", "S2A", "S2B", "S3A", "S3B"]
selected_class = st.selectbox("Select Class", classes)

num_students = st.number_input("Number of Students", min_value=1, step=1, value=5)

# --------------------------
# MULTIPLE TEST ENTRY
# --------------------------
st.subheader("🧮 Enter Test Details")

num_tests = st.number_input("Number of Tests", min_value=1, step=1, value=1)
tests = []

for t in range(num_tests):
    st.markdown(f"### Test {t+1}")
    test_name = st.text_input(f"Name for Test {t+1}", f"Test {t+1}")
    test_date = st.date_input(f"Date for {test_name}", datetime.today())
    max_marks = st.number_input(f"Maximum Marks for {test_name}", min_value=1, max_value=100, value=100, key=f"max_{t}")
    marks = []
    for i in range(1, num_students + 1):
        marks.append(st.number_input(f"Student {i} - {test_name}", min_value=0, max_value=max_marks, step=1, value=0, key=f"{t}_{i}"))
    tests.append({
        "name": test_name,
        "date": test_date,
        "max": max_marks,
        "marks": marks
    })

# --------------------------
# SAVE ALL TESTS TO EXCEL
# --------------------------
if st.button("💾 Save / Update Marks"):
    wb = Workbook()
    ws = wb.active
    ws.title = selected_class

    # Headers
    ws.merge_cells('A1:D1')
    ws['A1'] = "MARKS RECORD"
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal="center")

    ws.merge_cells('A2:D2')
    ws['A2'] = f"{school_name} | {district} | Class: {selected_class} | Subject: {subject} | Teacher: {teacher}"
    ws['A2'].alignment = Alignment(horizontal="center")

    start_row = 4
    for t, test in enumerate(tests):
        ws.cell(row=start_row, column=1, value=f"{test['name']} - {test['date'].strftime('%d/%m/%Y')} - Max: {test['max']}")
        ws.cell(row=start_row, column=1).font = Font(bold=True)
        start_row += 1

        df = pd.DataFrame({
            "Student No.": range(1, num_students + 1),
            "Marks": test["marks"]
        })
        df["Rank"] = df["Marks"].rank(ascending=False, method='min').astype(int)

        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        start_row = ws.max_row + 2

    # Adjust width
    for col in ws.columns:
        max_len = max(len(str(cell.value)) for cell in col if cell.value)
        ws.column_dimensions[col[0].column_letter].width = max_len + 3

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    st.success("✅ Marks and test dates saved successfully!")

    st.download_button(
        label="⬇️ Download Excel Report",
        data=buffer,
        file_name=f"{selected_class}_{subject}_AllTests.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# --------------------------
# PDF REPORT (All Tests)
# --------------------------
if st.button("📄 Generate PDF Report"):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "MARKS RECORD", 0, 1, "C")

    pdf.set_font("Arial", size=11)
    pdf.multi_cell(0, 8, f"{school_name}\nDistrict: {district}\nClass: {selected_class}\nSubject: {subject}\nTeacher: {teacher}", 0, "L")
    pdf.ln(5)

    for t, test in enumerate(tests):
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 8, f"{test['name']} — {test['date'].strftime('%d/%m/%Y')} — Max: {test['max']}", 0, 1)
        pdf.set_font("Arial", "B", 11)
        pdf.cell(40, 10, "Student No.", 1, 0, "C")
        pdf.cell(40, 10, "Marks", 1, 0, "C")
        pdf.cell(40, 10, "Rank", 1, 1, "C")

        pdf.set_font("Arial", size=11)
        df = pd.DataFrame({
            "Student No.": range(1, num_students + 1),
            "Marks": test["marks"]
        })
        df["Rank"] = df["Marks"].rank(ascending=False, method='min').astype(int)
        for i, row in df.iterrows():
            pdf.cell(40, 10, str(row["Student No."]), 1, 0, "C")
            pdf.cell(40, 10, str(row["Marks"]), 1, 0, "C")
            pdf.cell(40, 10, str(row["Rank"]), 1, 1, "C")
        pdf.ln(5)

    pdf_buffer = io.BytesIO()
    pdf.output(pdf_buffer)
    pdf_buffer.seek(0)

    st.download_button(
        label="⬇️ Download PDF Report",
        data=pdf_buffer,
        file_name=f"{selected_class}_{subject}_AllTests.pdf",
        mime="application/pdf"
    )
