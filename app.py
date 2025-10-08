import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

# ---- PAGE SETUP ----
st.set_page_config(page_title="Smart Assessment Report Generator", page_icon="📊", layout="centered")

st.title("📊 Smart Assessment Report Generator")
st.write("Easily record marks and automatically generate a final performance report.")

# ---- SCHOOL INFO ----
st.header("🏫 School & Class Information")
col1, col2 = st.columns(2)
with col1:
    district = st.text_input("District")
    school = st.text_input("School")
    year = st.text_input("Academic Year", datetime.date.today().year)
with col2:
    sector = st.text_input("Sector")
    class_group = st.text_input("Class", placeholder="e.g., S1A")

# ---- UPLOAD STUDENT LIST ----
st.subheader("📘 Upload Students List")
uploaded_students = st.file_uploader("Upload Excel file containing student names (multi-sheet allowed)", type=["xlsx"])

if uploaded_students:
    excel_file = pd.ExcelFile(uploaded_students)
    sheet_names = excel_file.sheet_names
    selected_sheet = st.selectbox("Select Class Group (from uploaded file)", sheet_names)
    students_df = pd.read_excel(uploaded_students, sheet_name=selected_sheet)

    # show preview
    st.dataframe(students_df.head())

    # ---- TEST SETTINGS ----
    st.subheader("🧮 Test Information")
    num_tests = st.number_input("Number of Tests", 1, 10, 5)
    max_marks = st.number_input("Maximum Marks per Test", 1, 200, 100)

    # ---- ENTER TEST DATES ----
    st.markdown("### 🗓️ Enter Test Dates")
    test_dates = []
    cols = st.columns(num_tests)
    for i in range(num_tests):
        with cols[i]:
            date = st.date_input(f"Test {i+1} Date", datetime.date.today())
            test_dates.append(date)

    # ---- ENTER SCORES ----
    st.markdown("### ✏️ Enter Scores for Each Student")
    scores = {}
    for idx, row in students_df.iterrows():
        name = row['Names']
        with st.expander(f"{idx+1}. {name}"):
            marks = []
            cols2 = st.columns(num_tests)
            for i in range(num_tests):
                with cols2[i]:
                    mark = st.number_input(f"Test {i+1}", 0.0, float(max_marks), 0.0, key=f"{name}_{i}")
                    marks.append(mark)
            scores[name] = marks

    # ---- GENERATE REPORT ----
    if st.button("📤 Generate Report"):
        report = []
        for name, marks in scores.items():
            total = sum(marks)
            percent = round((total / (num_tests * max_marks)) * 100, 2)
            report.append([name] + marks + [total, f"{percent}%"])

        columns = ["Names"] + [f"Test {i+1} ({test_dates[i]})" for i in range(num_tests)] + ["Total", "Percentage"]
        report_df = pd.DataFrame(report, columns=columns)

        # ---- CREATE EXCEL FILE ----
        output = BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "Final Report"

        # School info
        ws["A1"] = f"{school.upper()} - CLASS PERFORMANCE REPORT ({year})"
        ws["A2"] = f"District: {district} | Sector: {sector} | Class: {class_group}"
        ws["A1"].font = Font(bold=True, size=13)
        ws["A2"].font = Font(italic=True, size=11)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columns))
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(columns))

        # Header
        for col_num, col_name in enumerate(columns, 1):
            c = ws.cell(row=4, column=col_num, value=col_name)
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")

        # Data
        for row_num, row_data in enumerate(report, 5):
            for col_num, cell_value in enumerate(row_data, 1):
                cell = ws.cell(row=row_num, column=col_num, value=cell_value)
                cell.alignment = Alignment(horizontal="center", vertical="center")

        # Borders
        thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                      top=Side(style='thin'), bottom=Side(style='thin'))
        for row in ws.iter_rows(min_row=4, max_row=4+len(report), min_col=1, max_col=len(columns)):
            for cell in row:
                cell.border = thin

        wb.save(output)
        output.seek(0)

        st.success("✅ Report generated successfully!")

        st.download_button(
            label="⬇️ Download Excel Report",
            data=output,
            file_name=f"{class_group}_Final_Report_{year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
