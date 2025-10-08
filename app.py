import streamlit as st
import pandas as pd
import datetime
import io
from fpdf import FPDF

st.set_page_config(page_title="Student Marks Record System", layout="wide")

st.title("📘 Student Marks Record System")

# --- Input Section ---
st.header("School Information")

col1, col2, col3 = st.columns(3)
with col1:
    district = st.text_input("District")
    sector = st.text_input("Sector")
with col2:
    school = st.text_input("School")
    academic_year = st.text_input("Academic Year")
with col3:
    class_name = st.text_input("Class")
    term = st.text_input("Term")

subject = st.text_input("Subject")
teacher = st.text_input("Teacher’s Name")

st.divider()

# --- Marks Section ---
st.header("Marks Entry")

uploaded_file = st.file_uploader("Upload Excel file with student names", type=["xlsx"])
max_marks_subject = st.number_input("Maximum marks for subject", min_value=1, value=40)

num_tests = st.number_input("Number of Tests", min_value=1, max_value=10, value=3)

test_names, test_dates, test_maximums = [], [], []
for i in range(1, num_tests + 1):
    c1, c2, c3 = st.columns(3)
    with c1:
        test_name = st.text_input(f"Name of Test {i}", f"Test {i}")
    with c2:
        test_date = st.date_input(f"Date of {test_name}", datetime.date.today())
    with c3:
        test_max = st.number_input(f"Max marks for {test_name}", min_value=1, value=20)
    test_names.append(test_name)
    test_dates.append(test_date.strftime("%d-%m-%Y"))
    test_maximums.append(test_max)

if uploaded_file:
    df_names = pd.read_excel(uploaded_file)
    student_names = df_names.iloc[:, 0].tolist()
    marks_data = {}

    st.write("### Enter Marks for Each Test")

    for test in test_names:
        st.subheader(test)
        test_marks = []
        for student in student_names:
            mark = st.number_input(f"{student} - {test}", min_value=0, value=0, step=1, key=f"{student}_{test}")
            test_marks.append(mark)
        marks_data[test] = test_marks

    # Prepare DataFrame
    results_df = pd.DataFrame(marks_data)
    results_df.insert(0, "Names", student_names)

    # Add total, percent, rank
    results_df["Total"] = results_df[test_names].sum(axis=1)
    results_df["%"] = round((results_df["Total"] / sum(test_maximums)) * 100, 2)
    results_df["Rank"] = results_df["Total"].rank(ascending=False, method="min").astype(int)

    # --- Swap rows (Date row on top, Test names row below it) ---
    dates_row = pd.DataFrame([["Date"] + test_dates + ["", "", ""]], columns=results_df.columns)
    tests_row = pd.DataFrame([["Test"] + test_names + ["", "", ""]], columns=results_df.columns)
    results_df = pd.concat([dates_row, tests_row, results_df], ignore_index=True)

    # Display
    st.write("### Final Results")
    st.dataframe(results_df, use_container_width=True)

    # --- Excel Download ---
    def to_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # Write school info
            info = [
                ["District", district],
                ["Sector", sector],
                ["School", school],
                ["Class", class_name],
                ["Term", term],
                ["Academic Year", academic_year],
                ["Subject", subject],
                ["Teacher", teacher],
                ["Max Marks (Subject)", max_marks_subject]
            ]
            info_df = pd.DataFrame(info, columns=["Field", "Value"])
            info_df.to_excel(writer, sheet_name="Report", index=False, startrow=0)
            df.to_excel(writer, sheet_name="Report", index=False, startrow=len(info) + 2)
        processed_data = output.getvalue()
        return processed_data

    excel_data = to_excel(results_df)
    st.download_button("⬇️ Download Excel Report", data=excel_data,
                       file_name=f"{subject}_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # --- PDF Download ---
    def generate_pdf(df):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=11)

        # Header
        pdf.cell(200, 10, txt=f"STUDENT MARKS REPORT - {subject.upper()}", ln=True, align="C")
        pdf.ln(5)

        # School Info
        info_lines = [
            f"District: {district}",
            f"Sector: {sector}",
            f"School: {school}",
            f"Class: {class_name}",
            f"Term: {term}",
            f"Academic Year: {academic_year}",
            f"Teacher: {teacher}",
            f"Maximum Marks (Subject): {max_marks_subject}",
        ]
        for line in info_lines:
            pdf.cell(200, 8, txt=line, ln=True)
        pdf.ln(5)

        # Table
        col_width = pdf.w / (len(df.columns) + 1)
        for i, row in df.iterrows():
            for value in row:
                pdf.cell(col_width, 8, txt=str(value), border=0)
            pdf.ln(8)

        pdf_output = io.BytesIO()
        pdf.output(pdf_output)
        return pdf_output.getvalue()

    pdf_data = generate_pdf(results_df)
    st.download_button("⬇️ Download PDF Report", data=pdf_data, file_name=f"{subject}_Report.pdf", mime="application/pdf")
