import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date

st.set_page_config(page_title="Student Marks Recorder", layout="wide")

st.title("📘 Student Marks Recorder")

st.write("Enter marks for each student, then download the updated workbook with test details and totals.")

# Upload Excel workbook
uploaded_file = st.file_uploader("📂 Upload the class Excel workbook", type=["xlsx"])

if uploaded_file:
    workbook = pd.ExcelFile(uploaded_file)
    
    # Let user choose the class (sheet)
    sheet_name = st.selectbox("Select the class (sheet)", workbook.sheet_names)
    
    # 👇 Parse the sheet that the user selected (not default one)
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

    if "Name" not in df.columns:
        st.error("❌ 'Name' column not found in your Excel file. Please ensure one column has student names.")
    else:
        # User inputs
        district = st.text_input("District")
        sector = st.text_input("Sector")
        school = st.text_input("School")
        academic_year = st.text_input("Academic Year (e.g. 2025-2026)")
        term = st.selectbox("Term", ["Term I", "Term II", "Term III"])
        test_count = st.number_input("Number of tests", min_value=1, max_value=10, step=1)
        max_mark = st.number_input("Maximum marks for each test", min_value=1, max_value=100, value=40)
        test_date = st.date_input("Date of the test", value=date.today())

        st.markdown("### ✏️ Enter Scores for Each Student")
        names = df["Name"].tolist()
        scores = {}

        for name in names:
            scores[name] = []
            for i in range(test_count):
                score = st.number_input(f"{name} - Test {i+1}", min_value=0, max_value=int(max_mark), key=f"{sheet_name}_{name}_{i}")
                scores[name].append(score)

        if st.button("💾 Generate Excel File"):
            new_df = df.copy()

            # Add test columns
            for i in range(test_count):
                new_df[f"Test {i+1}"] = [scores[n][i] for n in names]

            # Add total and percentage
            new_df["Total"] = new_df[[f"Test {i+1}" for i in range(test_count)]].sum(axis=1)
            new_df["% (out of 100)"] = (new_df["Total"] / (test_count * max_mark)) * 100

            # Add headers (date, max marks, etc.)
            header_info = pd.DataFrame({
                "": [
                    f"District: {district}",
                    f"Sector: {sector}",
                    f"School: {school}",
                    f"Class: {sheet_name}",
                    f"Academic Year: {academic_year}",
                    f"Term: {term}",
                    f"Test Date: {test_date}",
                    f"Maximum Marks per Test: {max_mark}"
                ]
            })

            # Combine info + table
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                header_info.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=0)
                new_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=len(header_info) + 2)
            output.seek(0)

            st.success(f"✅ Marks file for class {sheet_name} generated successfully!")
            st.download_button(
                label="⬇️ Download Updated Workbook",
                data=output,
                file_name=f"{sheet_name}_Marks_Record.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
