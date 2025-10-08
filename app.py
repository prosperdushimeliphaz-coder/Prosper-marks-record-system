import streamlit as st
import pandas as pd
import io
from fpdf import FPDF
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import date
import os

st.set_page_config(page_title="Prosper Marks Recorder", layout="wide")

# Title
st.title("📘 Student Marks Management System")

# ---- School Info ----
st.subheader("School & Class Information")
district = st.text_input("District:")
school = st.text_input("School:")
_class = st.text_input("Class:")
teacher = st.text_input("Teacher:")
subject = st.text_input("Subject:")

uploaded_file = st.file_uploader("📂 Upload Workbook (.xlsx)", type=["xlsx"])

if uploaded_file:
    wb = load_workbook(uploaded_file)
    if _class in wb.sheetnames:
        sheet = wb[_class]
    else:
        sheet = wb.active
    data = sheet.values
    cols = next(data)
    df = pd.DataFrame(data, columns=cols)
else:
    st.warning("Please upload your class workbook to continue.")
    st.stop()

# ---- Test Input ----
st.subheader("Test Information")
test_name = st.text_input("Test Name (e.g. Test 1):", "Test 1")
test_date = st.date_input("Date for Test", date.today())
max_marks = st.number_input("Max Marks for This Test", min_value=1, max_value=100, value=30)

# ---- Enter Marks ----
st.subheader("Enter Marks for Each Student")
if "Name" not in df.columns:
    st.error("Workbook must contain a 'Name' column.")
    st.stop()

marks = []
for i, name in enumerate(df["Name"]):
    mark = st.number_input(f"{name}'s marks:", min_value=0, max_value=int(max_marks), key=f"mark_{i}")
    marks.append(mark)

# Add or update this test column
df[test_name] = marks

# ---- Save/Update Workbook ----
if st.button("💾 Save / Update Workbook (Add Test)"):

    # Reload workbook to ensure all previous tests are kept
    workbook_path = "Updated_Marks_Record.xlsx"
    if os.path.exists(workbook_path):
        existing_wb = load_workbook(workbook_path)
        if _class in existing_wb.sheetnames:
            ws = existing_wb[_class]
            existing_data = ws.values
            cols = next(existing_data)
            old_df = pd.DataFrame(existing_data, columns=cols)
            # Merge new column (test) with existing data
            if test_name not in old_df.columns:
                old_df[test_name] = marks
            df = old_df
        else:
            df = df
    else:
        df = df

    with pd.ExcelWriter(workbook_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=_class)
    st.success("✅ Workbook updated successfully with new test!")

# ---- Reload for Full Report ----
if os.path.exists("Updated_Marks_Record.xlsx"):
    wb = load_workbook("Updated_Marks_Record.xlsx")
    sheet = wb[_class]
    data = sheet.values
    cols = next(data)
    df = pd.DataFrame(data, columns=cols)

# ---- Calculate Totals and Ranks ----
numeric_cols = df.select_dtypes(include=['number']).columns
df["Total"] = df[numeric_cols].sum(axis=1)
df["Rank"] = df["Total"].rank(ascending=False, method="min").astype(int)

st.subheader("📊 Marks Record (All Tests)")
st.dataframe(df)

# ---- Export Excel ----
excel_buffer = io.BytesIO()
with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name="Marks_Record")
excel_data = excel_buffer.getvalue()

# ---- Export PDF ----
pdf = FPDF()
pdf.add_page()
pdf.set_font("Arial", "B", 16)
pdf.cell(200, 10, f"MARKS REPORT - {_class}", ln=True, align="C")
pdf.ln(4)

pdf.set_font("Arial", "", 12)
pdf.cell(200, 8, f"District: {district}", ln=True)
pdf.cell(200, 8, f"School: {school}", ln=True)
pdf.cell(200, 8, f"Teacher: {teacher}", ln=True)
pdf.cell(200, 8, f"Subject: {subject}", ln=True)
pdf.ln(6)

pdf.set_font("Arial", "B", 11)
pdf.set_fill_color(220, 220, 220)
for col in df.columns:
    pdf.cell(190/len(df.columns), 8, str(col), border=1, align="C", fill=True)
pdf.ln()

pdf.set_font("Arial", "", 10)
for _, row in df.iterrows():
    for col in df.columns:
        pdf.cell(190/len(df.columns), 7, str(row[col]), border=1, align="C")
    pdf.ln()

pdf.ln(6)
pdf.set_font("Arial", "B", 12)
pdf.cell(200, 8, "Summary:", ln=True)
pdf.set_font("Arial", "", 11)
pdf.cell(200, 8, f"Total Students: {len(df)}", ln=True)
pdf.cell(200, 8, f"Class Average: {df[numeric_cols].mean().mean():.2f}", ln=True)

pdf_output = pdf.output(dest='S').encode('latin-1')

col1, col2 = st.columns(2)
col1.download_button(
    "⬇️ Download Excel Report (All Tests)",
    data=excel_data,
    file_name=f"{_class}_Full_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
col2.download_button(
    "⬇️ Download PDF Report (All Tests)",
    data=pdf_output,
    file_name=f"{_class}_Full_Report.pdf",
    mime="application/pdf"
)
