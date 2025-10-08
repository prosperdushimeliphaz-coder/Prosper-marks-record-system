import streamlit as st
import pandas as pd
import io
from fpdf import FPDF
from openpyxl import load_workbook, Workbook
from datetime import date
import os

st.set_page_config(page_title="Prosper Marks Recorder", layout="wide")

# Title
st.title("📘 Student Marks Management System")

# ---- Student & Class Info ----
st.subheader("School & Class Information")
district = st.text_input("District:")
school = st.text_input("School:")
_class = st.text_input("Class:")
teacher = st.text_input("Teacher:")
subject = st.text_input("Subject:")

uploaded_file = st.file_uploader("📂 Upload Workbook (.xlsx)", type=["xlsx"])

if uploaded_file:
    wb = load_workbook(uploaded_file)
    class_list = wb.sheetnames
    selected_class = st.selectbox("Select Class Sheet:", class_list)
    sheet = wb[selected_class]
    data = sheet.values
    cols = next(data)
    df = pd.DataFrame(data, columns=cols)
else:
    st.warning("Please upload your class workbook to continue.")
    st.stop()

# ---- Test Input Section ----
st.subheader("Test Information")
test_name = st.text_input("Test Name (e.g. Test 1):", "Test 1")
test_date = st.date_input("Date for Test", date.today())
max_marks = st.number_input("Max Marks for This Test", min_value=1, max_value=100, value=30)

# ---- Enter Marks ----
st.subheader("Enter Marks for Each Student")
if "Name" in df.columns:
    marks = []
    for i, name in enumerate(df["Name"]):
        mark = st.number_input(f"{name}'s marks:", min_value=0, max_value=int(max_marks), key=f"mark_{i}")
        marks.append(mark)
    df[test_name] = marks
else:
    st.error("Workbook must contain a 'Name' column for students.")
    st.stop()

# ---- Calculate Totals and Ranks ----
numeric_cols = df.select_dtypes(include=['number']).columns
df["Total"] = df[numeric_cols].sum(axis=1)
df["Rank"] = df["Total"].rank(ascending=False, method="min").astype(int)

st.subheader("Marks Record Table")
st.dataframe(df)

# ---- Save/Update Workbook Button ----
st.subheader("💾 Save or Download Reports")

save_buffer = io.BytesIO()

def update_workbook(existing_file, new_data, sheet_name):
    if os.path.exists(existing_file):
        wb = load_workbook(existing_file)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for r in dataframe_to_rows(new_data, index=False, header=True):
                ws.append(r)
        else:
            ws = wb.create_sheet(sheet_name)
            for r in dataframe_to_rows(new_data, index=False, header=True):
                ws.append(r)
        wb.save(existing_file)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        for r in dataframe_to_rows(new_data, index=False, header=True):
            ws.append(r)
        wb.save(existing_file)

if st.button("💾 Save / Update Workbook (Add Test)"):
    with pd.ExcelWriter(save_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=_class)
    st.success("✅ Test results added to workbook successfully!")

# ---- Generate Excel for download ----
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name="Marks_Record")
excel_data = buffer.getvalue()

# ---- Generate PDF ----
pdf = FPDF()
pdf.add_page()

pdf.set_font("Arial", "B", 16)
pdf.cell(200, 10, f"MARKS REPORT - {_class}", ln=True, align="C")
pdf.ln(4)

pdf.set_font("Arial", "", 12)
pdf.cell(200, 8, f"District: {district}", ln=True, align="L")
pdf.cell(200, 8, f"School: {school}", ln=True, align="L")
pdf.cell(200, 8, f"Teacher: {teacher}", ln=True, align="L")
pdf.cell(200, 8, f"Subject: {subject}", ln=True, align="L")
pdf.cell(200, 8, f"Test: {test_name}", ln=True, align="L")
pdf.cell(200, 8, f"Date: {test_date}", ln=True, align="L")
pdf.cell(200, 8, f"Max Marks: {max_marks}", ln=True, align="L")
pdf.ln(6)

# Table
pdf.set_font("Arial", "B", 11)
pdf.set_fill_color(220, 220, 220)
for col in df.columns:
    pdf.cell(38, 10, str(col), border=1, align="C", fill=True)
pdf.ln()
pdf.set_font("Arial", "", 10)
for _, row in df.iterrows():
    for col in df.columns:
        pdf.cell(38, 8, str(row[col]), border=1, align="C")
    pdf.ln()

pdf.ln(6)
pdf.set_font("Arial", "B", 12)
pdf.cell(200, 8, "Summary:", ln=True)
pdf.set_font("Arial", "", 11)
pdf.cell(200, 8, f"Total Students: {len(df)}", ln=True)
if 'Total' in df.columns:
    pdf.cell(200, 8, f"Class Average: {df['Total'].mean():.2f}", ln=True)

pdf_output = pdf.output(dest='S').encode('latin-1')

col1, col2 = st.columns(2)
col1.download_button(
    "⬇️ Download Excel Report",
    data=excel_data,
    file_name="All_Tests_Record.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
col2.download_button(
    "⬇️ Download PDF Report",
    data=pdf_output,
    file_name=f"{_class}_Report.pdf",
    mime="application/pdf"
)

st.success("✅ You can now save this test and add another later!")
