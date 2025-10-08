import streamlit as st
import pandas as pd
from datetime import date
import io
from fpdf import FPDF

st.set_page_config(page_title="Marks Management", layout="wide")

st.title("📘 Student Marks Management System")

# ==============================
# 1️⃣ Upload workbook
# ==============================
uploaded_file = st.file_uploader("Upload Excel workbook containing student list", type=["xlsx"])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    if 'Name' not in df.columns:
        st.error("⚠️ The Excel file must contain a column named 'Name'.")
        st.stop()
else:
    st.warning("Please upload the Excel file first.")
    st.stop()

# ==============================
# 2️⃣ Basic Info Section
# ==============================
st.subheader("Class Information")
col1, col2, col3, col4 = st.columns(4)
district = col1.text_input("District:")
school = col2.text_input("School:")
teacher = col3.text_input("Teacher:")
_class = col4.selectbox("Class:", ["S1A", "S1B", "S2A", "S2B", "S3A", "S3B", "S4A", "S5A", "S6A"])

st.divider()

# ==============================
# 3️⃣ Test Information
# ==============================
st.subheader("Enter Test Details")
test_name = st.text_input("Test Name (e.g. Test 1, Midterm, etc.):")
test_date = st.date_input("Date for this Test", date.today())
max_marks = st.number_input("Maximum Marks", min_value=1, max_value=100, value=30)

st.divider()

# ==============================
# 4️⃣ Enter Marks
# ==============================
st.subheader(f"Enter Marks for Each Student in {test_name}")
marks = []
for student in df['Name']:
    mark = st.number_input(f"{student}", min_value=0, max_value=max_marks, value=0, key=student)
    marks.append(mark)

df[test_name] = marks

# ==============================
# 5️⃣ Calculate Total and Rank
# ==============================
if 'Total' not in df.columns:
    df['Total'] = 0
df['Total'] = df.filter(like="Test").sum(axis=1)
df['Rank'] = df['Total'].rank(ascending=False, method='min').astype(int)

# ==============================
# 6️⃣ Save Updates to Excel
# ==============================
st.divider()
st.subheader("Save or Download Reports")

save_button = st.button("💾 Update Workbook (All Tests)")
if save_button:
    with pd.ExcelWriter("All_Tests_Record.xlsx", engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=_class)
    st.success("✅ Workbook updated successfully!")

# ==============================
# 7️⃣ Download Options
# ==============================
excel_buffer = io.BytesIO()
df.to_excel(excel_buffer, index=False)
excel_data = excel_buffer.getvalue()

# ---- PDF Report ----
pdf = FPDF()
pdf.add_page()
pdf.set_font("Arial", "B", 14)
pdf.cell(200, 10, f"Marks Report - {_class}", ln=True, align="C")
pdf.set_font("Arial", "", 12)
pdf.cell(200, 10, f"District: {district} | School: {school} | Teacher: {teacher}", ln=True, align="L")
pdf.cell(200, 10, f"Test: {test_name} | Date: {test_date} | Max Marks: {max_marks}", ln=True, align="L")
pdf.ln(8)

pdf.set_font("Arial", "B", 11)
for col in df.columns:
    pdf.cell(38, 10, str(col), border=1)
pdf.ln()

pdf.set_font("Arial", "", 10)
for _, row in df.iterrows():
    for col in df.columns:
        pdf.cell(38, 8, str(row[col]), border=1)
    pdf.ln()

pdf_buffer = io.BytesIO()
pdf.output(pdf_buffer)
pdf_data = pdf_buffer.getvalue()

col1, col2 = st.columns(2)
col1.download_button("⬇️ Download Excel", data=excel_data, file_name="All_Tests_Record.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
col2.download_button("⬇️ Download PDF", data=pdf_data, file_name=f"{_class}_Report.pdf", mime="application/pdf")

st.success("✅ Ready! You can now enter another test later — it will update automatically.")
