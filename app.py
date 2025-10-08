import streamlit as st
import pandas as pd
from datetime import date
from io import BytesIO
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="Prosper Marks Record", layout="centered")
st.title("📘 Student Marks Record System")

# =====================
# 1️⃣  SCHOOL INFO
# =====================
st.subheader("🏫 School Information")
district = st.text_input("District")
school = st.text_input("School")
teacher = st.text_input("Teacher Name")
subject = st.text_input("Subject")
selected_class = st.text_input("Class")

# =====================
# 2️⃣  UPLOAD STUDENTS
# =====================
st.subheader("📂 Upload Student List")
uploaded_file = st.file_uploader("Upload Excel file with student names", type=["xlsx", "csv"])

if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
else:
    st.warning("⚠️ Please upload a student list file to continue.")
    st.stop()

# =====================
# 3️⃣  TEST SETTINGS
# =====================
st.subheader("🧾 Test Setup")
num_tests = st.number_input("Enter number of tests", min_value=1, max_value=20, step=1, value=3)
test_names, test_dates, max_marks = [], [], []

for i in range(num_tests):
    st.markdown(f"**Test {i+1} Details:**")
    t_name = st.text_input(f"Name for Test {i+1}", f"Test {i+1}")
    t_date = st.date_input(f"Date for {t_name}", date.today(), key=f"date_{i}")
    t_max = st.number_input(f"Maximum Marks for {t_name}", min_value=1, value=30, step=1, key=f"max_{i}")
    test_names.append(t_name)
    test_dates.append(t_date)
    max_marks.append(t_max)
    st.markdown("---")

# =====================
# 4️⃣  ENTER MARKS
# =====================
st.subheader("✏️ Enter Marks for Each Student")
marks_data = []
for idx, row in df.iterrows():
    st.markdown(f"### 🧍 {row.iloc[0]}")  # assuming first column is name
    student_marks = {"Name": row.iloc[0]}
    for i, test in enumerate(test_names):
        mark = st.number_input(
            f"{test} ({test_dates[i]}) - {row.iloc[0]}",
            min_value=0,
            max_value=max_marks[i],
            step=1,
            key=f"{row.iloc[0]}_{test}"
        )
        student_marks[test] = mark
    marks_data.append(student_marks)
    st.divider()

marks_df = pd.DataFrame(marks_data)

# =====================
# 5️⃣  CALCULATE TOTALS & RANKS
# =====================
marks_df["Total"] = marks_df[test_names].sum(axis=1)
marks_df["Rank"] = marks_df["Total"].rank(ascending=False, method="min").astype(int)

st.divider()
st.subheader("📊 Marks Record Table")
st.dataframe(marks_df, use_container_width=True)

# =====================
# 6️⃣  DOWNLOAD EXCEL
# =====================
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Marks Record")
    return output.getvalue()

excel_bytes = to_excel(marks_df)
st.download_button(
    "⬇️ Download as Excel",
    data=excel_bytes,
    file_name=f"{selected_class}_Marks_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# =====================
# 7️⃣  DOWNLOAD PDF
# =====================
def to_pdf(df):
    output = BytesIO()
    pdf = SimpleDocTemplate(output, pagesize=landscape(A4))
    elements = []
    styles = getSampleStyleSheet()
    elements.append(Paragraph(f"<b>{school} - {subject} ({selected_class})</b>", styles["Title"]))
    elements.append(Paragraph(f"Teacher: {teacher} | District: {district}", styles["Normal"]))
    elements.append(Spacer(1, 12))

    data = [list(df.columns)] + df.values.tolist()
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightblue),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
    ]))
    elements.append(table)
    pdf.build(elements)
    return output.getvalue()

pdf_bytes = to_pdf(marks_df)
st.download_button(
    "⬇️ Download as PDF",
    data=pdf_bytes,
    file_name=f"{selected_class}_Marks_Report.pdf",
    mime="application/pdf"
)

st.success("✅ All done! Marks report ready for download in both Excel and PDF formats.")
