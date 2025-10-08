import streamlit as st
import pandas as pd
import datetime
from io import BytesIO

st.set_page_config(page_title="Prosper Marks Manager Pro", layout="wide")

st.title("📘 Prosper Marks Manager Pro")

# --- Header Information ---
st.subheader("School Marks Management Portal")
school_name = st.text_input("School Name", "Bright Future School")
subject = st.text_input("Subject", "Biology")
teacher = st.text_input("Teacher", "Mr. Prosper Dushimirimana")
term = st.selectbox("Term", ["Term 1", "Term 2", "Term 3"])
class_selected = st.selectbox("Select Class", ["Senior 1", "Senior 2", "Senior 3", "Senior 4", "Senior 5", "Senior 6"])

st.divider()

# --- Student Names Upload ---
uploaded_file = st.file_uploader("📎 Upload Excel File of Students (with a column 'Name')", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    if "Name" not in df.columns:
        st.error("The Excel file must contain a 'Name' column.")
    else:
        st.success(f"Loaded {len(df)} student names successfully!")

        # --- Test Information ---
        date_today = st.date_input("Date of Test", datetime.date.today())
        max_marks = st.number_input("Max Marks for Test 1", min_value=1, value=30)
        st.divider()

        # --- Enter Marks ---
        st.subheader("Enter Marks for Each Student")
        default_mark = st.number_input("Enter marks for all (can adjust later in Excel):", min_value=0, max_value=max_marks, value=0)
        df["Test 1"] = default_mark
        df["Total"] = df["Test 1"]
        df["Rank"] = df["Total"].rank(ascending=False, method="min").astype(int)

        # --- Display Marks Table ---
        st.subheader("Marks Record Table")
        st.dataframe(df.style.hide(axis="index"))

        # --- Save File per Class ---
        filename = f"{class_selected.replace(' ', '_')}_marks.xlsx"
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Marks")
        st.download_button(
            label="⬇️ Download Excel File",
            data=buffer.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # --- PDF Export ---
        pdf_name = f"{class_selected.replace(' ', '_')}_report.pdf"
        pdf_buffer = BytesIO()
        from reportlab.lib.pagesizes import A4
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet

        doc = SimpleDocTemplate(pdf_buffer, pagesize=A4)
        elements = []
        styles = getSampleStyleSheet()
        elements.append(Paragraph(f"<b>{school_name}</b>", styles['Title']))
        elements.append(Paragraph(f"Subject: {subject} | Class: {class_selected} | Teacher: {teacher}", styles['Heading3']))
        elements.append(Paragraph(f"Date: {date_today} | Term: {term}", styles['Normal']))
        elements.append(Spacer(1, 12))

        table_data = [df.columns.tolist()] + df.values.tolist()
        table = Table(table_data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
        ]))
        elements.append(table)
        doc.build(elements)

        st.download_button(
            label="📄 Download PDF Report",
            data=pdf_buffer.getvalue(),
            file_name=pdf_name,
            mime="application/pdf",
        )

        # --- Auto-save data ---
        if "saved_data" not in st.session_state:
            st.session_state["saved_data"] = {}
        st.session_state["saved_data"][class_selected] = df
        st.success(f"✅ Data auto-saved for {class_selected}")

else:
    st.info("Please upload an Excel file containing a column named 'Name'.")
