import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="Prosper Marks Register", layout="wide")

st.title("📘 Student Marks Register")

uploaded_file = st.file_uploader("📂 Upload Excel file", type=["xlsx"])

if uploaded_file:
    # Read Excel
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names

    # Select Class
    selected_class = st.selectbox("🏫 Select Class", sheet_names)
    df = pd.read_excel(xls, sheet_name=selected_class)

    # Detect name column
    if "Name" in df.columns:
        names = df["Name"].dropna().tolist()
    else:
        names = df.iloc[:, 0].dropna().tolist()  # assume first column

    st.success(f"✅ Loaded class: {selected_class} with {len(names)} students.")

    # Basic info
    st.subheader("School Information")
    col1, col2 = st.columns(2)
    with col1:
        district = st.text_input("District:")
        school = st.text_input("School:")
    with col2:
        teacher_name = st.text_input("Teacher’s Name:")
        subject_name = st.text_input("Subject Name:")

    # Test setup
    st.subheader("Tests Setup")
    num_tests = st.number_input("Number of Tests:", min_value=1, max_value=10, value=1)
    test_dates, test_max = [], []

    for i in range(num_tests):
        c1, c2 = st.columns(2)
        with c1:
            t_date = st.date_input(f"Date of Test {i+1}", value=date.today())
        with c2:
            max_mark = st.number_input(f"Max marks for Test {i+1}", min_value=1, value=20)
        test_dates.append(str(t_date))
        test_max.append(max_mark)

    # Enter marks
    st.subheader("✏️ Enter Marks for Each Student")
    marks_data = {}
    for i, name in enumerate(names):
        st.markdown(f"**{i+1}. {name}**")
        scores = []
        for j in range(num_tests):
            score = st.number_input(
                f"{name} - Test {j+1}",
                min_value=0,
                max_value=test_max[j],
                value=0,
                key=f"{i}_{j}"
            )
            scores.append(score)
        marks_data[name] = scores

    if st.button("💾 Save / Update Marks"):
        # Prepare data
        headers = ["Student Name"] + [f"Test {i+1}" for i in range(num_tests)] + ["Total", "Average (%)"]
        table_data = [headers]

        for name, scores in marks_data.items():
            total = sum(scores)
            average = round((total / sum(test_max)) * 100, 2)
            table_data.append([name] + scores + [total, average])

        df_report = pd.DataFrame(table_data[1:], columns=table_data[0])
        st.success("✅ Marks updated successfully!")
        st.dataframe(df_report, use_container_width=True)

        # --- PDF GENERATION ---
        pdf_buffer = BytesIO()
        doc = SimpleDocTemplate(pdf_buffer, pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        elements = []

        # Header info
        header = f"""
        <b>{school}</b> — <b>{selected_class}</b><br/>
        District: {district}<br/>
        <b>Subject:</b> {subject_name} | <b>Teacher:</b> {teacher_name}
        """
        elements.append(Paragraph(header, styles["Title"]))
        elements.append(Spacer(1, 12))

        # Table
        t = Table(table_data, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        elements.append(t)
        doc.build(elements)

        st.download_button(
            "📄 Download PDF Report",
            data=pdf_buffer.getvalue(),
            file_name=f"{selected_class}_{subject_name}_Report.pdf",
            mime="application/pdf"
        )

        # --- Excel output ---
        excel_buffer = BytesIO()
        df_report.to_excel(excel_buffer, index=False)
        st.download_button(
            "📊 Download Excel Report",
            data=excel_buffer.getvalue(),
            file_name=f"{selected_class}_{subject_name}_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
