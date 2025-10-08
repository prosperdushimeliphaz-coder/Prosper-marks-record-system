import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="Prosper Marks Register", layout="wide")

st.title("📘 Student Marks Register")

# Upload Excel file
uploaded_file = st.file_uploader("📂 Upload Excel file", type=["xlsx"])
if uploaded_file:
    # Read Excel and extract sheet names
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names

    # Dropdown for classes (auto from sheet names)
    selected_class = st.selectbox("Select Class:", sheet_names)
    df = pd.read_excel(xls, sheet_name=selected_class)
    
    st.success(f"✅ Loaded class: {selected_class}")

    # Basic information
    st.subheader("School Information")
    col1, col2 = st.columns(2)
    with col1:
        district = st.text_input("District:")
        school = st.text_input("School:")
    with col2:
        teacher_name = st.text_input("Teacher’s Name:")
        subject_name = st.text_input("Subject Name:")

    # Tests info
    st.subheader("Test Setup")
    num_tests = st.number_input("Number of Tests:", min_value=1, max_value=10, value=1)
    test_dates, test_max = [], []
    for i in range(num_tests):
        c1, c2 = st.columns(2)
        with c1:
            date = st.date_input(f"Date of Test {i+1}")
        with c2:
            max_mark = st.number_input(f"Max marks for Test {i+1}", min_value=1, value=20)
        test_dates.append(str(date))
        test_max.append(max_mark)

    # Enter Marks
    st.subheader("Enter Marks for Each Student")
    marks_data = {}
    for i, name in enumerate(df.iloc[:, 0]):  # first column = names
        st.write(f"**{name}**")
        student_marks = []
        for j in range(num_tests):
            mark = st.number_input(f"{name} - Test {j+1}", min_value=0, max_value=test_max[j], value=0, key=f"{i}_{j}")
            student_marks.append(mark)
        marks_data[name] = student_marks

    # Generate reports
    if st.button("Generate Report"):
        # Prepare table
        headers = ["Name"] + [f"Test {i+1}" for i in range(num_tests)] + ["Total", "Average (%)"]
        data = [headers]
        for name, marks in marks_data.items():
            total = sum(marks)
            avg = round(total / sum(test_max) * 100, 2)
            data.append([name] + marks + [total, avg])
        
        df_report = pd.DataFrame(data[1:], columns=data[0])
        st.dataframe(df_report, use_container_width=True)
        
        # Create PDF
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        elements = []

        header = f"""
        <b>{school}</b> — <b>{selected_class}</b><br/>
        {district}<br/>
        <b>Subject:</b> {subject_name} | <b>Teacher:</b> {teacher_name}
        """
        elements.append(Paragraph(header, styles["Title"]))
        elements.append(Spacer(1, 12))

        table = Table(data, repeatRows=1)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
        ]))
        elements.append(table)
        doc.build(elements)

        # Download PDF
        st.download_button(
            label="📄 Download PDF Report",
            data=buffer.getvalue(),
            file_name=f"{selected_class}_{subject_name}_Report.pdf",
            mime="application/pdf"
        )

        # Download Excel
        excel_buffer = BytesIO()
        df_report.to_excel(excel_buffer, index=False)
        st.download_button(
            label="📊 Download Excel Report",
            data=excel_buffer.getvalue(),
            file_name=f"{selected_class}_{subject_name}_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
