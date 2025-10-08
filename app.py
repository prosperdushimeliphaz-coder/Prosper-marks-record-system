import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import base64

# --- PAGE SETUP ---
st.set_page_config(page_title="Prosper Marks Manager Pro", layout="wide")

st.title("📊 Prosper Marks Manager Pro")
st.markdown("Easily record, update and generate student performance reports.")


# --- SCHOOL INFO SECTION ---
st.sidebar.header("🏫 School Information")
district = st.sidebar.text_input("District")
sector = st.sidebar.text_input("Sector")
school = st.sidebar.text_input("School name")
academic_year = st.sidebar.text_input("Academic year (e.g. 2024-2025)")
term = st.sidebar.selectbox("Term", ["Term 1", "Term 2", "Term 3"])
subject = st.sidebar.text_input("Subject")
teacher = st.sidebar.text_input("Teacher name")


# --- FILE UPLOAD ---
st.header("📘 Upload Students Workbook")
uploaded_file = st.file_uploader("Upload Excel file containing students list", type=["xlsx"])

if uploaded_file:
    excel_file = pd.ExcelFile(uploaded_file)
    sheet_names = excel_file.sheet_names
    selected_sheet = st.selectbox("Select class (sheet)", sheet_names)

    df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
    
    if "Name" not in df.columns:
        st.error("No 'Name' column found in the sheet! Please ensure a column named 'Name' exists.")
    else:
        st.success(f"✅ Loaded class list for: {selected_sheet}")
        st.dataframe(df[["Name"]])

        # --- TEST ENTRY SECTION ---
        st.subheader("🧾 Enter Tests and Marks")

        num_tests = st.number_input("Number of tests", min_value=1, max_value=10, step=1, value=3)
        test_names, test_dates, max_marks = [], [], []

        for i in range(num_tests):
            col1, col2, col3 = st.columns(3)
            with col1:
                test_name = st.text_input(f"Test {i+1} Name", f"Test {i+1}")
            with col2:
                test_date = st.date_input(f"Date for {test_name}", datetime.today())
            with col3:
                max_mark = st.number_input(f"Max marks for {test_name}", min_value=1, max_value=200, value=30)
            test_names.append(test_name)
            test_dates.append(test_date)
            max_marks.append(max_mark)

        # --- MARK ENTRY FOR EACH STUDENT ---
        st.subheader("✏️ Enter Marks for Each Student")

        marks_data = []
        for idx, row in df.iterrows():
            st.markdown(f"**{row['Name']}**")
            student_marks = {"Name": row["Name"]}
            cols = st.columns(num_tests)
            for i, test in enumerate(test_names):
                with cols[i]:
                    mark = st.number_input(
                        f"{test} ({test_dates[i].strftime('%d-%b')})",
                        min_value=0,
                        max_value=max_marks[i],
                        step=1,
                        key=f"{row['Name']}_{test}"
                    )
                    student_marks[test] = mark
            marks_data.append(student_marks)
            st.divider()

        marks_df = pd.DataFrame(marks_data)

        # --- CALCULATE TOTALS, PERCENT, RANK ---
        for t in test_names:
            if t not in marks_df.columns:
                marks_df[t] = 0

        marks_df["Total"] = marks_df[test_names].sum(axis=1)
        total_max = sum(max_marks)
        marks_df["%"] = (marks_df["Total"] / total_max) * 100
        marks_df["Rank"] = marks_df["Total"].rank(ascending=False, method='min').astype(int)

        # --- FINAL REPORT MERGE ---
        report_df = marks_df.copy()
        report_df = report_df[["Name"] + test_names + ["Total", "%", "Rank"]]

        st.write("### 🧮 Marks Record Table")
        st.dataframe(report_df, use_container_width=True)

        # --- SUMMARY INFO (Header of report) ---
        st.write("### 🏫 Report Header Information")
        st.markdown(f"""
        **District:** {district}  
        **Sector:** {sector}  
        **School:** {school}  
        **Class:** {selected_sheet}  
        **Academic Year:** {academic_year}  
        **Term:** {term}  
        **Subject:** {subject}  
        **Teacher:** {teacher}  
        **Date Generated:** {datetime.today().strftime("%d-%b-%Y")}
        """)

        # --- EXPORT SECTION ---
        st.write("---")
        st.subheader("⬇️ Download Options")

        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Report')
            return output.getvalue()

        excel_data = to_excel(report_df)
        st.download_button(
            label="📥 Download Excel Report",
            data=excel_data,
            file_name=f"{selected_sheet}_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # PDF EXPORT (simple text-based)
        def generate_pdf(df):
            html = f"""
            <h2>School Marks Report</h2>
            <p><b>District:</b> {district}<br>
            <b>Sector:</b> {sector}<br>
            <b>School:</b> {school}<br>
            <b>Class:</b> {selected_sheet}<br>
            <b>Academic Year:</b> {academic_year}<br>
            <b>Term:</b> {term}<br>
            <b>Subject:</b> {subject}<br>
            <b>Teacher:</b> {teacher}<br>
            <b>Date Generated:</b> {datetime.today().strftime("%d-%b-%Y")}</p>
            {df.to_html(index=False)}
            """
            return BytesIO(html.encode('utf-8'))

        pdf_data = generate_pdf(report_df)
        st.download_button(
            label="📄 Download PDF Report",
            data=pdf_data,
            file_name=f"{selected_sheet}_Report.pdf",
            mime="application/pdf"
        )

else:
    st.info("👆 Please upload the Excel workbook first to begin.")
