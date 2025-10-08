import streamlit as st
import pandas as pd
from io import BytesIO
from fpdf import FPDF

st.set_page_config(page_title="Prosper Marks Record", layout="wide")

st.title("📘 Prosper Marks Record App")

# ---- SCHOOL INFO SECTION ----
st.subheader("School Information")
col1, col2, col3 = st.columns(3)
with col1:
    school = st.text_input("School Name", "")
    district = st.text_input("District", "")
with col2:
    level = st.text_input("Level", "")
    term = st.text_input("Term", "")
with col3:
    subject = st.text_input("Subject", "")
    teacher = st.text_input("Teacher Name", "")

st.markdown("---")

# ---- CLASS SELECTION ----
st.subheader("Select Class and Upload Workbook")
class_list = ["S1A", "S1B", "S2A", "S2B", "S3A", "S3B"]
selected_class = st.selectbox("Select Class", class_list)

uploaded_file = st.file_uploader("Upload Excel Workbook", type=["xlsx"])

# ---- PROCESS WORKBOOK ----
if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name=selected_class)
        df.columns = df.columns.str.strip()
        st.success(f"Workbook loaded successfully from class: {selected_class}")

        st.subheader("📋 Marks Record")
        st.write("Enter test names and dates:")

        num_tests = st.number_input("How many tests?", min_value=1, step=1)
        test_names, test_dates, max_marks = [], [], []

        for i in range(num_tests):
            cols = st.columns(3)
            with cols[0]:
                test_name = st.text_input(f"Test {i+1} Name", key=f"name_{i}")
            with cols[1]:
                test_date = st.date_input(f"Date for Test {i+1}", key=f"date_{i}")
            with cols[2]:
                max_mark = st.number_input(f"Max Marks for {test_name or f'Test {i+1}'}", min_value=1, step=1, key=f"max_{i}")

            test_names.append(test_name)
            test_dates.append(test_date)
            max_marks.append(max_mark)

        # Create marks entry table
        st.write("---")
        st.subheader("Enter Marks for Each Student")

        if "Name" not in df.columns:
            st.error("No 'Name' column found in the sheet! Please fix your Excel file.")
        else:
            for test_name in test_names:
                df[test_name] = st.number_input(
                    f"Enter marks for {test_name} (applied equally or adjust later in Excel):",
                    min_value=0,
                    max_value=100,
                    step=1,
                    key=f"marks_{test_name}"
                )

            # Calculate total and rank
            df["Total"] = df[test_names].sum(axis=1)
            df["Rank"] = df["Total"].rank(ascending=False, method='min').astype(int)

            st.write("### Marks Record Table")
            st.dataframe(df, use_container_width=True)

            # ---- EXPORT SECTION ----
            st.write("---")
            st.subheader("📤 Export Options")

            def convert_df_to_excel(dataframe):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    dataframe.to_excel(writer, index=False, sheet_name="Marks Record")
                return output.getvalue()

            excel_data = convert_df_to_excel(df)

            # --- Excel Export ---
            st.download_button(
                label="⬇️ Download Excel Report",
                data=excel_data,
                file_name=f"{school}_{subject}_{selected_class}_Marks.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # --- PDF Export ---
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", "B", 14)
            pdf.cell(200, 10, f"{school} - {district}", ln=True, align="C")
            pdf.cell(200, 10, f"Subject: {subject} | Class: {selected_class}", ln=True, align="C")
            pdf.cell(200, 10, f"Teacher: {teacher} | Term: {term} | Level: {level}", ln=True, align="C")
            pdf.ln(10)
            pdf.set_font("Arial", "B", 12)
            pdf.cell(200, 10, "Marks Record", ln=True, align="C")
            pdf.ln(5)

            pdf.set_font("Arial", size=10)
            for i, row in df.iterrows():
                line = ", ".join([f"{col}: {row[col]}" for col in df.columns])
                pdf.multi_cell(0, 8, line)

            pdf_output = BytesIO(pdf.output(dest="S").encode("latin1"))
            st.download_button(
                label="📄 Download PDF Report",
                data=pdf_output,
                file_name=f"{school}_{subject}_{selected_class}_Marks.pdf",
                mime="application/pdf"
            )

    except Exception as e:
        st.error(f"Error processing file: {e}")
else:
    st.info("Please upload a workbook to get started.")
