import streamlit as st
import pandas as pd
from datetime import date
from io import BytesIO

st.set_page_config(page_title="Marks Recorder", page_icon="🧮", layout="wide")

st.title("📘 Student Marks Recorder")
st.markdown("Enter marks for each student, then download the updated workbook with test details and totals.")

# -----------------------------
# File uploader
# -----------------------------
uploaded_file = st.file_uploader("📂 Upload the class Excel workbook", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # --- Detect the column for student names automatically ---
    name_col = None
    for possible in ["Names", "Name", "names", "Student Name", "student_name"]:
        if possible in df.columns:
            name_col = possible
            break

    if not name_col:
        st.error("❌ 'Name' column not found in your Excel file. Please ensure one column has student names.")
        st.stop()

    # --- Show detected column ---
    st.success(f"✅ Using column **{name_col}** for student names.")

    # --- Let user select class or sheet ---
    st.markdown("### 🧑‍🏫 Select Class or Section")
    class_name = st.text_input("Enter class name (e.g. S1 Biology)", "")

    # --- Date for this record ---
    record_date = st.date_input("📅 Date", date.today())

    # --- Input scores ---
    st.markdown("## ✏️ Enter Scores for Each Student")

    test_cols = ["Test 1", "Test 2", "Test 3", "Test 4", "Test 5"]
    max_marks = st.number_input("Maximum marks per test", 1, 100, 20)

    # Create an empty dataframe to store marks
    marks_data = {name_col: df[name_col].tolist()}

    for test in test_cols:
        marks_data[test] = [0] * len(df)

    marks_df = pd.DataFrame(marks_data)

    for i, row in df.iterrows():
        name = row[name_col]
        st.markdown(f"### 👤 {name}")
        for test in test_cols:
            marks_df.loc[i, test] = st.number_input(f"{test} score for {name}", 0, max_marks, 0, key=f"{name}_{test}")

    # --- Compute totals and percentage ---
    marks_df["Total (/Max)"] = marks_df[test_cols].sum(axis=1).astype(int).astype(str) + f" / {max_marks*5}"
    marks_df["Total (/100)"] = round((marks_df[test_cols].sum(axis=1) / (max_marks * 5)) * 100, 2)

    # --- Add class info and date ---
    marks_df["Class"] = class_name
    marks_df["Record Date"] = record_date
    marks_df["Year"] = record_date.year

    st.markdown("### ✅ Preview of recorded marks")
    st.dataframe(marks_df)

    # --- Download updated Excel ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        marks_df.to_excel(writer, index=False, sheet_name="Marks")
    processed_data = output.getvalue()

    st.download_button(
        label="📥 Download Updated Workbook",
        data=processed_data,
        file_name=f"{class_name}_marks_{record_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("👆 Please upload a workbook to begin.")
