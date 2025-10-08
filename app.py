import streamlit as st
import pandas as pd
from datetime import date

st.set_page_config(page_title="Prosper Marks Record", layout="centered")

st.title("📘 Student Marks Record System")

# --- Class selection ---
classes = ["S1", "S2", "S3", "S4", "S5", "S6"]
selected_class = st.selectbox("Select Class", classes)

# --- Test configuration ---
st.subheader("🧾 Test Details")
test_names = ["Test 1", "Test 2", "Test 3"]
test_dates = []
max_marks = []

for test in test_names:
    test_dates.append(
        st.date_input(f"📅 Date for {test}", date.today())
    )
    max_marks.append(
        st.number_input(f"Maximum Marks for {test}", min_value=1, value=30, step=1)
    )

st.divider()

# --- Upload Excel or create default student list ---
uploaded_file = st.file_uploader("📂 Upload Excel file with student names (optional)", type=["xlsx", "csv"])

if uploaded_file is not None:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
else:
    # Default student list
    df = pd.DataFrame({
        "S/N": [1, 2, 3],
        "Name": ["ABIZERIMANA Deborah", "AKARIZA ESTHER", "AKIMANA Lea"],
        "Gender": ["FEMALE", "FEMALE", "FEMALE"]
    })

st.subheader("✏️ Enter Marks for Each Student")

marks_data = []
for idx, row in df.iterrows():
    st.markdown(f"### 🧍 {row['Name']}")
    student_marks = {"S/N": row["S/N"], "Name": row["Name"], "Gender": row["Gender"]}

    for i, test in enumerate(test_names):
        mark = st.number_input(
            f"Marks for {test} ({test_dates[i].strftime('%d-%b-%Y')}) - {row['Name']}",
            min_value=0,
            max_value=max_marks[i],
            step=1,
            key=f"{row['Name']}_{test}"
        )
        student_marks[test] = mark

    marks_data.append(student_marks)
    st.divider()

marks_df = pd.DataFrame(marks_data)

# --- Calculate totals and ranks ---
marks_df["Total"] = marks_df[test_names].sum(axis=1)
marks_df["Rank"] = marks_df["Total"].rank(ascending=False, method="min").astype(int)

st.divider()
st.subheader("📊 Marks Record Table")
st.dataframe(marks_df, use_container_width=True)

# --- Download Report ---
def convert_to_excel(df):
    from io import BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Marks Record")
    return output.getvalue()

excel_data = convert_to_excel(marks_df)

st.download_button(
    label="⬇️ Download Marks Report (Excel)",
    data=excel_data,
    file_name=f"{selected_class}_Marks_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.success("✅ Marks record ready and downloadable!")
