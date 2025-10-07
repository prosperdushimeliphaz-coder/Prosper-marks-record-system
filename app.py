import streamlit as st
import pandas as pd

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Smart Assessment Report", layout="wide")

st.title("📊 Smart Assessment Report Generator")
st.markdown("Easily record marks and automatically generate a final performance report.")

# --- BASIC SCHOOL INFO ---
st.header("🏫 School & Class Information")

col1, col2, col3 = st.columns(3)
district = col1.text_input("District")
sector = col2.text_input("Sector")
school = col3.text_input("School")

col4, col5, col6 = st.columns(3)
academic_year = col4.text_input("Academic Year (e.g. 2025)")
_class = col5.text_input("Class (e.g. S2A)")
term = col6.selectbox("Term", ["Term I", "Term II", "Term III"])

st.divider()

# --- MARK ENTRY SECTION ---
st.header("🧮 Marks Record Entry")

num_students = st.number_input("Number of students", min_value=1, step=1)
num_tests = st.number_input("Number of assessments", min_value=1, step=1)

subject_max = st.number_input("Subject Overall Max (e.g., 40)", min_value=1, step=1)

students = []
for i in range(int(num_students)):
    st.subheader(f"Student {i+1}")
    student_name = st.text_input(f"Full name of student {i+1}")
    student_scores = []
    student_maxes = []
    for j in range(int(num_tests)):
        c1, c2 = st.columns(2)
        with c1:
            mark = st.number_input(f"Assessment {j+1} marks for {student_name}", min_value=0.0, step=0.5, key=f"m_{i}_{j}")
        with c2:
            maxm = st.number_input(f"Max marks for Assessment {j+1}", min_value=1.0, step=0.5, key=f"x_{i}_{j}")
        student_scores.append(mark)
        student_maxes.append(maxm)
    students.append((student_name, student_scores, student_maxes))

st.divider()

# --- GENERATE REPORT ---
if st.button("📄 Generate Final Report"):
    rows = []
    for s in students:
        name = s[0]
        scores = s[1]
        maxes = s[2]
        total = sum(scores)
        total_max = sum(maxes)
        percentage = round((total / total_max) * 100, 1) if total_max > 0 else 0
        final_mark = round((total / total_max) * subject_max, 1) if total_max > 0 else 0
        rows.append([name] + scores + [total, total_max, final_mark, percentage])

    # --- CREATE DATAFRAME ---
    columns = ["Name"] + [f"Assess {i+1}" for i in range(int(num_tests))] + ["Total", "Total Max", f"Final (/ {subject_max})", "%"]
    df = pd.DataFrame(rows, columns=columns)
    df["Rank"] = df["Total"].rank(ascending=False, method="min").astype(int)

    # --- DISPLAY REPORT ---
    st.header("📋 Generated Report")
    st.dataframe(df, use_container_width=True)

    # --- SUMMARY INFO ---
    st.markdown(f"**District:** {district}  **Sector:** {sector}  **School:** {school}")
    st.markdown(f"**Class:** {_class}  **Academic Year:** {academic_year}  **Term:** {term}")

    # --- EXPORT TO EXCEL ---
    excel_name = f"{school}_{_class}_{term}_Report.xlsx".replace(" ", "_")
    df.to_excel(excel_name, index=False)
    with open(excel_name, "rb") as f:
        st.download_button("⬇️ Download Report as Excel", f, file_name=excel_name)

st.divider()
st.caption("Developed by Prosper Dushimirimana | Powered by Streamlit & AI 🧠")
