# app.py
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, date
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="Smart Assessment — Final", layout="wide")
st.title("📊 Smart Assessment — Final (Class workbook + progressive tests)")

# ---------- Helpers ----------
def short_date(d):
    if isinstance(d, (datetime, date)):
        return d.strftime("%d/%m/%Y")
    return str(d)

def compute_totals(df_scores, test_cols, test_maxes):
    df = df_scores.copy()
    df["Total"] = df[test_cols].sum(axis=1)
    total_max = sum(test_maxes)
    df["/Max"] = total_max
    df["/100"] = (df["Total"] / total_max * 100).round(2) if total_max > 0 else 0
    # Rank by Total
    df["Rank"] = df["Total"].rank(ascending=False, method="min").astype(int)
    return df

def df_to_excel_bytes(df_final, metadata, test_dates, test_names, test_maxes):
    # Build Excel with metadata, date row on top, tests row next, then student rows
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        sheet = "Report"
        # Metadata block (vertical)
        meta_df = pd.DataFrame(list(metadata.items()), columns=["Field", "Value"])
        meta_df.to_excel(writer, sheet_name=sheet, index=False, header=False, startrow=0)
        startrow = len(meta_df) + 2

        # Build header rows: first row = Date row, second row = Test names row
        # We will construct a DataFrame whose columns match final table.
        cols = ["SN", "Name"] + [f"T{i+1}" for i in range(len(test_names))] + ["Total", "/Max", "/100", "Rank"]
        # Create empty df for headers
        header_df = pd.DataFrame(columns=cols)

        # Dates row values
        dates_vals = ["", "Date"] + [short_date(d) for d in test_dates] + ["", "", "", ""]
        tests_vals = ["", "Test"] + [f"{test_names[i]} (/ {test_maxes[i]})" for i in range(len(test_names))] + ["", "", "", ""]

        # Write the two header rows manually
        worksheet = writer.sheets[sheet]
        # write dates row
        for c, val in enumerate(dates_vals):
            worksheet.write(startrow, c, val)
        # write tests row
        for c, val in enumerate(tests_vals):
            worksheet.write(startrow + 1, c, val)

        # Now write student rows starting at startrow + 2
        df_final.to_excel(writer, sheet_name=sheet, index=False, startrow=startrow + 2)
        # Format columns widths a bit
        for i, col in enumerate(df_final.columns):
            width = max(12, min(40, int(max(df_final[col].astype(str).map(len).max(), len(col)) + 2)))
            worksheet.set_column(i, i, width)
    output.seek(0)
    return output.read()

def df_to_pdf_bytes(df_final, metadata, test_dates, test_names, test_maxes):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), leftMargin=18, rightMargin=18, topMargin=18, bottomMargin=18)
    styles = getSampleStyleSheet()
    elems = []
    title = f"MARKS RECORD — {metadata.get('Subject','')} — {metadata.get('Term','')}"
    elems.append(Paragraph(title, styles['Title']))
    elems.append(Spacer(1,6))
    meta_lines = "<br/>".join([f"{k}: {v}" for k,v in metadata.items()])
    elems.append(Paragraph(meta_lines, styles['Normal']))
    elems.append(Spacer(1,8))

    # Build table data: date row, test row, then student rows
    header = ["SN", "Name"] + [f"T{i+1}" for i in range(len(test_names))] + ["Total", "/Max", "/100", "Rank"]
    data = []
    data.append(["", "Date"] + [short_date(d) for d in test_dates] + ["", "", "", ""])
    data.append(["", "Test"] + [f"{test_names[i]} (/ {test_maxes[i]})" for i in range(len(test_names))] + ["", "", "", ""])
    # student rows
    for idx, row in df_final.iterrows():
        r = [idx+1, row["Name"]]
        for tcol in [c for c in df_final.columns if c.startswith("T")]:
            r.append(row.get(tcol, ""))
        r += [row.get("Total",""), row.get("/Max",""), row.get("/100",""), row.get("Rank","")]
        data.append(r)

    table = Table([header] + data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#dbe9f7")),
        ('BACKGROUND', (0,1), (-1,1), colors.HexColor("#f7f7f7")),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.4, colors.grey),
        ('FONTSIZE', (0,0), (-1,-1), 9),
    ]))
    elems.append(table)
    doc.build(elems)
    buffer.seek(0)
    return buffer.read()

# ---------- UI: metadata ----------
st.header("1) School & Report Info")
c1, c2, c3 = st.columns(3)
with c1:
    district = st.text_input("District")
    sector = st.text_input("Sector")
    school = st.text_input("School")
with c2:
    academic_year = st.text_input("Academic Year (e.g. 2025)")
    subject = st.text_input("Subject")
    teacher = st.text_input("Teacher")
with c3:
    term = st.selectbox("Term", ["Term I","Term II","Term III"])
    # choose how to provide student list
    upload_mode = st.radio("Student list source", ["Upload workbook (recommended)","Manual entry"], index=0)

# ---------- Input: workbook or manual ----------
uploaded = None
sheet_names = []
if upload_mode == "Upload workbook (recommended)":
    uploaded = st.file_uploader("Upload Excel workbook (.xlsx) with sheets per class (each sheet should have Names in first column)", type=["xlsx"])
    if uploaded:
        xls = pd.ExcelFile(uploaded)
        sheet_names = xls.sheet_names
        selected_sheet = st.selectbox("Select class (sheet)", sheet_names)
        # read selected sheet
        df_sheet = pd.read_excel(xls, sheet_name=selected_sheet, header=0)
        # detect name column
        name_col = None
        for possible in ["Names","Name","Student Name","Full Name"]:
            if possible in df_sheet.columns:
                name_col = possible
                break
        if name_col is None:
            # fallback to first column
            name_col = df_sheet.columns[0]
        student_names = df_sheet[name_col].dropna().astype(str).tolist()
else:
    # manual entry
    n = st.number_input("Number of students (manual)", min_value=1, value=10, step=1)
    student_names = []
    for i in range(n):
        student_names.append(st.text_input(f"Student {i+1} name", key=f"manual_name_{i}"))

# ---------- Tests setup ----------
st.header("2) Tests Setup")
mode = st.radio("Mode", ["Enter all tests now", "Update/Enter single test at a time"], index=1)

# We'll keep tests metadata in session_state so user can update progressively
if 'tests_meta' not in st.session_state:
    st.session_state['tests_meta'] = []  # list of dicts: {'name','date','max'}

if mode == "Enter all tests now":
    m = st.number_input("Number of tests to create", min_value=1, value=max(1,len(st.session_state['tests_meta'])), key="num_tests_all")
    temp_tests = []
    for i in range(m):
        col_a, col_b, col_c = st.columns([3,2,2])
        with col_a:
            tname = st.text_input(f"Test {i+1} name", value=(st.session_state['tests_meta'][i]['name'] if i < len(st.session_state['tests_meta']) else f"Test {i+1}"), key=f"all_name_{i}")
        with col_b:
            tdate = st.date_input(f"Date for Test {i+1}", value=(st.session_state['tests_meta'][i]['date'] if i < len(st.session_state['tests_meta']) else date.today()), key=f"all_date_{i}")
        with col_c:
            tmax = st.number_input(f"Max for Test {i+1}", min_value=1, value=(st.session_state['tests_meta'][i]['max'] if i < len(st.session_state['tests_meta']) else 20), key=f"all_max_{i}")
        temp_tests.append({'name': tname, 'date': tdate, 'max': int(tmax)})
    # store in session
    st.session_state['tests_meta'] = temp_tests

else:
    # Update single test mode (progressive)
    # Select test number to update or create new
    existing = st.session_state.get('tests_meta', [])
    existing_names = [t['name'] for t in existing]
    choice = st.selectbox("Choose existing test to update or create new", options = existing_names + ["<Create new test>"])
    if choice == "<Create new test>":
        new_name = st.text_input("New test name", "Test X", key="new_test_name")
        new_date = st.date_input("Date for new test", date.today(), key="new_test_date")
        new_max = st.number_input("Max for new test", min_value=1, value=20, key="new_test_max")
        if st.button("Add new test"):
            st.session_state['tests_meta'].append({'name': new_name, 'date': new_date, 'max': int(new_max)})
            st.success("New test added to session.")
    else:
        # update marks for selected existing test
        sel_idx = existing_names.index(choice)
        st.info(f"Updating marks for: {choice}")
        # allow changing date and max
        u_date = st.date_input("Date", value=st.session_state['tests_meta'][sel_idx]['date'], key=f"upd_date_{sel_idx}")
        u_max = st.number_input("Max marks", min_value=1, value=st.session_state['tests_meta'][sel_idx]['max'], key=f"upd_max_{sel_idx}")
        if st.button("Update test metadata"):
            st.session_state['tests_meta'][sel_idx]['date'] = u_date
            st.session_state['tests_meta'][sel_idx]['max'] = int(u_max)
            st.success("Test metadata updated")

# ---------- Marks entry ----------
st.header("3) Enter / Update Marks")
tests_meta = st.session_state.get('tests_meta', [])
if len(tests_meta) == 0:
    st.info("No tests defined yet. Please add tests (in 'Enter all tests now' mode or 'Add new test').")
else:
    # Build or load marks_df in session (index: student names)
    if 'marks_df' not in st.session_state:
        # create empty marks df with Name column and T1..Tn columns as available tests
        cols = ['Name'] + [f"T{i+1}" for i in range(len(tests_meta))]
        marks_init = pd.DataFrame(columns=cols)
        marks_init['Name'] = student_names
        # initialize test cols to NaN
        for i in range(len(tests_meta)):
            marks_init[f"T{i+1}"] = pd.NA
        st.session_state['marks_df'] = marks_init

    marks_df = st.session_state['marks_df']

    # If number of tests changed, ensure columns match
    expected_cols = ['Name'] + [f"T{i+1}" for i in range(len(tests_meta))]
    for c in expected_cols:
        if c not in marks_df.columns:
            marks_df[c] = pd.NA
    # keep only expected columns
    marks_df = marks_df[['Name'] + [c for c in marks_df.columns if c.startswith('T')]]
    # if names differ from uploaded/manual, reset names column
    if len(student_names) != 0 and list(marks_df['Name'].astype(str)) != student_names:
        marks_df['Name'] = student_names

    st.session_state['marks_df'] = marks_df  # update

    st.write("Edit marks below (leave blank for tests not yet taken).")
    edited = st.experimental_data_editor(marks_df, num_rows="dynamic")  # editable table

    # Save edits back
    st.session_state['marks_df'] = edited.copy()

    # Buttons: Save progress / Generate final report / Download
    col_save, col_report, col_clear = st.columns([1,1,1])
    with col_save:
        if st.button("💾 Save progress (in session)"):
            st.success("Progress saved in session. You can continue later in this browser session.")
    with col_report:
        if st.button("📊 Generate Final Report"):
            # compute totals
            test_cols = [c for c in edited.columns if c.startswith('T')]
            test_maxes = [t['max'] for t in tests_meta]
            # ensure all test columns exist (if some tests added later)
            for i, t in enumerate(tests_meta):
                col_name = f"T{i+1}"
                if col_name not in edited.columns:
                    edited[col_name] = 0
            # fill NaN with 0 for computation
            comp_df = edited.copy()
            for c in test_cols:
                comp_df[c] = pd.to_numeric(comp_df[c], errors='coerce').fillna(0)
            final_df = compute_totals(comp_df[['Name'] + test_cols].rename(columns={c:c for c in test_cols}), test_cols, test_maxes)
            # final_df has Name, test cols, Total, /Max, /100, Rank
            # prepare metadata
            metadata = {
                "District": district,
                "Sector": sector,
                "School": school,
                "Class": selected_sheet if uploaded else "(manual)",
                "Academic Year": academic_year,
                "Term": term,
                "Subject": subject,
                "Teacher": teacher
            }
            # Excel bytes
            excel_bytes = df_to_excel_bytes(final_df[['Name'] + test_cols + ['Total','/Max','/100','Rank']].rename(columns=lambda x: ("T"+x[1]) if x.startswith("T") else x),
                                           metadata,
                                           [t['date'] for t in tests_meta],
                                           [t['name'] for t in tests_meta],
                                           [t['max'] for t in tests_meta])
            st.download_button("⬇️ Download Excel Report", data=excel_bytes, file_name=f"{subject}_{selected_sheet if uploaded else 'class'}_Report.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            # PDF bytes
            pdf_bytes = df_to_pdf_bytes(final_df[['Name'] + test_cols + ['Total','/Max','/100','Rank']].rename(columns=lambda x: ("T"+x[1]) if x.startswith("T") else x),
                                       metadata,
                                       [t['date'] for t in tests_meta],
                                       [t['name'] for t in tests_meta],
                                       [t['max'] for t in tests_meta])
            st.download_button("⬇️ Download PDF Report", data=pdf_bytes, file_name=f"{subject}_{selected_sheet if uploaded else 'class'}_Report.pdf", mime="application/pdf")

    with col_clear:
        if st.button("🧹 Clear session data"):
            for k in ['tests_meta','marks_df']:
                if k in st.session_state:
                    del st.session_state[k]
            st.experimental_rerun()

else:
    st.info("Please upload a workbook or enter manual student names and define tests to begin.")
