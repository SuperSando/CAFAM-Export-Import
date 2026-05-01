import streamlit as st
import pandas as pd
import io
import os
import base64
from datetime import datetime
from github import Github
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# --- HELPER FUNCTIONS ---
def format_to_hhmm(val):
    """Converts decimal hours (147.8) to Aviation HH:MM (147:48)"""
    if pd.isna(val) or val == "" or val == 0:
        return ""
    try:
        total_hours = float(val)
        hours = int(total_hours)
        minutes = int(round((total_hours - hours) * 60))
        if minutes == 60:
            hours += 1
            minutes = 0
        return f"{hours}:{minutes:02d}"
    except:
        return val

def log_to_github(filename):
    """Appends a row to log.csv in your GitHub repository"""
    try:
        token = st.secrets["GITHUB_TOKEN"]
        repo_name = st.secrets["REPO_NAME"]
        g = Github(token)
        repo = g.get_repo(repo_name)
        
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        new_row = f"{timestamp},{filename},Success\n"
        
        file_path = "log.csv"
        try:
            # Update existing log
            file_content = repo.get_contents(file_path)
            current_data = file_content.decoded_content.decode("utf-8")
            updated_data = current_data + new_row
            repo.update_file(file_path, f"Log update: {filename}", updated_data, file_content.sha)
        except:
            # Create new log if missing
            header = "Timestamp,Filename,Status\n"
            repo.create_file(file_path, "Initial log creation", header + new_row)
        return True
    except Exception as e:
        st.error(f"GitHub Logging failed: {e}")
        return False

# --- APP LAYOUT ---
st.set_page_config(page_title="CAFAM Transformer", page_icon="✈️", layout="wide")
st.title("✈️ CAFAM Record Transformer")

tab1, tab2 = st.tabs(["Transformer", "Permanent Audit Log"])

with tab1:
    st.write("Upload your handover export to apply RGV logic and formatting.")
    uploaded_file = st.file_uploader("Upload .xlsx file", type=["xlsx"])

    if uploaded_file:
        # Generate Dynamic Filename
        base_name = os.path.splitext(uploaded_file.name)[0]
        export_filename = f"{base_name} CAFAM Export.xlsx"

        # 1. Load Data
        df_raw = pd.read_excel(uploaded_file, engine='openpyxl')
        
        # 2. Capture B7 Reference Value (Aircraft Total Time)
        try:
            ref_b7_value = pd.to_numeric(df_raw.iloc[5, 1], errors='coerce')
        except:
            ref_b7_value = 0

        # 3. Delete Rows 2-9
        df = df_raw.drop(df_raw.index[0:8]).reset_index(drop=True)

        # 4. Amalgamate Description (Cols D, E, and others)
        col_d = df.iloc[:, 3].fillna('').astype(str)
        col_e = df.iloc[:, 4].fillna('').astype(str)
        others = df.iloc[:, 5:8].fillna('').astype(str).agg(' '.join, axis=1)
        description_data = (col_d + ": " + col_e + " " + others).str.replace(r'\s+', ' ', regex=True).str.strip().str.lstrip(': ')

        # 5. Reorganize and Rename
        cols_to_remove = df.columns[3:8]
        df = df.drop(columns=cols_to_remove)
        df.insert(3, 'Description', description_data)

        rename_map = {
            "DIRCTVE": "ATA / Task No.", "GROUP": "Group", "PARTNO": "Pn", "SERIAL": "Sn",
            "SCHD_HRS": "Int. FH", "SCHD_LDG": "Int. FC", "SCHD_DAYS": "Int. Cal.",
            "DATE_DUEBY": "Due date", "MANDTRY": "Mandatory", "REASON": "Remarks Office",
            "ACTION": "Action", "DATESAT": "Start date", "AC_HRS_SAT": "Ac. TT LSV",
            "TTSN": "Item TT", "TCSN": "Cycl. SN", "AC_LDG_SAT": "Item FC LSV"
        }
        df = df.rename(columns=rename_map)
        edited_headers = list(rename_map.values()) + ['Description', 'Comp.', 'Appl.', 'N/A', 'Ref_B7']

        # 6. Append DATE_DUEON to Remarks Office
        if 'DATE_DUEON' in df.columns:
            due_on_clean = pd.to_datetime(df['DATE_DUEON'], errors='coerce').dt.strftime('%Y-%m-%d').fillna('')
            mask = due_on_clean != ""
            df.loc[mask, 'Remarks Office'] = df.loc[mask, 'Remarks Office'].fillna('').astype(str).str.strip() + " " + due_on_clean
            df['Remarks Office'] = df['Remarks Office'].str.replace(r'\s+', ' ', regex=True).str.strip()

        # 7. Conversions (Cal to Months, Mandatory, Dates)
        if 'Int. Cal.' in df.columns:
            df['Int. Cal.'] = (pd.to_numeric(df['Int. Cal.'], errors='coerce') / 365 * 12).round(0).astype('Int64')
        if 'Mandatory' in df.columns:
            df['Mandatory'] = df['Mandatory'].astype(str).str.strip().replace('M', 'TRUE').replace('nan', '')
        for col in ['Due date', 'Start date']:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d').replace('NaT', '')

        # 8. Comp / Appl / N/A Logic
        if 'Pn' in df.columns:
            df['Comp.'] = ""
            df.loc[df['Pn'].notna() & (df['Pn'].astype(str).str.strip() != ""), 'Comp.'] = "TRUE"

        if 'Remarks Office' in df.columns:
            df['Appl.'] = ""
            keywords_appl = 'REPETITIVE|VERIFY IMMDT|COMPLY WITH|AS REQD'
            is_app = df['Remarks Office'].astype(str).str.contains(keywords_appl, case=False, na=False) | \
                     df['Remarks Office'].isna() | (df['Remarks Office'].astype(str).str.strip() == "")
            df.loc[is_app, 'Appl.'] = "TRUE"
            df['N/A'] = ""
            df.loc[df['Appl.'] != "TRUE", 'N/A'] = "TRUE"

        # 9. Component Control Math & Highlighting Prep
        modified_rows_green_cell = []
        if 'ATA / Task No.' in df.columns and 'Ac. TT LSV' in df.columns and 'Item TT' in df.columns:
            mask_cc = df['ATA / Task No.'].astype(str).str.contains('~COMPONENT CONTROL', na=False)
            item_tt_numeric = pd.to_numeric(df['Item TT'], errors='coerce')
            df.loc[mask_cc, 'Ac. TT LSV'] = ref_b7_value - item_tt_numeric
            modified_rows_green_cell = df.index[mask_cc].tolist()
            df['Ac. TT LSV'] = df['Ac. TT LSV'].apply(format_to_hhmm)

        # 10. Identify Highlights (Yellow and Green Rows)
        yellow_keywords = '8.33 KHZ CONVERSION|deleted|AIRCON/SVC.RECHARGE.ANN|NO LONGER A REQUIREMENT|WHEEL AND BRAKE CONFIRMATION|SANITISE AIRCRAFT|PROPELLER BALANCING'
        yellow_mask = (df['ATA / Task No.'].astype(str).str.contains(yellow_keywords, case=False, na=False)) | \
                      (df['Description'].astype(str).str.contains(yellow_keywords, case=False, na=False))
        modified_rows_yellow = df.index[yellow_mask].tolist()

        exclude_keywords = 'AD|SB|SIL'
        contains_mandatory = (df['ATA / Task No.'].astype(str).str.contains(exclude_keywords, case=False, na=False)) | \
                             (df['Description'].astype(str).str.contains(exclude_keywords, case=False, na=False))
        modified_rows_green_row = df.index[~contains_mandatory].tolist()

        # 11. Add Ref_B7 Column (Fixed for Strict Typing)
        while len(df.columns) < 34: df[f"Extra_{len(df.columns)}"] = None
        df.iloc[0, 33] = ref_b7_value
        cols = list(df.columns); cols[33] = "Ref_B7"; df.columns = cols

        # 12. Final Clean (Zeros)
        zero_clean_cols = ['Int. FH', 'Int. FC', 'Int. Cal.', 'Item FC LSV']
        for col in zero_clean_cols:
            if col in df.columns: df[col] = df[col].replace({0: pd.NA, 0.0: pd.NA, "0": pd.NA})

        # --- STYLING AND EXCEL GENERATION ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        
        output.seek(0)
        wb = load_workbook(output)
        ws = wb.active
        
        # Fills
        green_solid = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
        green_light = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        
        ac_tt_col_idx = None
        for cell in ws[1]:
            if cell.value in edited_headers: cell.fill = blue_fill
            if cell.value == "Ac. TT LSV": ac_tt_col_idx = cell.column

        # Priority Painting
        for row_idx in modified_rows_green_row:
            for cell in ws[row_idx + 2]: cell.fill = green_light
        for row_idx in modified_rows_yellow:
            for cell in ws[row_idx + 2]: cell.fill = yellow_fill
        if ac_tt_col_idx:
            for row_idx in modified_rows_green_cell:
                ws.cell(row=row_idx + 2, column=ac_tt_col_idx).fill = green_solid

        # Auto-Resize
        for col in ws.columns:
            max_len = 0
            for cell in col:
                if cell.value: max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 2

        final_buffer = io.BytesIO()
        wb.save(final_buffer)
        
        st.success("Transformation Complete!")
        
        # Logging
        if log_to_github(uploaded_file.name):
            st.info("Conversion recorded in GitHub log.csv")

        st.download_button(
            label="📥 Download Transformed Excel",
            data=final_buffer.getvalue(),
            file_name=export_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

with tab2:
    st.subheader("Persistent Audit Log (log.csv)")
    if st.button("Refresh Log"):
        st.rerun()
    try:
        token = st.secrets["GITHUB_TOKEN"]
        repo_name = st.secrets["REPO_NAME"]
        g = Github(token)
        repo = g.get_repo(repo_name)
        file_content = repo.get_contents("log.csv")
        log_df = pd.read_csv(io.StringIO(file_content.decoded_content.decode("utf-8")))
        st.dataframe(log_df, use_container_width=True)
    except:
        st.info("Log file is being generated or repository connection is pending.")
