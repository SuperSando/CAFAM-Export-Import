import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime
from github import Github
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils import get_column_letter

# --- SHARED HELPER FUNCTIONS ---

def format_to_hhmm(val):
    """Universal helper for HH:MM conversion"""
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

def log_to_github(filename, process_type):
    """Centralized GitHub Logging"""
    try:
        token = st.secrets["GITHUB_TOKEN"]
        repo_name = st.secrets["REPO_NAME"]
        g = Github(token)
        repo = g.get_repo(repo_name)
        
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        new_row = f"{timestamp},{filename},{process_type},Success\n"
        
        file_path = "log.csv"
        try:
            file_content = repo.get_contents(file_path)
            current_data = file_content.decoded_content.decode("utf-8")
            updated_data = current_data + new_row
            repo.update_file(file_path, f"Log: {filename}", updated_data, file_content.sha)
        except:
            header = "Timestamp,Filename,Process,Status\n"
            repo.create_file(file_path, "Initial log creation", header + new_row)
        return True
    except Exception as e:
        st.error(f"GitHub Logging failed: {e}")
        return False

# --- APP CONFIG ---
st.set_page_config(page_title="RGV Aviation Toolkit", page_icon="✈️", layout="wide")
st.title("✈️ RGV Maintenance Toolkit")

# Tab Names Updated
tab1, tab2, tab3 = st.tabs(["Recurring Maintenance", "Modlist Items", "Permanent Audit Log"])

# ---------------------------------------------------------
# TAB 1: RECURRING MAINTENANCE
# ---------------------------------------------------------
with tab1:
    st.subheader("Transform Recurring Maintenance File")
    uploaded_handover = st.file_uploader("Upload Maintenance File", type=["xlsx"], key="handover")

    if uploaded_handover:
        base_name = os.path.splitext(uploaded_handover.name)[0]
        export_filename = f"{base_name} CAFAM Export.xlsx"

        df_raw = pd.read_excel(uploaded_handover, engine='openpyxl')
        ref_b7_value = pd.to_numeric(df_raw.iloc[5, 1], errors='coerce') if not df_raw.empty else 0
        df = df_raw.drop(df_raw.index[0:8]).reset_index(drop=True)

        col_d, col_e = df.iloc[:, 3].fillna('').astype(str), df.iloc[:, 4].fillna('').astype(str)
        others = df.iloc[:, 5:8].fillna('').astype(str).agg(' '.join, axis=1)
        description_data = (col_d + ": " + col_e + " " + others).str.replace(r'\s+', ' ', regex=True).str.strip().str.lstrip(': ')

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

        if 'DATE_DUEON' in df.columns:
            due_on_clean = pd.to_datetime(df['DATE_DUEON'], errors='coerce').dt.strftime('%Y-%m-%d').fillna('')
            mask = due_on_clean != ""
            df.loc[mask, 'Remarks Office'] = df.loc[mask, 'Remarks Office'].fillna('').astype(str).str.strip() + " " + due_on_clean
            df['Remarks Office'] = df['Remarks Office'].str.replace(r'\s+', ' ', regex=True).str.strip()

        if 'Int. Cal.' in df.columns:
            df['Int. Cal.'] = (pd.to_numeric(df['Int. Cal.'], errors='coerce') / 365 * 12).round(0).astype('Int64')
        if 'Mandatory' in df.columns:
            df['Mandatory'] = df['Mandatory'].astype(str).str.strip().replace('M', 'TRUE').replace('nan', '')
        for col in ['Due date', 'Start date']:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d').replace('NaT', '')

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

        modified_rows_green_cell = []
        if 'ATA / Task No.' in df.columns and 'Ac. TT LSV' in df.columns and 'Item TT' in df.columns:
            mask_cc = df['ATA / Task No.'].astype(str).str.contains('~COMPONENT CONTROL', na=False)
            item_tt_numeric = pd.to_numeric(df['Item TT'], errors='coerce')
            df.loc[mask_cc, 'Ac. TT LSV'] = ref_b7_value - item_tt_numeric
            modified_rows_green_cell = df.index[mask_cc].tolist()
            df['Ac. TT LSV'] = df['Ac. TT LSV'].apply(format_to_hhmm)

        yellow_keywords = '8.33 KHZ CONVERSION|deleted|AIRCON/SVC.RECHARGE.ANN|NO LONGER A REQUIREMENT|WHEEL AND BRAKE CONFIRMATION|SANITISE AIRCRAFT|PROPELLER BALANCING'
        yellow_mask = (df['ATA / Task No.'].astype(str).str.contains(yellow_keywords, case=False, na=False)) | \
                      (df['Description'].astype(str).str.contains(yellow_keywords, case=False, na=False))
        modified_rows_yellow = df.index[yellow_mask].tolist()

        exclude_keywords = 'AD|SB|SIL'
        contains_mandatory = (df['ATA / Task No.'].astype(str).str.contains(exclude_keywords, case=False, na=False)) | \
                             (df['Description'].astype(str).str.contains(exclude_keywords, case=False, na=False))
        modified_rows_green_row = df.index[~contains_mandatory].tolist()

        while len(df.columns) < 34: df[f"Extra_{len(df.columns)}"] = None
        df.iloc[0, 33] = ref_b7_value
        cols = list(df.columns); cols[33] = "Ref_B7"; df.columns = cols

        zero_clean_cols = ['Int. FH', 'Int. FC', 'Int. Cal.', 'Item FC LSV']
        for col in zero_clean_cols:
            if col in df.columns: df[col] = df[col].replace({0: pd.NA, 0.0: pd.NA, "0": pd.NA})

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer: df.to_excel(writer, index=False)
        output.seek(0)
        wb = load_workbook(output)
        ws = wb.active
        
        green_solid, green_light = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid"), PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        blue_fill, yellow_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid"), PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        
        ac_tt_col_idx = None
        for cell in ws[1]:
            if cell.value in edited_headers: cell.fill = blue_fill
            if cell.value == "Ac. TT LSV": ac_tt_col_idx = cell.column

        for row_idx in modified_rows_green_row:
            for cell in ws[row_idx + 2]: cell.fill = green_light
        for row_idx in modified_rows_yellow:
            for cell in ws[row_idx + 2]: cell.fill = yellow_fill
        if ac_tt_col_idx:
            for row_idx in modified_rows_green_cell: ws.cell(row=row_idx + 2, column=ac_tt_col_idx).fill = green_solid

        for col in ws.columns:
            max_len = 0
            for cell in col:
                if cell.value: max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 2

        final_buffer = io.BytesIO()
        wb.save(final_buffer)
        
        st.success("Recurring Maintenance Transformation Complete!")
        if log_to_github(uploaded_handover.name, "Recurring Maintenance"):
            st.info("Logged to GitHub.")
        st.download_button("📥 Download Maintenance Export", data=final_buffer.getvalue(), file_name=export_filename)

# ---------------------------------------------------------
# TAB 2: MODLIST ITEMS
# ---------------------------------------------------------
with tab2:
    st.subheader("Transform Modlist Items")
    uploaded_modlist = st.file_uploader("Upload Modlist File", type=["xlsx"], key="modlist")

    if uploaded_modlist:
        base_name = os.path.splitext(uploaded_modlist.name)[0]
        modlist_export_filename = f"{base_name}_REFORMATTED.xlsx"

        wb = load_workbook(uploaded_modlist)
        ws = wb.active

        light_blue_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
        light_green_fill = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
        darker_yellow_fill = PatternFill(start_color='FFFF99', end_color='FFFF99', fill_type='solid')
        center_align = Alignment(horizontal='center', vertical='center')

        green_identifiers = ["AD", "SB", "MSB", "SIL", "CSB"]
        yellow_identifiers = [
            "8.33 KHZ CONVERSION", "AIRCON/SVC.RECHARGE.ANN", "NO LONGER A REQUIREMENT",
            "CAA SD-2024/001 V.2", "PROPELLER BALANCING", "DELETED",
            "WHEEL AND BRAKE CONFIRMATION: Please verify the wheel/ brake system"
        ]

        cell_value_b7 = ws['B7'].value
        ws.delete_rows(2, 8)

        for row in range(2, ws.max_row + 1):
            d_val = str(ws.cell(row=row, column=4).value or "").strip()
            e_val = str(ws.cell(row=row, column=5).value or "").strip()
            prefix = f"{d_val}: {e_val}" if d_val and e_val else (d_val or e_val)
            other_parts = [str(ws.cell(row=row, column=col).value).strip() for col in range(6, 9) if ws.cell(row=row, column=col).value is not None]
            ws.cell(row=row, column=4).value = " ".join([prefix] + other_parts).strip()

        ws.delete_cols(5, 4)
        ws.insert_cols(3); ws.cell(row=1, column=3).value = "AD"; ws.cell(row=1, column=3).fill = light_blue_fill
        ws.insert_cols(4); ws.cell(row=1, column=4).value = "Mod"; ws.cell(row=1, column=4).fill = light_blue_fill
        ws.insert_cols(5); ws.cell(row=1, column=5).value = "Main Type"; ws.cell(row=1, column=5).fill = light_blue_fill

        header_replacements = {
            "DIRCTVE": "SB/SL", "DESCR": "Description", "GROUP": "Valid", "REASON": "Method of compl.",
            "DATESAT": "C/W Date", "AC_HRS_SAT": "C/W FH", "AC_LDG_SAT": "C/W FC", "DATE_DUEBY": "Remarks",
            "ALTREFNO": "AD Foreign", "JOBNO": "C/W WO", "ACTION": "C/W"
        }
        col_sb_sl, col_ad, col_mt, col_cw_fh, col_descr, col_cw_date, col_cw_action, col_cw_wo = [None]*8
        for cell in ws[1]:
            if cell.value:
                hdr = str(cell.value).strip()
                if hdr in header_replacements:
                    cell.value = header_replacements[hdr]
                    cell.fill = light_blue_fill
                    hdr = cell.value
                if hdr == "SB/SL": col_sb_sl = cell.column
                if hdr == "AD": col_ad = cell.column
                if hdr == "Main Type": col_mt = cell.column
                if hdr == "C/W FH": col_cw_fh = cell.column
                if hdr == "Description": col_descr = cell.column
                if hdr == "C/W Date": col_cw_date = cell.column
                if hdr == "C/W": col_cw_action = cell.column
                if hdr == "C/W WO": col_cw_wo = cell.column

        for row_idx in range(2, ws.max_row + 1):
            cell_sb = ws.cell(row=row_idx, column=col_sb_sl)
            cell_ad = ws.cell(row=row_idx, column=col_ad)
            cell_mt = ws.cell(row=row_idx, column=col_mt)
            cell_fh = ws.cell(row=row_idx, column=col_cw_fh)
            cell_ds = ws.cell(row=row_idx, column=col_descr)
            cell_dt = ws.cell(row=row_idx, column=col_cw_date)
            cell_cw = ws.cell(row=row_idx, column=col_cw_action)
            cell_wo = ws.cell(row=row_idx, column=col_cw_wo)
            
            val_sb_orig = str(cell_sb.value or "").strip()
            if val_sb_orig.upper().startswith("AD"):
                cell_ad.value = cell_sb.value; cell_sb.value = None

            if cell_ad.value: cell_mt.value = "AD"
            elif cell_sb.value: cell_mt.value = "SB/SL"

            if cell_fh.value is not None:
                cell_fh.value = format_to_hhmm(cell_fh.value)

            if cell_dt.value:
                if cell_cw: cell_cw.value = "TRUE"
                cell_dt.number_format = 'yyyy-mm-dd'

            if cell_wo.value:
                cell_wo.value = str(cell_wo.value).strip().rstrip('/')

            v_sb, v_ad, v_ds = str(cell_sb.value or "").upper(), str(cell_ad.value or "").upper(), str(cell_ds.value or "").upper()
            if any(t in v_sb or t in v_ds for t in yellow_identifiers):
                for c in ws[row_idx]: c.fill = darker_yellow_fill
            elif any(t in v_sb or t in v_ad or t in v_ds for t in green_identifiers):
                for c in ws[row_idx]: c.fill = light_green_fill

        for row in ws.iter_rows():
            for c in row: c.alignment = center_align
        ws['AL2'] = cell_value_b7
        ws['AL2'].alignment = center_align

        for col in ws.columns:
            max_len = 0
            for cell in col:
                if cell.value: max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col[0].column_letter].width = max_len + 2

        mod_buffer = io.BytesIO()
        wb.save(mod_buffer)
        st.success("Modlist Transformation Complete!")
        if log_to_github(uploaded_modlist.name, "Modlist Items"):
            st.info("Logged to GitHub.")
        st.download_button("📥 Download Modlist Export", data=mod_buffer.getvalue(), file_name=modlist_export_filename)

# ---------------------------------------------------------
# TAB 3: AUDIT LOG
# ---------------------------------------------------------
with tab3:
    st.subheader("Permanent Audit Log (GitHub)")
    if st.button("Refresh Log"): st.rerun()
    try:
        token = st.secrets["GITHUB_TOKEN"]
        repo_name = st.secrets["REPO_NAME"]
        g = Github(token)
        repo = g.get_repo(repo_name)
        file_content = repo.get_contents("log.csv")
        log_df = pd.read_csv(io.StringIO(file_content.decoded_content.decode("utf-8")))
        st.dataframe(log_df, use_container_width=True)
    except:
        st.info("Log is pending first conversion.")
