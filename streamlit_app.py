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

def log_to_github(filename, process_type, username):
    try:
        token = st.secrets["GITHUB_TOKEN"]
        repo_name = st.secrets["REPO_NAME"]
        g = Github(token)
        repo = g.get_repo(repo_name)
        
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        new_row = f"{timestamp},{username},{filename},{process_type},Success\n"
        
        file_path = "log.csv"
        file_content = repo.get_contents(file_path)
        current_data = file_content.decoded_content.decode("utf-8")
        updated_data = current_data + new_row
        repo.update_file(file_path, f"Log: {filename} by {username}", updated_data, file_content.sha)
        return True
    except Exception as e:
        st.error(f"GitHub Logging failed: {e}")
        return False

# --- AUTHENTICATION GATEKEEPER ---

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if st.session_state["authenticated"]:
        return True

    _, col2, _ = st.columns([1, 2, 1])
    with col2:
        st.subheader("🔑 User Login")
        user = st.text_input("Username")
        password = st.text_input("Password", type="password")
        
        if st.button("Login"):
            if user in st.secrets["passwords"] and password == st.secrets["passwords"][user]:
                st.session_state["authenticated"] = True
                st.session_state["username"] = user
                st.rerun()
            else:
                st.error("Invalid username or password")
    return False

# --- MAIN APP ---

st.set_page_config(page_title="CAFAM Export Toolkit", page_icon="✈️", layout="wide")

# Inject Custom CSS to remove all Streamlit default branding, footer, and top menu
hide_streamlit_style = """
<style>
/* Hide the Streamlit main toolbar/menu */
[data-testid="stToolbar"] {visibility: hidden !important;}
/* Hide the "Made with Streamlit" footer */
footer {visibility: hidden !important;}
/* Remove default top margin for a completely clean layout */
header {visibility: hidden !important;}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

if check_password():
    current_user = st.session_state["username"]
    st.sidebar.write(f"👤 User: **{current_user}**")
    if st.sidebar.button("Logout"):
        st.session_state["authenticated"] = False
        st.rerun()

    st.title("✈️ CAFAM Export Toolkit")
    tab1, tab2, tab3 = st.tabs(["Recurring Maintenance", "Modlist Items", "Permanent Audit Log"])

    # ---------------------------------------------------------
    # TAB 1: RECURRING MAINTENANCE
    # ---------------------------------------------------------
    with tab1:
        st.subheader("Transform Records Handover File")
        uploaded_handover = st.file_uploader("Upload Maintenance File", type=["xlsx"], key="handover")

        if "handover_output" not in st.session_state:
            st.session_state.handover_output = None
            st.session_state.handover_filename = ""

        if uploaded_handover:
            if "last_file_h" not in st.session_state or st.session_state.last_file_h != uploaded_handover.name:
                st.session_state.handover_output = None
                st.session_state.last_file_h = uploaded_handover.name

            if st.button("🚀 Process", key="btn_h"):
                with st.spinner("Processing..."):
                    base_name = os.path.splitext(uploaded_handover.name)[0]
                    export_filename = f"{base_name} CAFAM Export.xlsx"

                    # 1. Load Data
                    df_raw = pd.read_excel(uploaded_handover, engine='openpyxl')
                    ref_b7_value = pd.to_numeric(df_raw.iloc[5, 1], errors='coerce') if not df_raw.empty else 0
                    df = df_raw.drop(df_raw.index[0:8]).reset_index(drop=True)

                    # 2. Amalgamate Description
                    col_d, col_e = df.iloc[:, 3].fillna('').astype(str), df.iloc[:, 4].fillna('').astype(str)
                    others = df.iloc[:, 5:8].fillna('').astype(str).agg(' '.join, axis=1)
                    description_data = (col_d + ": " + col_e + " " + others).str.replace(r'\s+', ' ', regex=True).str.strip().str.lstrip(': ')

                    # 3. Rename Map
                    df = df.drop(columns=df.columns[3:8])
                    df.insert(3, 'Description', description_data)
                    rename_map = {
                        "DIRCTVE": "ATA / Task No.", "GROUP": "Group", "PARTNO": "Pn", "SERIAL": "Sn",
                        "SCHD_HRS": "Int. FH", "SCHD_LDG": "Int. FC", "SCHD_DAYS": "Int. Cal.",
                        "DATE_DUEBY": "Due date", "MANDTRY": "Mandatory", "REASON": "Remarks Office",
                        "ACTION": "Action", "DATESAT": "Start date", "AC_HRS_SAT": "Ac. TT LSV",
                        "TTSN": "Item TT", "TCSN": "Cycl. SN", "AC_LDG_SAT": "Item FC LSV"
                    }
                    df = df.rename(columns=rename_map)

                    # 4. Mandatory, Comp, Appl, N/A Logic
                    if 'Mandatory' in df.columns:
                        df['Mandatory'] = df['Mandatory'].astype(str).str.strip().replace('M', 'TRUE').replace('nan', '')

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

                    if 'DATE_DUEON' in df.columns:
                        due_on_clean = pd.to_datetime(df['DATE_DUEON'], errors='coerce').dt.strftime('%Y-%m-%d').fillna('')
                        mask = due_on_clean != ""
                        df.loc[mask, 'Remarks Office'] = df.loc[mask, 'Remarks Office'].fillna('').astype(str).str.strip() + " " + due_on_clean
                        df['Remarks Office'] = df['Remarks Office'].str.replace(r'\s+', ' ', regex=True).str.strip()

                    # 5. Dates & Math
                    if 'Int. Cal.' in df.columns:
                        df['Int. Cal.'] = (pd.to_numeric(df['Int. Cal.'], errors='coerce') / 365 * 12).round(0).astype('Int64')
                    
                    for col in ['Due date', 'Start date']:
                        if col in df.columns:
                            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d').replace('NaT', '')

                    modified_rows_green_cell = []
                    if 'ATA / Task No.' in df.columns and 'Ac. TT LSV' in df.columns and 'Item TT' in df.columns:
                        mask_cc = df['ATA / Task No.'].astype(str).str.contains('~COMPONENT CONTROL', na=False)
                        df.loc[mask_cc, 'Ac. TT LSV'] = ref_b7_value - pd.to_numeric(df['Item TT'], errors='coerce')
                        modified_rows_green_cell = df.index[mask_cc].tolist()
                        df['Ac. TT LSV'] = df['Ac. TT LSV'].apply(format_to_hhmm)

                    # 6. Highlights Prep
                    yellow_keywords = '8.33 KHZ CONVERSION|deleted|AIRCON/SVC.RECHARGE.ANN|NO LONGER A REQUIREMENT|WHEEL AND BRAKE CONFIRMATION|SANITISE AIRCRAFT|PROPELLER BALANCING'
                    yellow_mask = (df['ATA / Task No.'].astype(str).str.contains(yellow_keywords, case=False, na=False)) | (df['Description'].astype(str).str.contains(yellow_keywords, case=False, na=False))
                    modified_rows_yellow = df.index[yellow_mask].tolist()

                    exclude_keywords = 'AD|SB|SIL'
                    contains_mandatory = (df['ATA / Task No.'].astype(str).str.contains(exclude_keywords, case=False, na=False)) | (df['Description'].astype(str).str.contains(exclude_keywords, case=False, na=False))
                    modified_rows_green_row = df.index[~contains_mandatory].tolist()

                    # 7. Column Padding (Ref_B7 at Index 33)
                    while len(df.columns) < 34: df[f"Extra_{len(df.columns)}"] = None
                    df.iloc[0, 33] = ref_b7_value
                    cols = list(df.columns); cols[33] = "Ref_B7"; df.columns = cols

                    zero_clean_cols = ['Int. FH', 'Int. FC', 'Int. Cal.', 'Item FC LSV']
                    for col in zero_clean_cols:
                        if col in df.columns: df[col] = df[col].replace({0: pd.NA, 0.0: pd.NA, "0": pd.NA})

                    # 8. Style and Save
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer: df.to_excel(writer, index=False)
                    output.seek(0)
                    wb = load_workbook(output)
                    ws = wb.active
                    
                    g_solid, g_light = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid"), PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
                    b_fill, y_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid"), PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                    
                    edited_headers = list(rename_map.values()) + ['Description', 'Comp.', 'Appl.', 'N/A', 'Ref_B7']
                    ac_tt_col_idx = None
                    for cell in ws[1]:
                        if cell.value in edited_headers: cell.fill = b_fill
                        if cell.value == "Ac. TT LSV": ac_tt_col_idx = cell.column

                    for r in modified_rows_green_row:
                        for c in ws[r + 2]: c.fill = g_light
                    for r in modified_rows_yellow:
                        for c in ws[r + 2]: c.fill = y_fill
                    if ac_tt_col_idx:
                        for r in modified_rows_green_cell: ws.cell(row=r + 2, column=ac_tt_col_idx).fill = g_solid

                    for col in ws.columns:
                        m_len = 0
                        for cell in col:
                            if cell.value: m_len = max(m_len, len(str(cell.value)))
                        ws.column_dimensions[get_column_letter(col[0].column)].width = m_len + 2

                    final_buffer = io.BytesIO()
                    wb.save(final_buffer)
                    st.session_state.handover_output = final_buffer.getvalue()
                    st.session_state.handover_filename = export_filename
                    
                    st.success("Transformation to Recurring Maintenance Complete!")
                    if log_to_github(uploaded_handover.name, "Recurring Maintenance", current_user):
                        st.info("Logged")

            if st.session_state.handover_output:
                st.download_button("📥 Download Export", data=st.session_state.handover_output, file_name=st.session_state.handover_filename)

    # ---------------------------------------------------------
    # TAB 2: MODLIST ITEMS
    # ---------------------------------------------------------
    with tab2:
        st.subheader("Transform to Modlist Items")
        uploaded_modlist = st.file_uploader("Upload CAFAM File", type=["xlsx"], key="modlist")

        if "mod_output" not in st.session_state:
            st.session_state.mod_output, st.session_state.mod_filename = None, ""

        if uploaded_modlist:
            if "last_file_m" not in st.session_state or st.session_state.last_file_m != uploaded_modlist.name:
                st.session_state.mod_output = None
                st.session_state.last_file_m = uploaded_modlist.name

            if st.button("🚀 Process", key="btn_m"):
                with st.spinner("Reformatting..."):
                    base_name = os.path.splitext(uploaded_modlist.name)[0]
                    modlist_filename = f"{base_name}_REFORMATTED.xlsx"

                    wb = load_workbook(uploaded_modlist)
                    ws = wb.active
                    l_blue, l_green = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid'), PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
                    d_yellow, c_align = PatternFill(start_color='FFFF99', end_color='FFFF99', fill_type='solid'), Alignment(horizontal='center', vertical='center')
                    green_ids = ["AD", "SB", "MSB", "SIL", "CSB"]
                    yellow_ids = ["8.33 KHZ CONVERSION", "AIRCON/SVC.RECHARGE.ANN", "NO LONGER A REQUIREMENT", "CAA SD-2024/001 V.2", "PROPELLER BALANCING", "DELETED", "WHEEL AND BRAKE CONFIRMATION: Please verify the wheel/ brake system"]

                    cell_b7 = ws['B7'].value
                    ws.delete_rows(2, 8)

                    for row in range(2, ws.max_row + 1):
                        d_val, e_val = str(ws.cell(row=row, column=4).value or "").strip(), str(ws.cell(row=row, column=5).value or "").strip()
                        prefix = f"{d_val}: {e_val}" if d_val and e_val else (d_val or e_val)
                        other_parts = [str(ws.cell(row=row, column=col).value).strip() for col in range(6, 9) if ws.cell(row=row, column=col).value is not None]
                        ws.cell(row=row, column=4).value = " ".join([prefix] + other_parts).strip()

                    ws.delete_cols(5, 4)
                    ws.insert_cols(3); ws.cell(row=1, column=3).value = "AD"; ws.cell(row=1, column=3).fill = l_blue
                    ws.insert_cols(4); ws.cell(row=1, column=4).value = "Mod"; ws.cell(row=1, column=4).fill = l_blue
                    ws.insert_cols(5); ws.cell(row=1, column=5).value = "Main Type"; ws.cell(row=1, column=5).fill = l_blue

                    h_map = {"DIRCTVE": "SB/SL", "DESCR": "Description", "GROUP": "Valid", "REASON": "Method of compl.", "DATESAT": "C/W Date", "AC_HRS_SAT": "C/W FH", "AC_LDG_SAT": "C/W FC", "DATE_DUEBY": "Remarks", "ALTREFNO": "AD Foreign", "JOBNO": "C/W WO", "ACTION": "C/W"}
                    c_idx = {h: None for h in ["SB/SL", "AD", "Main Type", "C/W FH", "Description", "C/W Date", "C/W", "C/W WO"]}
                    for cell in ws[1]:
                        if cell.value:
                            val = str(cell.value).strip()
                            if val in h_map:
                                cell.value = h_map[val]; cell.fill = l_blue; val = cell.value
                            if val in c_idx: c_idx[val] = cell.column

                    for r in range(2, ws.max_row + 1):
                        cl = {k: ws.cell(row=r, column=v) if v else None for k, v in c_idx.items()}
                        v_sb_o = str(cl["SB/SL"].value or "").strip()
                        if v_sb_o.upper().startswith("AD"): cl["AD"].value = cl["SB/SL"].value; cl["SB/SL"].value = None
                        if cl["AD"].value: cl["Main Type"].value = "AD"
                        elif cl["SB/SL"].value: cl["Main Type"].value = "SB/SL"
                        if cl["C/W FH"].value is not None: cl["C/W FH"].value = format_to_hhmm(cl["C/W FH"].value)
                        if cl["C/W Date"].value: cl["C/W"].value = "TRUE"; cl["C/W Date"].number_format = 'yyyy-mm-dd'
                        if cl["C/W WO"].value: cl["C/W WO"].value = str(cl["C/W WO"].value).strip().rstrip('/')
                        v_s, v_a, v_d = str(cl["SB/SL"].value or "").upper(), str(cl["AD"].value or "").upper(), str(cl["Description"].value or "").upper()
                        if any(t in v_s or t in v_d for t in yellow_ids):
                            for c in ws[r]: c.fill = d_yellow
                        elif any(t in v_s or t in v_a or t in v_d for t in green_ids):
                            for c in ws[r]: c.fill = l_green

                    for row in ws.iter_rows():
                        for c in row: c.alignment = c_align
                    ws['AL2'] = cell_b7
                    ws['AL2'].alignment = c_align

                    for col in ws.columns:
                        m_len = 0
                        for cell in col:
                            if cell.value: m_len = max(m_len, len(str(cell.value)))
                        ws.column_dimensions[col[0].column_letter].width = m_len + 2

                    mod_buf = io.BytesIO()
                    wb.save(mod_buf)
                    st.session_state.mod_output = mod_buf.getvalue()
                    st.session_state.mod_filename = modlist_filename
                    st.success("Modlist Transformation Complete!")
                    if log_to_github(uploaded_modlist.name, "Modlist Items", current_user):
                        st.info("Logged")

            if st.session_state.mod_output:
                st.download_button("📥 Download Modlist Export", data=st.session_state.mod_output, file_name=st.session_state.mod_filename)

    # ---------------------------------------------------------
    # TAB 3: AUDIT LOG
    # ---------------------------------------------------------
    with tab3:
        st.subheader("Permanent Audit Log (GitHub)")
        if st.button("🔄 Force Refresh Log"): 
            st.cache_data.clear()
            st.rerun()
        try:
            token, repo_name = st.secrets["GITHUB_TOKEN"], st.secrets["REPO_NAME"]
            g = Github(token); repo = g.get_repo(repo_name)
            file_content = repo.get_contents("log.csv")
            log_df = pd.read_csv(io.StringIO(file_content.decoded_content.decode("utf-8")))
            st.dataframe(log_df.iloc[::-1], use_container_width=True, hide_index=True)
        except Exception as e:
            st.error(f"Audit Log Error: {e}")
