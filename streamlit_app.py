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
        if minutes == 60: hours += 1; minutes = 0
        return f"{hours}:{minutes:02d}"
    except: return val

def log_to_github(filename, process_type, username):
    try:
        token = st.secrets["GITHUB_TOKEN"]
        repo_name = st.secrets["REPO_NAME"]
        g = Github(token)
        repo = g.get_repo(repo_name)
        
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # 5-column format: Timestamp, User, Filename, Process, Status
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
    """Returns True if the user had the correct password."""
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if st.session_state["authenticated"]:
        return True

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

# --- MAIN APP START ---
st.set_page_config(page_title="CAFAM Export Toolkit", page_icon="✈️", layout="wide")

if check_password():
    # User is logged in, show the app
    current_user = st.session_state["username"]
    st.sidebar.write(f"👤 Logged in as: **{current_user}**")
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

        if uploaded_handover:
            if st.button("🚀 Process"):
                base_name = os.path.splitext(uploaded_handover.name)[0]
                export_filename = f"{base_name} CAFAM Export.xlsx"

                # ... (Transformation Logic - Same as before) ...
                df_raw = pd.read_excel(uploaded_handover, engine='openpyxl')
                ref_b7_value = pd.to_numeric(df_raw.iloc[5, 1], errors='coerce') if not df_raw.empty else 0
                df = df_raw.drop(df_raw.index[0:8]).reset_index(drop=True)
                
                # [Amalgamation, Renaming, Highlighting logic goes here...]
                # (Assuming full logic from previous turns is pasted here)

                # Style and Save logic...
                
                st.success("Transformation to Recurring Maintenance Complete!")
                # LOGGING: Now includes the username
                if log_to_github(uploaded_handover.name, "Recurring Maintenance", current_user):
                    st.info("Logged")
                
                # Download button...

    # ---------------------------------------------------------
    # TAB 2: MODLIST ITEMS
    # ---------------------------------------------------------
    with tab2:
        st.subheader("Transform to Modlist Items")
        uploaded_modlist = st.file_uploader("Upload CAFAM File", type=["xlsx"], key="modlist")

        if uploaded_modlist:
            if st.button("🚀 Process", key="proc_mod"):
                # ... (Modlist Logic - Same as before) ...
                
                st.success("Modlist Transformation Complete!")
                # LOGGING: Now includes the username
                if log_to_github(uploaded_modlist.name, "Modlist Items", current_user):
                    st.info("Logged")
                
                # Download button...

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
