import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from streamlit_gsheets import GSheetsConnection # New Connection Library

# --- HELPER FUNCTIONS ---
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

# --- STREAMLIT UI ---
st.set_page_config(page_title="CAFAM Transformer", page_icon="✈️")
st.title("✈️ CAFAM Record Transformer")

# --- GOOGLE SHEETS CONNECTION ---
# This looks for your Sheet URL in the Streamlit "Secrets" dashboard
conn = st.connection("gsheets", type=GSheetsConnection)

tab1, tab2 = st.tabs(["Transformer", "Permanent Activity Log"])

with tab1:
    uploaded_file = st.file_uploader("Upload Handover File", type=["xlsx"])

    if uploaded_file:
        base_name = os.path.splitext(uploaded_file.name)[0]
        export_filename = f"{base_name} CAFAM Export.xlsx"

        # [REDACTED FOR BREVITY: YOUR 1-13 STEPS OF TRANSFORMATION LOGIC HERE]
        # (Ensure you keep the exact logic we built in the previous version)
        
        # --- (Your full logic from the previous script goes here) ---
        # ... (Step 1 through 12) ...

        # --- AFTER SUCCESSFUL PROCESSING ---
        st.success("Transformation Complete!")
        
        # LOG TO GOOGLE SHEETS
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            # We create a small dataframe for the new row
            new_log = pd.DataFrame([{"Timestamp": timestamp, "Filename": uploaded_file.name, "Status": "Success"}])
            
            # Read existing data
            existing_data = conn.read(ttl=0) # ttl=0 ensures we get fresh data
            
            # Append and update
            updated_log = pd.concat([existing_data, new_log], ignore_index=True)
            conn.update(data=updated_log)
            st.info("Activity recorded in permanent log.")
        except Exception as e:
            st.warning(f"Transformation worked, but couldn't reach the log: {e}")

        # (Insert Download Button Logic here)

with tab2:
    st.subheader("Permanent Conversion History")
    try:
        # Fetch the log from Google Sheets
        log_data = conn.read(ttl="10m") # Refresh every 10 mins
        st.table(log_data)
    except:
        st.info("The log is currently empty or the connection isn't set up yet.")
