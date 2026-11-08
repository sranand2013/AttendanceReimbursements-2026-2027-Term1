import streamlit as st
import pandas as pd
import sqlite3
import os
from pathlib import Path
import io
import json
import datetime

# -----------------------------------------------------------------------------
# Authentication Setup
# -----------------------------------------------------------------------------
st.set_page_config(page_title="Attendance Data Manager", layout="wide")

if 'authenticated' not in st.session_state:
    st.session_state['authenticated'] = False

if not st.session_state['authenticated']:
    st.title("🔒 Login Required")
    st.write("Please enter your credentials to access the Attendance Manager.")
    
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submit = st.form_submit_button("Login")
        
        if submit:
            if username == "admin" and password == "xyz1234":
                st.session_state['authenticated'] = True
                st.rerun()
            else:
                st.error("Invalid username or password.")
                
    st.stop()

# -----------------------------------------------------------------------------
# Database Setup (Updated Schema)
# -----------------------------------------------------------------------------
def get_db_connection(db_name="attendance_records_v2.db"):
    conn = sqlite3.connect(db_name, check_same_thread=False)
    return conn

def init_db(conn):
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            usn TEXT,
            class_name TEXT,
            date TEXT,
            time TEXT,
            event_name TEXT,
            file_name TEXT,
            file_path TEXT,
            file_timestamp TEXT,
            extra_data TEXT,
            UNIQUE(usn, date, time, event_name)
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS skipped_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT,
            file_path TEXT,
            file_timestamp TEXT,
            reason TEXT,
            UNIQUE(file_path)
        )
    ''')
    conn.commit()

# Helper to expand JSON data into pandas columns for UI/Exports
def expand_extra_data(df):
    if 'Extra Data' in df.columns:
        def safe_parse(val):
            try:
                return json.loads(val) if pd.notna(val) else {}
            except:
                return {}
        
        expanded_df = pd.json_normalize(df['Extra Data'].apply(safe_parse))
        df = pd.concat([df.drop(columns=['Extra Data']), expanded_df], axis=1)
    return df

# -----------------------------------------------------------------------------
# Data Processing Engine
# -----------------------------------------------------------------------------
def process_excel_file(file_obj, file_name_str, file_path_str, file_timestamp, conn):
    COLUMN_ALIASES = {
        'NAME': ['NAME', 'STUDENT NAME', 'FULL NAME', 'STUDENT'],
        'USN': ['USN', 'REGISTER NUMBER', 'REGISTRATION NUMBER', 'REG NO', 'ROLL NO', 'REGISTER NO'],
        'CLASS': ['CLASS', 'SECTION', 'SEM AND SEC', 'COURSE', 'PROGRAM', 'BATCH', 'CLASS AND SECTION'],
        'DATE': ['DATE', 'DATES', 'DAY', 'DATE OF REIMBURSEMENT'],
        'TIME': ['TIME', 'TIMINGS', 'TIMING', 'FULLDAY', 'FULL DAY', 'HOURS', 'DURATION', 'SESSION'],
        'EVENT': ['REASON', 'EVENT', 'EVENT NAME', 'PROGRAMME', 'PURPOSE', 'REASON FOR'],
        'SIGNATORY': [
            'SIGNATORY', 'SIGNATURE', 'AUTHORISED', 'AUTHORIZED', 'AUTHRISED', 
            'VERIFIED BY', 'APPROVED BY', 'COORDINATOR', 'ASSIGNED TEACHER', 
            'ASSIGNED NAME', 'SARITHA'
        ]
    }
    
    IGNORE_KEYWORDS = [
        'SL.NO', 'S.NO', 'SL NO', 'S.N', 'SL. NO', 'NAN', 
        'TELEPHONE', 'MOBILE', 'PASSPORT', 'FATHER', 'MOTHER', 
        'YEAR', 'TOKEN', 'UTR', 'TRANSACTION', 'REFERENCE'
    ]

    def is_useful_column(col_name):
        col_upper = str(col_name).strip().upper()
        if any(kw in col_upper for kw in IGNORE_KEYWORDS):
            return False
        if col_upper.startswith('UNNAMED'):
            return False
        return True

    # Helper functions for auto-correcting swapped Name/USN data
    def has_numbers(s):
        return any(c.isdigit() for c in s)
    def has_letters(s):
        return any(c.isalpha() for c in s)

    try:
        sheets_dict = pd.read_excel(file_obj, header=None, sheet_name=None)
        total_records_added = 0
        cursor = conn.cursor()

        for sheet_name, temp_df in sheets_dict.items():
            if temp_df.empty:
                continue
                
            header_row_index = -1
            standardized_rename_map = {}
            
            # Find the header row using a smarter 2-PASS match system
            for i, row in temp_df.iterrows():
                row_vals = [str(x).strip().upper() for x in row.values]
                current_map = {}
                
                # PASS 1: Strict EXACT matching (prevents "STUDENT REG NO" mapping to "NAME")
                for orig_col_name in row_vals:
                    if orig_col_name in ('NAN', 'NONE', ''): continue
                    for std_col, aliases in COLUMN_ALIASES.items():
                        if std_col not in current_map.values():
                            if any(alias == orig_col_name for alias in aliases):
                                current_map[orig_col_name] = std_col
                                break
                
                # PASS 2: Substring matching (catches leftover messy headers)
                for orig_col_name in row_vals:
                    if orig_col_name in ('NAN', 'NONE', '') or orig_col_name in current_map: continue
                    for std_col, aliases in COLUMN_ALIASES.items():
                        if std_col not in current_map.values():
                            if any(alias in orig_col_name for alias in aliases):
                                current_map[orig_col_name] = std_col
                                break
                
                if 'NAME' in current_map.values() and 'USN' in current_map.values():
                    header_row_index = i
                    actual_cols = [str(x).strip().upper() for x in temp_df.iloc[i].values]
                    for col in actual_cols:
                        if col in current_map:
                            standardized_rename_map[col] = current_map[col]
                    break
                    
            if header_row_index == -1:
                continue
                
            df = temp_df.iloc[header_row_index + 1:].copy()
            df.columns = [str(c).strip().upper() for c in temp_df.iloc[header_row_index].values]
            
            # Remove duplicate columns to prevent pandas crashes
            df = df.loc[:, ~df.columns.duplicated()].copy()
            
            threshold = max(1, int(0.1 * len(df)))
            df.dropna(axis=1, thresh=threshold, inplace=True)
            
            extra_columns = [col for col in df.columns if col not in standardized_rename_map and is_useful_column(col)]
            
            df.rename(columns=standardized_rename_map, inplace=True)
            
            for req_col in ['CLASS', 'DATE', 'TIME', 'EVENT']:
                if req_col not in df.columns:
                    df[req_col] = "Not Specified"
                    
            if 'NAME' not in df.columns or 'USN' not in df.columns:
                continue
                
            df = df.dropna(subset=['NAME', 'USN'])
            df = df[df['NAME'].astype(str).str.strip().str.upper() != 'NAME']
            df = df[~df['USN'].astype(str).str.strip().str.upper().isin(['USN', 'REGISTRATION NUMBER', 'REGISTER NUMBER', 'REG NO'])]
            
            for _, row in df.iterrows():
                name_val = str(row['NAME']).strip()
                usn_val = str(row['USN']).strip()
                
                # SMART AUTO-CORRECT: If Name has numbers and USN has only letters, they were typed in the wrong columns!
                if has_numbers(name_val) and not has_numbers(usn_val) and has_letters(usn_val):
                    name_val, usn_val = usn_val, name_val # Swap them back to correct order
                
                extra_data_dict = {col: row[col] for col in extra_columns if pd.notna(row[col]) and str(row[col]).strip() != ""}
                
                if 'SIGNATORY' in df.columns and pd.notna(row.get('SIGNATORY')) and str(row['SIGNATORY']).strip() != "":
                    extra_data_dict['Authorized Signatory'] = str(row['SIGNATORY']).strip()
                    
                extra_data_json = json.dumps(extra_data_dict, default=str)

                try:
                    cursor.execute('''
                        INSERT OR IGNORE INTO attendance (
                            name, usn, class_name, date, time, event_name, 
                            file_name, file_path, file_timestamp, extra_data
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        name_val, 
                        usn_val, 
                        str(row.get('CLASS', 'Not Specified')).strip(), 
                        str(row.get('DATE', 'Not Specified')).strip(), 
                        str(row.get('TIME', 'Not Specified')).strip(), 
                        str(row.get('EVENT', 'Not Specified')).strip(),
                        f"{file_name_str} [{sheet_name}]",
                        file_path_str,
                        file_timestamp,
                        extra_data_json
                    ))
                    if cursor.rowcount > 0:
                        total_records_added += 1
                except sqlite3.Error:
                    pass 
                    
        conn.commit()
        return True, f"Processed all sheets. Added {total_records_added} new records."
        
    except Exception as e:
        return False, f"File Error: {str(e)}"

# -----------------------------------------------------------------------------
# Streamlit UI
# -----------------------------------------------------------------------------
col1, col2 = st.columns([0.9, 0.1])
with col1:
    st.title("Student Attendance & Reimbursement Manager")
with col2:
    if st.button("Logout"):
        st.session_state['authenticated'] = False
        st.rerun()

conn = get_db_connection()
init_db(conn)

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📂 Folder Scan", 
    "📥 Drag & Drop", 
    "🔍 Student Query",
    "📅 Event Search",
    "📊 Reports & Downloads",
    "⚠️ Skipped Files"
])

# TAB 1: Folder Scan
with tab1:
    st.markdown("### Scan Local Directory Recursively")
    folder_path = st.text_input("Enter the absolute root folder path (e.g., C:/Data/Attendance):")
    
    if st.button("Start Folder Scan"):
        if os.path.isdir(folder_path):
            root_dir = Path(folder_path)
            
            excel_files = [
                p for p in root_dir.rglob("*") 
                if p.is_file() and p.suffix.lower() in ['.xlsx', '.xls'] and not p.name.startswith("~$")
            ]
            
            if not excel_files:
                st.warning("No Excel files found.")
            else:
                progress_text = st.empty()
                bar = st.progress(0)
                
                total_files = len(excel_files)
                success_count = 0
                
                for i, file_path in enumerate(excel_files):
                    rel_path = str(file_path.relative_to(root_dir))
                    mtime = os.path.getmtime(file_path)
                    sys_timestamp = datetime.datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
                    
                    success, msg = process_excel_file(file_path, file_path.name, rel_path, sys_timestamp, conn)
                    
                    if success and "Added 0 new records" not in msg:
                        success_count += 1
                    else:
                        reason = msg if not success else "File processed but contained no valid matching student records."
                        st.warning(f"Skipped {file_path.name}: {reason}")
                        
                        try:
                            conn.cursor().execute('''
                                INSERT OR REPLACE INTO skipped_files (
                                    file_name, file_path, file_timestamp, reason
                                ) VALUES (?, ?, ?, ?)
                            ''', (file_path.name, rel_path, sys_timestamp, reason))
                            conn.commit()
                        except sqlite3.Error:
                            pass
                        
                    progress_text.text(f"Processing {i+1}/{total_files}: {file_path.name}")
                    bar.progress((i + 1) / total_files)
                    
                st.success(f"Scan complete! Processed {success_count} valid files.")
        else:
            st.error("Invalid directory path.")

# TAB 2: Drag and Drop Upload
with tab2:
    st.markdown("### Upload Individual Excel Files")
    uploaded_files = st.file_uploader("Drag and drop files here", type=['xlsx', 'xls'], accept_multiple_files=True)
    
    if st.button("Process Uploaded Files") and uploaded_files:
        with st.spinner("Processing files and all sheets..."):
            upload_timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S (Upload Time)')
            for file in uploaded_files:
                success, msg = process_excel_file(file, file.name, f"Direct Upload: {file.name}", upload_timestamp, conn)
                if success:
                    st.success(f"{file.name}: {msg}")
                else:
                    st.error(f"{file.name}: Failed - {msg}")

# TAB 3: Student Query
with tab3:
    st.markdown("### Lookup Student Records")
    search_term = st.text_input("Enter Student Name or USN:")
    
    if search_term:
        query = f"%{search_term}%"
        df_results = pd.read_sql_query('''
            SELECT name as Name, usn as USN, class_name as Class, date as Date, time as Time, 
                   event_name as Event, file_name as 'File Name', file_path as 'Source Path', 
                   file_timestamp as 'File Timestamp', extra_data as 'Extra Data'
            FROM attendance 
            WHERE name LIKE ? OR usn LIKE ?
            ORDER BY file_timestamp DESC, date DESC
        ''', conn, params=(query, query))
        
        if not df_results.empty:
            df_results = expand_extra_data(df_results)
            st.dataframe(df_results, use_container_width=True)
            st.info(f"Total entries found: {len(df_results)}")
        else:
            st.warning("No records found.")

# TAB 4: Event Search
with tab4:
    st.markdown("### Search by Event or File Name")
    
    df_events = pd.read_sql_query("SELECT DISTINCT event_name FROM attendance WHERE event_name != 'Not Specified'", conn)
    event_list = df_events['event_name'].dropna().tolist()
    
    df_files = pd.read_sql_query("SELECT DISTINCT file_name FROM attendance", conn)
    file_list = df_files['file_name'].dropna().tolist()
    
    combined_list = sorted(list(set(event_list + file_list)))
    
    selected_event = st.selectbox("Start typing or select an Event/Reason or File Name:", [""] + combined_list, index=0)
    
    if selected_event:
        df_event_results = pd.read_sql_query('''
            SELECT name as Name, usn as USN, class_name as Class, date as Date, time as Time, 
                   event_name as Event, file_name as 'File Name', file_path as 'Source Path',
                   file_timestamp as 'File Timestamp', extra_data as 'Extra Data'
            FROM attendance
            WHERE event_name = ? OR file_name = ?
            ORDER BY file_timestamp DESC, date DESC, name ASC
        ''', conn, params=(selected_event, selected_event))
        
        if not df_event_results.empty:
            df_event_results = expand_extra_data(df_event_results)
            st.dataframe(df_event_results, use_container_width=True)
            st.info(f"Total students found for '{selected_event}': {len(df_event_results)}")
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_event_results.to_excel(writer, index=False, sheet_name='Event Report')
            
            safe_filename = "".join([c for c in selected_event if c.isalpha() or c.isdigit() or c==' ']).rstrip()
            
            st.download_button(
                label=f"📥 Download '{selected_event}' Report",
                data=buffer.getvalue(),
                file_name=f"Event_Report_{safe_filename[:20]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# TAB 5: Reports & Downloads
with tab5:
    st.markdown("### 📊 Export Consolidated Database")
    st.write("Download every single record currently stored in the database, ordered by the latest received files.")
    
    if st.button("Download All Consolidated Data"):
        all_data_query = '''
            SELECT name as Name, usn as USN, class_name as Class, date as Date, time as Time, 
                   event_name as Event, file_name as 'File Name', file_path as 'Source Path',
                   file_timestamp as 'File Timestamp', extra_data as 'Extra Data'
            FROM attendance
            ORDER BY file_timestamp DESC
        '''
        df_all = pd.read_sql_query(all_data_query, conn)
        
        if not df_all.empty:
            df_all = expand_extra_data(df_all)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_all.to_excel(writer, index=False, sheet_name='All Data')
            
            st.download_button(
                label="📥 Download All Data (Excel)",
                data=buffer.getvalue(),
                file_name="All_Consolidated_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Database is empty. Please upload data first.")
            
    st.markdown("---")
    
    st.markdown("### 🏆 Reimbursement Analysis")
    st.write("Generate a report of students with the highest number of attendance days.")
    
    if st.button("Generate Reimbursement Report"):
        report_query = '''
            SELECT 
                usn as USN,
                name as Name,
                class_name as Class,
                COUNT(DISTINCT date) as 'Reimbursement Days Allowed',
                COUNT(id) as 'Total Session Entries'
            FROM attendance
            GROUP BY usn
            ORDER BY 'Reimbursement Days Allowed' DESC, 'Total Session Entries' DESC
        '''
        df_report = pd.read_sql_query(report_query, conn)
        
        if not df_report.empty:
            st.dataframe(df_report, use_container_width=True)
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_report.to_excel(writer, index=False, sheet_name='Reimbursement Report')
            
            st.download_button(
                label="📥 Download Reimbursement Report",
                data=buffer.getvalue(),
                file_name="Reimbursement_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Database is empty. Please upload data first.")

# TAB 6: Skipped Files Log
with tab6:
    st.markdown("### ⚠️ Skipped Files Log")
    st.write("Files that failed to process or contained no extractable data during scans.")
    
    df_skipped = pd.read_sql_query('''
        SELECT file_name as 'File Name', 
               file_path as 'Source Path', 
               file_timestamp as 'File Timestamp', 
               reason as 'Reason' 
        FROM skipped_files 
        ORDER BY file_timestamp DESC
    ''', conn)
    
    if not df_skipped.empty:
        st.dataframe(df_skipped, use_container_width=True)
        
        buffer_skipped = io.BytesIO()
        with pd.ExcelWriter(buffer_skipped, engine='xlsxwriter') as writer:
            df_skipped.to_excel(writer, index=False, sheet_name='Skipped Files')
        
        col_dl, col_clear = st.columns([0.2, 0.8])
        with col_dl:
            st.download_button(
                label="📥 Download Log",
                data=buffer_skipped.getvalue(),
                file_name="Skipped_Files_Log.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col_clear:
            if st.button("🗑️ Clear Skipped Files Log"):
                conn.cursor().execute("DELETE FROM skipped_files")
                conn.commit()
                st.success("Log cleared!")
                st.rerun()
    else:
        st.success("No skipped files recorded! All files processed contained valid data.")
