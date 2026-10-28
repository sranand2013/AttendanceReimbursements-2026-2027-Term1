import streamlit as st
import pandas as pd
import sqlite3
import os
from pathlib import Path
import io

# -----------------------------------------------------------------------------
# Database Setup
# -----------------------------------------------------------------------------
def get_db_connection(db_name="attendance_records.db"):
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
            UNIQUE(usn, date, time, event_name)
        )
    ''')
    conn.commit()

# -----------------------------------------------------------------------------
# Data Processing Engine
# -----------------------------------------------------------------------------
def process_excel_file(file_obj, file_name_str, file_path_str, conn):
    """Reads the excel file, dynamically maps headers, cleans internal headers, and inserts to DB."""
    
    COLUMN_ALIASES = {
        'NAME': ['NAME', 'STUDENT NAME', 'FULL NAME', 'STUDENT'],
        'USN': ['USN', 'REGISTER NUMBER', 'REGISTRATION NUMBER', 'REG NO', 'ROLL NO', 'REGISTER NO'],
        'CLASS': ['CLASS', 'SECTION', 'SEM AND SEC', 'COURSE', 'PROGRAM', 'BATCH', 'CLASS AND SECTION'],
        'DATE': ['DATE', 'DATES', 'DAY', 'DATE OF REIMBURSEMENT'],
        'TIME': ['TIME', 'TIMINGS', 'TIMING', 'FULLDAY', 'FULL DAY', 'HOURS', 'DURATION', 'SESSION'],
        'EVENT': ['REASON', 'EVENT', 'EVENT NAME', 'PROGRAMME', 'PURPOSE']
    }

    try:
        # Read without headers first to locate the actual header row
        temp_df = pd.read_excel(file_obj, header=None)
        
        header_row_index = -1
        standardized_rename_map = {}
        
        # Scan row by row to find headers
        for i, row in temp_df.iterrows():
            row_vals = [str(x).strip().upper() for x in row.values]
            current_map = {}
            
            for orig_col_name in row_vals:
                if orig_col_name == 'NAN': continue
                
                # Check against aliases
                for std_col, aliases in COLUMN_ALIASES.items():
                    if any(alias in orig_col_name for alias in aliases):
                        if std_col not in current_map.values():
                            current_map[orig_col_name] = std_col
                            break
            
            # A valid header row MUST contain at least a Name and USN
            if 'NAME' in current_map.values() and 'USN' in current_map.values():
                header_row_index = i
                actual_cols = [str(x).strip().upper() for x in temp_df.iloc[i].values]
                for col in actual_cols:
                    if col in current_map:
                        standardized_rename_map[col] = current_map[col]
                break
                
        if header_row_index == -1:
            return False, "Could not identify header row (Missing Name or USN variations)."
            
        # Re-read with the correct header
        df = pd.read_excel(file_obj, header=header_row_index)
        
        # Uppercase all columns to match the map
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        # Rename columns to standard format
        df.rename(columns=standardized_rename_map, inplace=True)
        
        # Handle optional/missing columns by adding them with default values
        for req_col in ['CLASS', 'DATE', 'TIME', 'EVENT']:
            if req_col not in df.columns:
                df[req_col] = "Not Specified"
                
        # Drop rows missing primary identifiers (handles entirely blank rows)
        df = df.dropna(subset=['NAME', 'USN'])
        
        # Filter Internal Headers: Remove rows where the data itself is just the repeating header name
        df = df[df['NAME'].astype(str).str.strip().str.upper() != 'NAME']
        df = df[~df['USN'].astype(str).str.strip().str.upper().isin(['USN', 'REGISTRATION NUMBER', 'REGISTER NUMBER', 'REG NO'])]
        
        cursor = conn.cursor()
        records_added = 0
        for _, row in df.iterrows():
            try:
                cursor.execute('''
                    INSERT OR IGNORE INTO attendance (name, usn, class_name, date, time, event_name, file_name, file_path)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    str(row['NAME']).strip(), 
                    str(row['USN']).strip(), 
                    str(row.get('CLASS', 'Not Specified')).strip(), 
                    str(row.get('DATE', 'Not Specified')).strip(), 
                    str(row.get('TIME', 'Not Specified')).strip(), 
                    str(row.get('EVENT', 'Not Specified')).strip(),
                    file_name_str,
                    file_path_str
                ))
                if cursor.rowcount > 0:
                    records_added += 1
            except sqlite3.Error:
                pass # Ignored due to UNIQUE constraint
                
        conn.commit()
        return True, f"Successfully added {records_added} new records."
        
    except Exception as e:
        return False, f"File Error: {str(e)}"

# -----------------------------------------------------------------------------
# Streamlit UI
# -----------------------------------------------------------------------------
st.set_page_config(page_title="Attendance Data Manager", layout="wide")
st.title("Student Attendance & Reimbursement Manager")

# Initialize DB on first run
conn = get_db_connection()
init_db(conn)

# UI Layout using Tabs
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📂 Folder Scan", 
    "📥 Drag & Drop", 
    "🔍 Student Query",
    "📅 Event Search",
    "📊 Reports & Downloads"
])

# TAB 1: Recursive Folder Scan
with tab1:
    st.markdown("### Scan Local Directory Recursively")
    folder_path = st.text_input("Enter the absolute root folder path (e.g., C:/Data/Attendance):")
    
    if st.button("Start Folder Scan"):
        if os.path.isdir(folder_path):
            root_dir = Path(folder_path)
            
            # Case-insensitive recursive scan that ignores temporary ~$ lock files
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
                    success, msg = process_excel_file(file_path, file_path.name, rel_path, conn)
                    if success:
                        success_count += 1
                    else:
                        st.warning(f"Skipped {file_path.name}: {msg}")
                        
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
        with st.spinner("Processing files..."):
            for file in uploaded_files:
                success, msg = process_excel_file(file, file.name, f"Direct Upload: {file.name}", conn)
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
            SELECT name as Name, usn as USN, class_name as Class, date as Date, time as Time, event_name as Event, file_name as 'File Name', file_path as 'Source Path' 
            FROM attendance 
            WHERE name LIKE ? OR usn LIKE ?
            ORDER BY date DESC
        ''', conn, params=(query, query))
        
        if not df_results.empty:
            st.dataframe(df_results, use_container_width=True)
            st.info(f"Total entries found: {len(df_results)}")
        else:
            st.warning("No records found.")

# TAB 4: Event Search
with tab4:
    st.markdown("### Search by Event or File Name")
    
    # Fetch unique event names (ignoring defaults)
    df_events = pd.read_sql_query("SELECT DISTINCT event_name FROM attendance WHERE event_name != 'Not Specified'", conn)
    event_list = df_events['event_name'].dropna().tolist()
    
    # Fetch unique file names
    df_files = pd.read_sql_query("SELECT DISTINCT file_name FROM attendance", conn)
    file_list = df_files['file_name'].dropna().tolist()
    
    # Combine both lists and remove duplicates
    combined_list = sorted(list(set(event_list + file_list)))
    
    selected_event = st.selectbox("Start typing or select an Event/Reason or File Name:", [""] + combined_list, index=0)
    
    if selected_event:
        # Search WHERE selected item matches EITHER the event_name OR the file_name
        df_event_results = pd.read_sql_query('''
            SELECT name as Name, usn as USN, class_name as Class, date as Date, time as Time, event_name as Event, file_name as 'File Name', file_path as 'Source Path'
            FROM attendance
            WHERE event_name = ? OR file_name = ?
            ORDER BY date DESC, name ASC
        ''', conn, params=(selected_event, selected_event))
        
        if not df_event_results.empty:
            st.dataframe(df_event_results, use_container_width=True)
            st.info(f"Total students found for '{selected_event}': {len(df_event_results)}")
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_event_results.to_excel(writer, index=False, sheet_name='Event Report')
            
            # Sanitize the file name for download
            safe_filename = "".join([c for c in selected_event if c.isalpha() or c.isdigit() or c==' ']).rstrip()
            
            st.download_button(
                label=f"📥 Download '{selected_event}' Report",
                data=buffer.getvalue(),
                file_name=f"Event_Report_{safe_filename[:20]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# TAB 5: Reports & Downloads
with tab5:
    st.markdown("### Reimbursement Analysis")
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
                label="📥 Download Excel Report",
                data=buffer.getvalue(),
                file_name="Reimbursement_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Database is empty. Please upload data first.")