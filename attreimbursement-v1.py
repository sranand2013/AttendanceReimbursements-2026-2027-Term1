import streamlit as st
import pandas as pd
import os
from pathlib import Path
import io
import zipfile

# -----------------------------------------------------------------------------
# 1. Page & Authentication Setup (Always renders first)
# -----------------------------------------------------------------------------
st.set_page_config(page_title="Attendance & Reimbursement Manager", layout="wide")

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
# 2. Safe Dependency Check & Turso Database Setup
# -----------------------------------------------------------------------------
try:
    import libsql_client
except ImportError as e:
    st.error(f"Missing dependency `libsql-client`: {e}")
    st.info("Ensure `libsql-client` is listed in `requirements.txt` on GitHub and click **Reboot app** in Streamlit Cloud.")
    st.stop()

@st.cache_resource
def get_turso_client():
    """Creates a cached synchronous HTTP client for Turso DB 'eshaaletursocmslb'."""
    if "turso" not in st.secrets:
        raise KeyError("Missing `[turso]` section in Streamlit Cloud Secrets.")
    url = st.secrets["turso"]["url"].replace("libsql://", "https://").strip()
    auth_token = st.secrets["turso"]["auth_token"].strip()
    return libsql_client.create_client_sync(url=url, auth_token=auth_token)

def init_db(client):
    client.execute('''
        CREATE TABLE IF NOT EXISTS attendance_reimbursement (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            usn TEXT,
            class_name TEXT,
            date TEXT,
            time TEXT,
            event_name TEXT,
            term TEXT,
            file_name TEXT,
            file_path TEXT,
            UNIQUE(usn, date, time, event_name, term)
        )
    ''')

def run_select_query(client, query, params=None):
    """Executes a SELECT query on Turso and returns a Pandas DataFrame."""
    rs = client.execute(query, params or [])
    if not rs.rows:
        return pd.DataFrame(columns=list(rs.columns))
    return pd.DataFrame([list(row) for row in rs.rows], columns=list(rs.columns))

def get_available_terms(client):
    """Fetches distinct terms from the database for UI dropdowns."""
    df_terms = run_select_query(
        client, 
        "SELECT DISTINCT term FROM attendance_reimbursement WHERE term IS NOT NULL AND term != '' ORDER BY term DESC"
    )
    return df_terms['term'].tolist() if not df_terms.empty else []

# -----------------------------------------------------------------------------
# 3. Data Processing Engine (Returns explicit status: 'processed', 'skipped', 'error')
# -----------------------------------------------------------------------------
def process_excel_file(file_obj, file_name_str, file_path_str, default_term, client):
    """
    Reads all sheets, maps headers dynamically, assigns Term, and batch-upserts to Turso.
    Returns: (status, records_count, message) where status in {'processed', 'skipped', 'error'}
    """
    COLUMN_ALIASES = {
        'NAME': ['NAME', 'STUDENT NAME', 'FULL NAME', 'STUDENT'],
        'USN': ['USN', 'REGISTER NUMBER', 'REGISTRATION NUMBER', 'REG NO', 'ROLL NO', 'REGISTER NO'],
        'CLASS': ['CLASS', 'SECTION', 'SEM AND SEC', 'COURSE', 'PROGRAM', 'BATCH', 'CLASS AND SECTION'],
        'DATE': ['DATE', 'DATES', 'DAY', 'DATE OF REIMBURSEMENT'],
        'TIME': ['TIME', 'TIMINGS', 'TIMING', 'FULLDAY', 'FULL DAY', 'HOURS', 'DURATION', 'SESSION'],
        'EVENT': ['REASON', 'EVENT', 'EVENT NAME', 'PROGRAMME', 'PURPOSE'],
        'TERM': ['TERM', 'ACADEMIC TERM', 'SEMESTER TERM']
    }

    try:
        sheets_dict = pd.read_excel(file_obj, header=None, sheet_name=None)
        total_sheets = len(sheets_dict)
        valid_sheets = 0
        statements = []

        for sheet_name, temp_df in sheets_dict.items():
            if temp_df.empty:
                continue
                
            header_row_index = -1
            standardized_rename_map = {}
            
            for i, row in temp_df.iterrows():
                row_vals = [str(x).strip().upper() for x in row.values]
                current_map = {}
                
                for orig_col_name in row_vals:
                    if orig_col_name in ('NAN', 'NONE', ''):
                        continue
                    
                    for std_col, aliases in COLUMN_ALIASES.items():
                        if any(alias in orig_col_name for alias in aliases):
                            if std_col not in current_map.values():
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
            df.rename(columns=standardized_rename_map, inplace=True)
            
            for req_col in ['CLASS', 'DATE', 'TIME', 'EVENT']:
                if req_col not in df.columns:
                    df[req_col] = "Not Specified"
            
            if 'TERM' not in df.columns:
                df['TERM'] = default_term.strip()
            else:
                df['TERM'] = df['TERM'].fillna(default_term.strip())
                
            if 'NAME' not in df.columns or 'USN' not in df.columns:
                continue
                
            df = df.dropna(subset=['NAME', 'USN'])
            df = df[df['NAME'].astype(str).str.strip().str.upper() != 'NAME']
            df = df[~df['USN'].astype(str).str.strip().str.upper().isin(
                ['USN', 'REGISTRATION NUMBER', 'REGISTER NUMBER', 'REG NO']
            )]
            
            if df.empty:
                continue

            valid_sheets += 1
            for _, row in df.iterrows():
                row_term = str(row.get('TERM', default_term)).strip()
                if not row_term or row_term.upper() in ('NAN', 'NONE'):
                    row_term = default_term.strip()

                statements.append(
                    libsql_client.Statement(
                        '''
                        INSERT OR REPLACE INTO attendance_reimbursement 
                        (name, usn, class_name, date, time, event_name, term, file_name, file_path)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''',
                        [
                            str(row['NAME']).strip(), 
                            str(row['USN']).strip(), 
                            str(row.get('CLASS', 'Not Specified')).strip(), 
                            str(row.get('DATE', 'Not Specified')).strip(), 
                            str(row.get('TIME', 'Not Specified')).strip(), 
                            str(row.get('EVENT', 'Not Specified')).strip(),
                            row_term,
                            f"{file_name_str} [{sheet_name}]",
                            file_path_str
                        ]
                    )
                )

        if not statements:
            return "skipped", 0, f"Skipped (no valid Name/USN student rows found across {total_sheets} sheet(s))."

        total_records_processed = 0
        chunk_size = 100
        for idx in range(0, len(statements), chunk_size):
            batch_results = client.batch(statements[idx:idx + chunk_size])
            total_records_processed += sum(res.rows_affected for res in batch_results)

        return "processed", total_records_processed, f"Synced {total_records_processed} records from {valid_sheets}/{total_sheets} sheet(s)."
        
    except Exception as e:
        return "error", 0, f"Error reading file: {str(e)}"

# -----------------------------------------------------------------------------
# 4. Streamlit UI (Authenticated Area)
# -----------------------------------------------------------------------------
col1, col2 = st.columns([0.9, 0.1])
with col1:
    st.title("Student Attendance & Reimbursement Manager")
    st.caption("Connected to Turso Database: `eshaaletursocmslb` | Table: `attendance_reimbursement`")
with col2:
    if st.button("Logout"):
        st.session_state['authenticated'] = False
        st.rerun()

try:
    client = get_turso_client()
    init_db(client)
    existing_terms = get_available_terms(client)
except Exception as db_err:
    st.error(f"Turso Database Connection Failed: {db_err}")
    st.info("Open **App Settings → Secrets** in Streamlit Cloud and verify `[turso]` has valid `url` (`https://...`) and `auth_token`.")
    st.stop()

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📥 Drag & Drop (Files / ZIP)", 
    "📂 Local Folder Scan", 
    "🔍 Student Query",
    "📅 Event Search",
    "📊 Reports & Downloads"
])

# -----------------------------------------------------------------------------
# TAB 1: Drag and Drop Upload (Detailed Scan & Skip Reporting)
# -----------------------------------------------------------------------------
with tab1:
    st.markdown("### Upload Excel Files or Zipped Folders")
    upload_term = st.text_input("Academic Term (e.g., 2026-27 Term1):", value="2026-27 Term1", key="upload_term")
    uploaded_files = st.file_uploader(
        "Drag and drop Excel files (.xlsx, .xls) or a compressed folder (.zip) here", 
        type=['xlsx', 'xls', 'zip'], 
        accept_multiple_files=True
    )
    
    if st.button("Process Uploaded Files") and uploaded_files:
        if not upload_term.strip():
            st.error("Please enter an Academic Term before uploading.")
        else:
            progress_text = st.empty()
            progress_bar = st.progress(0)

            # Collect all candidate items first so we have exact total counts
            excel_tasks = []       # (display_path, clean_name, file_bytes_or_obj)
            ignored_non_excel = [] # (display_path, reason)
            zip_errors = []

            for file in uploaded_files:
                if file.name.lower().endswith(".zip"):
                    try:
                        with zipfile.ZipFile(file) as z:
                            all_entries = [info for info in z.infolist() if not info.is_dir()]
                            for zip_info in all_entries:
                                entry_name = Path(zip_info.filename).name
                                # Ignore macOS metadata folders or Excel lock files
                                if zip_info.filename.startswith("__MACOSX/") or entry_name.startswith("~$"):
                                    ignored_non_excel.append((f"{file.name}/{zip_info.filename}", "Temporary/system metadata file ignored"))
                                elif zip_info.filename.lower().endswith(('.xlsx', '.xls')):
                                    with z.open(zip_info) as excel_bytes:
                                        excel_tasks.append((
                                            f"{file.name}/{zip_info.filename}",
                                            entry_name,
                                            io.BytesIO(excel_bytes.read())
                                        ))
                                else:
                                    ignored_non_excel.append((f"{file.name}/{zip_info.filename}", "Non-Excel file ignored"))
                    except zipfile.BadZipFile:
                        zip_errors.append((file.name, "Invalid or corrupted ZIP archive"))
                else:
                    if file.name.startswith("~$"):
                        ignored_non_excel.append((file.name, "Temporary Excel lock file ignored"))
                    else:
                        excel_tasks.append((f"Direct Upload: {file.name}", file.name, file))

            total_excel_files = len(excel_tasks)
            processed_logs = []
            skipped_logs = list(ignored_non_excel)
            error_logs = list(zip_errors)
            total_records_synced = 0

            if total_excel_files == 0:
                progress_bar.empty()
                progress_text.empty()
                st.session_state["upload_report"] = {
                    "term": upload_term,
                    "total_excel": 0,
                    "processed": [],
                    "skipped": skipped_logs,
                    "errors": error_logs,
                    "total_records": 0
                }
                st.rerun()

            for idx, (display_path, clean_name, file_obj) in enumerate(excel_tasks):
                progress_text.text(f"Scanning Excel file {idx + 1} of {total_excel_files}: {clean_name}")
                status, rec_count, msg = process_excel_file(
                    file_obj, clean_name, display_path, upload_term, client
                )
                if status == "processed":
                    processed_logs.append((display_path, msg))
                    total_records_synced += rec_count
                elif status == "skipped":
                    skipped_logs.append((display_path, msg))
                else:
                    error_logs.append((display_path, msg))

                progress_bar.progress((idx + 1) / total_excel_files)

            progress_text.empty()
            progress_bar.empty()

            st.session_state["upload_report"] = {
                "term": upload_term,
                "total_excel": total_excel_files,
                "processed": processed_logs,
                "skipped": skipped_logs,
                "errors": error_logs,
                "total_records": total_records_synced
            }
            st.rerun()

    # Render Persistent Scan Summary Report after rerun
    if "upload_report" in st.session_state:
        rep = st.session_state["upload_report"]
        p_count = len(rep["processed"])
        s_count = len(rep["skipped"])
        e_count = len(rep["errors"])
        t_excel = rep["total_excel"]

        st.markdown(f"#### 📋 Scan Summary for `{rep['term']}`")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Excel Files Scanned", f"{p_count + e_count} / {t_excel}")
        m2.metric("Files Imported", p_count)
        m3.metric("Skipped / Ignored", s_count)
        m4.metric("Records Synced", rep["total_records"])

        if p_count > 0:
            st.success(f"Successfully imported **{p_count} out of {t_excel}** Excel files ({rep['total_records']} total student entries synced).")

        if s_count > 0:
            st.warning(f"**{s_count} file(s) were skipped or ignored.** Expand below to inspect which files were skipped and why.")

        if e_count > 0:
            st.error(f"**{e_count} file(s) failed due to errors.** Expand below for details.")

        with st.expander("🔍 View Detailed File-by-File Scan Log", expanded=(s_count > 0 or e_count > 0)):
            if rep["errors"]:
                st.markdown("**❌ Failed Files:**")
                for fname, reason in rep["errors"]:
                    st.write(f"- `{fname}` — {reason}")
            if rep["skipped"]:
                st.markdown("**⚠️ Skipped / Ignored Files:**")
                for fname, reason in rep["skipped"]:
                    st.write(f"- `{fname}` — {reason}")
            if rep["processed"]:
                st.markdown("**✅ Processed Excel Files:**")
                for fname, reason in rep["processed"]:
                    st.write(f"- `{fname}` — {reason}")

# -----------------------------------------------------------------------------
# TAB 2: Recursive Folder Scan (For Local Machine Execution Only)
# -----------------------------------------------------------------------------
with tab2:
    st.markdown("### Scan Local Directory Recursively")
    st.info("Note: Local folder paths work only when running Streamlit locally on your PC. On Streamlit Cloud, compress your folder to `.zip` and use **Tab 1**.")
    
    col_t1_a, col_t1_b = st.columns([0.4, 0.6])
    with col_t1_a:
        scan_term = st.text_input("Academic Term (e.g., 2026-27 Term1):", value="2026-27 Term1", key="scan_term")
    with col_t1_b:
        folder_path = st.text_input("Absolute root folder path (e.g., C:/Data/Attendance):")
    
    if st.button("Start Folder Scan"):
        clean_path = folder_path.strip().strip('"').strip("'")
        if not scan_term.strip():
            st.error("Please enter an Academic Term before scanning.")
        elif clean_path and os.path.isdir(clean_path):
            root_dir = Path(clean_path)
            all_files = [p for p in root_dir.rglob("*") if p.is_file()]
            excel_files = []
            skipped_local = []

            for p in all_files:
                rel_p = str(p.relative_to(root_dir))
                if p.name.startswith("~$"):
                    skipped_local.append((rel_p, "Temporary Excel lock file ignored"))
                elif p.suffix.lower() in ['.xlsx', '.xls']:
                    excel_files.append((p, rel_p))
                else:
                    skipped_local.append((rel_p, "Non-Excel file ignored"))

            if not excel_files:
                st.warning(f"No Excel files found in the specified folder ({len(skipped_local)} non-Excel files ignored).")
            else:
                progress_text = st.empty()
                bar = st.progress(0)
                total_files = len(excel_files)
                processed_local = []
                error_local = []
                total_recs = 0
                
                for i, (file_path, rel_path) in enumerate(excel_files):
                    status, rec_count, msg = process_excel_file(file_path, file_path.name, rel_path, scan_term, client)
                    if status == "processed":
                        processed_local.append((rel_path, msg))
                        total_recs += rec_count
                    elif status == "skipped":
                        skipped_local.append((rel_path, msg))
                    else:
                        error_local.append((rel_path, msg))
                        
                    progress_text.text(f"Processing {i+1}/{total_files}: {file_path.name}")
                    bar.progress((i + 1) / total_files)
                    
                st.session_state["local_scan_report"] = {
                    "term": scan_term,
                    "total_excel": total_files,
                    "processed": processed_local,
                    "skipped": skipped_local,
                    "errors": error_local,
                    "total_records": total_recs
                }
                st.rerun()
        else:
            st.error(f"Invalid directory path: `{clean_path}`. Ensure the path exists on the machine running Streamlit, or upload a `.zip` file in Tab 1.")

    if "local_scan_report" in st.session_state:
        rep = st.session_state["local_scan_report"]
        st.success(f"Scanned {len(rep['processed'])}/{rep['total_excel']} Excel files ({rep['total_records']} records synced). Skipped/Ignored: {len(rep['skipped'])}.")
        with st.expander("View Local Scan Log", expanded=len(rep["skipped"]) > 0):
            for fname, reason in rep["skipped"]:
                st.write(f"- ⚠️ `{fname}` — {reason}")
            for fname, reason in rep["errors"]:
                st.write(f"- ❌ `{fname}` — {reason}")
            for fname, reason in rep["processed"]:
                st.write(f"- ✅ `{fname}` — {reason}")

# -----------------------------------------------------------------------------
# TAB 3: Student Query (Term-Wise)
# -----------------------------------------------------------------------------
with tab3:
    st.markdown("### Lookup Student Records")
    col_q1, col_q2 = st.columns([0.7, 0.3])
    with col_q1:
        search_term = st.text_input("Enter Student Name or USN:")
    with col_q2:
        selected_query_term = st.selectbox("Filter by Term:", ["All Terms"] + existing_terms, key="query_term_filter")
    
    if search_term:
        like_pattern = f"%{search_term}%"
        if selected_query_term == "All Terms":
            df_results = run_select_query(client, '''
                SELECT term as Term, name as Name, usn as USN, class_name as Class, 
                       date as Date, time as Time, event_name as Event, 
                       file_name as 'File Name', file_path as 'Source Path' 
                FROM attendance_reimbursement 
                WHERE name LIKE ? OR usn LIKE ?
                ORDER BY term DESC, date DESC
            ''', [like_pattern, like_pattern])
        else:
            df_results = run_select_query(client, '''
                SELECT term as Term, name as Name, usn as USN, class_name as Class, 
                       date as Date, time as Time, event_name as Event, 
                       file_name as 'File Name', file_path as 'Source Path' 
                FROM attendance_reimbursement 
                WHERE (name LIKE ? OR usn LIKE ?) AND term = ?
                ORDER BY date DESC
            ''', [like_pattern, like_pattern, selected_query_term])
        
        if not df_results.empty:
            st.dataframe(df_results, use_container_width=True)
            st.info(f"Total entries found: {len(df_results)}")
        else:
            st.warning("No records found.")

# -----------------------------------------------------------------------------
# TAB 4: Event Search (Term-Wise)
# -----------------------------------------------------------------------------
with tab4:
    st.markdown("### Search by Event or File Name")
    selected_event_term = st.selectbox("Filter Events by Term:", ["All Terms"] + existing_terms, key="event_term_filter")
    
    if selected_event_term == "All Terms":
        df_events = run_select_query(client, "SELECT DISTINCT event_name FROM attendance_reimbursement WHERE event_name != 'Not Specified'")
        df_files = run_select_query(client, "SELECT DISTINCT file_name FROM attendance_reimbursement")
    else:
        df_events = run_select_query(client, "SELECT DISTINCT event_name FROM attendance_reimbursement WHERE event_name != 'Not Specified' AND term = ?", [selected_event_term])
        df_files = run_select_query(client, "SELECT DISTINCT file_name FROM attendance_reimbursement WHERE term = ?", [selected_event_term])
        
    event_list = df_events['event_name'].dropna().tolist() if not df_events.empty else []
    file_list = df_files['file_name'].dropna().tolist() if not df_files.empty else []
    combined_list = sorted(list(set(event_list + file_list)))
    
    selected_event = st.selectbox("Start typing or select an Event/Reason or File Name:", [""] + combined_list, index=0)
    
    if selected_event:
        if selected_event_term == "All Terms":
            df_event_results = run_select_query(client, '''
                SELECT term as Term, name as Name, usn as USN, class_name as Class, 
                       date as Date, time as Time, event_name as Event, 
                       file_name as 'File Name', file_path as 'Source Path'
                FROM attendance_reimbursement
                WHERE event_name = ? OR file_name = ?
                ORDER BY term DESC, date DESC, name ASC
            ''', [selected_event, selected_event])
        else:
            df_event_results = run_select_query(client, '''
                SELECT term as Term, name as Name, usn as USN, class_name as Class, 
                       date as Date, time as Time, event_name as Event, 
                       file_name as 'File Name', file_path as 'Source Path'
                FROM attendance_reimbursement
                WHERE (event_name = ? OR file_name = ?) AND term = ?
                ORDER BY date DESC, name ASC
            ''', [selected_event, selected_event, selected_event_term])
        
        if not df_event_results.empty:
            st.dataframe(df_event_results, use_container_width=True)
            st.info(f"Total students found for '{selected_event}' ({selected_event_term}): {len(df_event_results)}")
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_event_results.to_excel(writer, index=False, sheet_name='Event Report')
            
            safe_filename = "".join([c for c in selected_event if c.isalpha() or c.isdigit() or c == ' ']).rstrip()
            term_tag = selected_event_term.replace(" ", "_")
            
            st.download_button(
                label=f"📥 Download '{selected_event}' Report ({selected_event_term})",
                data=buffer.getvalue(),
                file_name=f"Event_Report_{term_tag}_{safe_filename[:20]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# -----------------------------------------------------------------------------
# TAB 5: Term-Wise Reports & Downloads
# -----------------------------------------------------------------------------
with tab5:
    st.markdown("### 📊 Export Consolidated Database (Term-Wise)")
    selected_report_term = st.selectbox(
        "Select Term for Reports:", 
        ["All Terms"] + existing_terms, 
        key="report_term_filter"
    )
    
    # 1. Consolidated Data Download
    if selected_report_term == "All Terms":
        df_all = run_select_query(client, '''
            SELECT term as Term, name as Name, usn as USN, class_name as Class, 
                   date as Date, time as Time, event_name as Event, 
                   file_name as 'File Name', file_path as 'Source Path'
            FROM attendance_reimbursement
            ORDER BY term DESC, date DESC, name ASC
        ''')
    else:
        df_all = run_select_query(client, '''
            SELECT term as Term, name as Name, usn as USN, class_name as Class, 
                   date as Date, time as Time, event_name as Event, 
                   file_name as 'File Name', file_path as 'Source Path'
            FROM attendance_reimbursement
            WHERE term = ?
            ORDER BY date DESC, name ASC
        ''', [selected_report_term])
        
    if not df_all.empty:
        buffer_all = io.BytesIO()
        with pd.ExcelWriter(buffer_all, engine='xlsxwriter') as writer:
            df_all.to_excel(writer, index=False, sheet_name='Consolidated Data')
        
        term_slug = selected_report_term.replace(" ", "_")
        st.download_button(
            label=f"📥 Download Consolidated Data ({selected_report_term} - {len(df_all)} rows)",
            data=buffer_all.getvalue(),
            file_name=f"Consolidated_Data_{term_slug}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning(f"No records found for {selected_report_term}.")
            
    st.markdown("---")
    
    # 2. Term-Wise Reimbursement Analysis
    st.markdown("### 🏆 Term-Wise Reimbursement Analysis")
    st.write("Summary of student reimbursement days and total session entries for the selected term.")
    
    if selected_report_term == "All Terms":
        df_report = run_select_query(client, '''
            SELECT 
                term as Term,
                usn as USN,
                MAX(name) as Name,
                MAX(class_name) as Class,
                COUNT(DISTINCT date) as "Reimbursement Days Allowed",
                COUNT(id) as "Total Session Entries"
            FROM attendance_reimbursement
            GROUP BY term, usn
            ORDER BY term DESC, "Reimbursement Days Allowed" DESC, "Total Session Entries" DESC
        ''')
    else:
        df_report = run_select_query(client, '''
            SELECT 
                term as Term,
                usn as USN,
                MAX(name) as Name,
                MAX(class_name) as Class,
                COUNT(DISTINCT date) as "Reimbursement Days Allowed",
                COUNT(id) as "Total Session Entries"
            FROM attendance_reimbursement
            WHERE term = ?
            GROUP BY term, usn
            ORDER BY "Reimbursement Days Allowed" DESC, "Total Session Entries" DESC
        ''', [selected_report_term])
        
    if not df_report.empty:
        st.dataframe(df_report, use_container_width=True)
        
        buffer_rep = io.BytesIO()
        with pd.ExcelWriter(buffer_rep, engine='xlsxwriter') as writer:
            df_report.to_excel(writer, index=False, sheet_name='Reimbursement Report')
        
        term_slug = selected_report_term.replace(" ", "_")
        st.download_button(
            label=f"📥 Download Reimbursement Report ({selected_report_term})",
            data=buffer_rep.getvalue(),
            file_name=f"Reimbursement_Report_{term_slug}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning(f"No reimbursement records available for {selected_report_term}.")
