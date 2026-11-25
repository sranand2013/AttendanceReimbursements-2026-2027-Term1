import pandas as pd
import libsql_client
from pathlib import Path

# 1. Configure your Turso credentials and local folder
TURSO_URL = "https://eshaaletursocmslb-<your-org>.turso.io"
TURSO_AUTH_TOKEN = "your-turso-auth-token-here"

LOCAL_FOLDER = r"C:\Users\ANAND.ANAND-PC\Desktop\CMS Data\1sem 2026\Reimbursement Related\Attendance Reimbursement Folder 2026 Term 1"
ACADEMIC_TERM = "2026-27 Term1"

COLUMN_ALIASES = {
    'NAME': ['NAME', 'STUDENT NAME', 'FULL NAME', 'STUDENT'],
    'USN': ['USN', 'REGISTER NUMBER', 'REGISTRATION NUMBER', 'REG NO', 'ROLL NO', 'REGISTER NO'],
    'CLASS': ['CLASS', 'SECTION', 'SEM AND SEC', 'COURSE', 'PROGRAM', 'BATCH', 'CLASS AND SECTION'],
    'DATE': ['DATE', 'DATES', 'DAY', 'DATE OF REIMBURSEMENT'],
    'TIME': ['TIME', 'TIMINGS', 'TIMING', 'FULLDAY', 'FULL DAY', 'HOURS', 'DURATION', 'SESSION'],
    'EVENT': ['REASON', 'EVENT', 'EVENT NAME', 'PROGRAMME', 'PURPOSE'],
    'TERM': ['TERM', 'ACADEMIC TERM', 'SEMESTER TERM']
}

def sync_local_folder():
    root_dir = Path(LOCAL_FOLDER)
    if not root_dir.is_dir():
        print(f"Error: Folder not found -> {LOCAL_FOLDER}")
        return

    client = libsql_client.create_client_sync(url=TURSO_URL, auth_token=TURSO_AUTH_TOKEN)
    client.execute('''
        CREATE TABLE IF NOT EXISTS attendance_reimbursement (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT, usn TEXT, class_name TEXT, date TEXT,
            time TEXT, event_name TEXT, term TEXT, file_name TEXT, file_path TEXT,
            UNIQUE(usn, date, time, event_name, term)
        )
    ''')

    excel_files = [
        p for p in root_dir.rglob("*")
        if p.is_file() and p.suffix.lower() in ['.xlsx', '.xls'] and not p.name.startswith("~$")
    ]
    print(f"Found {len(excel_files)} Excel files. Scanning and syncing to Turso...")

    statements = []
    for file_path in excel_files:
        rel_path = str(file_path.relative_to(root_dir))
        try:
            sheets_dict = pd.read_excel(file_path, header=None, sheet_name=None)
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
                    df['TERM'] = ACADEMIC_TERM
                else:
                    df['TERM'] = df['TERM'].fillna(ACADEMIC_TERM)

                if 'NAME' not in df.columns or 'USN' not in df.columns:
                    continue

                df = df.dropna(subset=['NAME', 'USN'])
                df = df[df['NAME'].astype(str).str.strip().str.upper() != 'NAME']
                df = df[~df['USN'].astype(str).str.strip().str.upper().isin(
                    ['USN', 'REGISTRATION NUMBER', 'REGISTER NUMBER', 'REG NO']
                )]

                for _, row in df.iterrows():
                    row_term = str(row.get('TERM', ACADEMIC_TERM)).strip()
                    if not row_term or row_term.upper() in ('NAN', 'NONE'):
                        row_term = ACADEMIC_TERM

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
                                f"{file_path.name} [{sheet_name}]",
                                rel_path
                            ]
                        )
                    )
            print(f"Parsed: {rel_path}")
        except Exception as e:
            print(f"Skipped {file_path.name}: {e}")

    total_synced = 0
    for idx in range(0, len(statements), 100):
        results = client.batch(statements[idx:idx + 100])
        total_synced += sum(r.rows_affected for r in results)

    print(f"\nDone! Synced {total_synced} records to Turso. Refresh your Streamlit Cloud app to view them.")

if __name__ == "__main__":
    sync_local_folder()
