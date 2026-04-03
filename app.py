import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Network Report Auditor", layout="wide")
st.title("📡 Sector & Carrier Keyed Comparison Tool")

# --- STYLES ---
RED_FILL = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
LIGHT_RED_FILL = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
GREEN_FILL = PatternFill(start_color='CCFFCC', end_color='CCFFCC', fill_type='solid')
YELLOW_FILL = PatternFill(start_color='FFFFCC', end_color='FFFFCC', fill_type='solid')
HEADER_FILL = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
HEADER_FONT = Font(color='FFFFFF', bold=True)
BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

# --- INSTRUCTION PANEL (Sidebar) ---
with st.sidebar:
    st.header("📋 Audit Instructions")
    st.markdown("""
    1. **Upload Files**: Select **PRE** and **POST** reports.
    2. **Wait for Upload**: Large files (20MB+) take a moment to transfer.
    3. **Automated Skip**: 
        * 1st sheet and *General Information* are ignored.
        * Sheets ending in **'Pivot'** are ignored.
    4. **Output**: Two tabs (Summary + Detailed) with side-by-side comparison.
    """)
    st.divider()
    st.caption("v3.1 | High-Speed Indexing | PRE vs POST")

# --- HELPER FUNCTIONS ---

def streaming_load(file_obj, sheet_name):
    """Fast streaming load that only grabs data rows."""
    try:
        file_obj.seek(0)
        # Using read_only for memory efficiency
        wb = load_workbook(file_obj, read_only=True, data_only=True)
        if sheet_name not in wb.sheetnames:
            return None
        ws = wb[sheet_name]
        data = []
        header_row_idx = None
        # Scan for headers
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            if i > 50: break 
            row_vals = [str(v).strip().lower() if v is not None else "" for v in row]
            if 'sector name' in row_vals and 'carrier' in row_vals:
                header_row_idx = i
                headers = [str(v).strip() if v is not None else f"Col_{j}" for j, v in enumerate(row)]
                break
        if header_row_idx:
            for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
                if any(v is not None for v in row):
                    data.append(row)
            return pd.DataFrame(data, columns=headers)
    except Exception as e:
        st.error(f"Error reading {sheet_name}: {e}")
    return None

def create_comparison_report(df1, df2):
    """High-speed comparison using Pandas Indexing."""
    # Find key columns
    cols_pre_map = {str(c).strip().lower(): c for c in df1.columns}
    cols_post_map = {str(c).strip().lower(): c for c in df2.columns}
    sec_col = cols_pre_map['sector name']
    car_col = cols_pre_map['carrier']

    # Create keys and set index for O(1) lookup speed
    df1['K'] = df1[sec_col].astype(str).str.strip() + '|' + df1[car_col].astype(str).str.strip()
    df2['K'] = df2[cols_post_map['sector name']].astype(str).str.strip() + '|' + df2[cols_post_map['carrier']].astype(str).str.strip()
    
    # Drop rows with duplicate keys to prevent errors
    df1 = df1.drop_duplicates(subset='K').set_index('K')
    df2 = df2.drop_duplicates(subset='K').set_index('K')
    
    all_keys = sorted(set(df1.index).union(set(df2.index)))
    other_cols = sorted(set(df1.columns) | set(df2.columns) - {sec_col, car_col})
    
    wb = Workbook()
    ws_det = wb.active
    ws_det.title = "Comparison Results"
    
    # Headers
    headers = [sec_col, car_col, 'Status']
    for col in other_cols:
        headers += [f"{col}\n(PRE)", f"{col}\n(POST)", f"{col}\nMatch?"]
    
    for c_idx, h in enumerate(headers, 1):
        cell = ws_det.cell(row=1, column=c_idx, value=h)
        cell.fill, cell.font, cell.border = HEADER_FILL, HEADER_FONT, BORDER
        ws_det.column_dimensions[cell.column_letter].width = 18

    stats = {"match": 0, "mismatch": 0, "only_pre": 0, "only_post": 0}
    
    # Comparison Loop (Now much faster with indexing)
    for row_idx, key in enumerate(all_keys, 2):
        in_pre = key in df1.index
        in_post = key in df2.index
        k_parts = key.split('|', 1)
        
        ws_det.cell(row=row_idx, column=1, value=k_parts[0]).border = BORDER
        ws_det.cell(row=row_idx, column=2, value=k_parts[1]).border = BORDER
        status_cell = ws_det.cell(row=row_idx, column=3)
        status_cell.border = BORDER

        if not in_pre:
            status_cell.value, status_cell.fill, stats["only_post"] = "ONLY IN POST", YELLOW_FILL, stats["only_post"] + 1
        elif not in_post:
            status_cell.value, status_cell.fill, stats["only_pre"] = "ONLY IN PRE", YELLOW_FILL, stats["only_pre"] + 1
        else:
            status_cell.value = "IN BOTH FILES"
            row1, row2 = df1.loc[key], df2.loc[key]
            has_mismatch, col_ptr = False, 4
            for col in other_cols:
                v1 = str(row1.get(col, 'NULL'))
                v2 = str(row2.get(col, 'NULL'))
                c1 = ws_det.cell(row=row_idx, column=col_ptr, value=v1)
                c2 = ws_det.cell(row=row_idx, column=col_ptr+1, value=v2)
                cm = ws_det.cell(row=row_idx, column=col_ptr+2)
                
                if v1 == v2:
                    cm.value, cm.fill = "✓ MATCH", GREEN_FILL
                else:
                    cm.value, cm.fill, has_mismatch = "✗ MISMATCH", LIGHT_RED_FILL, True
                    c1.fill, c2.fill = RED_FILL, RED_FILL
                
                for c in [c1, c2, cm]: c.border = BORDER
                col_ptr += 3
            
            if has_mismatch:
                status_cell.fill, stats["mismatch"] = LIGHT_RED_FILL, stats["mismatch"] + 1
            else:
                status_cell.fill, stats["match"] = GREEN_FILL, stats["match"] + 1

    # Summary Tab
    ws_sum = wb.create_sheet("Summary", 0)
    summary_rows = [
        ["COMPARISON SUMMARY", ""], [""], ["File Information:", ""],
        ["Total Records in PRE", len(df1)], ["Total Records in POST", len(df2)],
        ["Total Unique Keys", len(all_keys)], ["", ""], ["Comparison Results:", ""],
        ["✓ Matching Records", stats["match"]], ["✗ Mismatching Records", stats["mismatch"]],
        ["📄 Records Only in PRE", stats["only_pre"]], ["📄 Records Only in POST", stats["only_post"]],
        ["", ""], ["Generated on:", pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")]
    ]
    for r_idx, row_data in enumerate(summary_rows, 1):
        for c_idx, val in enumerate(row_data, 1):
            cell = ws_sum.cell(row=r_idx, column=c_idx, value=val)
            if r_idx == 1: cell.font, cell.fill = Font(bold=True, size=14, color="FFFFFF"), HEADER_FILL

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- MAIN APP ---

col1, col2 = st.columns(2)
with col1:
    pre_files = st.file_uploader("Upload PRE Reports", accept_multiple_files=True)
with col2:
    post_files = st.file_uploader("Upload POST Reports", accept_multiple_files=True)

if st.button("🚀 Run Global Audit"):
    if pre_files and post_files:
        pre_dict = {f.name: f for f in pre_files}
        post_dict = {f.name: f for f in post_files}
        zip_buffer = io.BytesIO()
        processed_any = False

        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for fname, fobj in pre_dict.items():
                if fname in post_dict:
                    st.info(f"📁 Processing Report: {fname}")
                    # Fast sheet listing using Calamine
                    xl = pd.ExcelFile(fobj, engine='calamine')
                    for i, sname in enumerate(xl.sheet_names):
                        if i == 0 or sname.lower().endswith('pivot') or sname == "General Information":
                            st.write(f"   ⏩ Skipping: {sname}")
                            continue
                        
                        df_pre = streaming_load(fobj, sname)
                        df_post = streaming_load(post_dict[fname], sname)
                        
                        if df_pre is not None and df_post is not None:
                            st.write(f"   ⚙️ Analyzing Sheet: {sname}...")
                            report_bytes = create_comparison_report(df_pre, df_post)
                            zf.writestr(f"{fname.split('.')[0]}/{sname}_Audit.xlsx", report_bytes)
                            processed_any = True
        
        if processed_any:
            st.success("🏁 Audit Complete!")
            st.download_button("📥 Download ZIP", zip_buffer.getvalue(), "Network_Audit_Results.zip")
        else:
            st.error("No valid
