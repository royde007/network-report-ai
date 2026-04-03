import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Network Audit Portal", layout="wide")
st.title("📡 Universal Network Audit Comparison Portal")

# --- STYLES ---
RED_FILL = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
LIGHT_RED_FILL = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
GREEN_FILL = PatternFill(start_color='CCFFCC', end_color='CCFFCC', fill_type='solid')
YELLOW_FILL = PatternFill(start_color='FFFFCC', end_color='FFFFCC', fill_type='solid')
HEADER_FILL = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
HEADER_FONT = Font(color='FFFFFF', bold=True)
THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

# --- REPORT CONFIGURATION (The "Switch Statement" Logic) ---
# Add new reports here. The app will look for these keys to do the matching.
REPORT_CONFIG = {
    "Access Distance": {
        "primary_key": "Sector Name",
        "secondary_key": "Carrier",
        "description": "Standard Access Distance Histogram Analysis"
    },
    "Throughput Report": {
        "primary_key": "Cell Name",
        "secondary_key": "Frequency",
        "description": "User Throughput and Cell Load Analysis"
    }
}

# --- SIDEBAR ---
with st.sidebar:
    st.header("⚙️ Configuration")
    report_type = st.selectbox("Select Report Type", options=list(REPORT_CONFIG.keys()))
    st.info(f"**Mode:** {REPORT_CONFIG[report_type]['description']}")
    st.divider()
    st.markdown("""
    **Global Rules:**
    - Skips 1st sheet & 'General Information'.
    - Skips any sheet ending in 'Pivot'.
    - Preserves PRE column order.
    """)

# --- INTERFACE ---
col1, col2 = st.columns(2)
with col1:
    pre_files = st.file_uploader("Upload PRE Reports", accept_multiple_files=True)
with col2:
    post_files = st.file_uploader("Upload POST Reports", accept_multiple_files=True)

def get_data_with_headers(file_obj, sheet_name, p_key, s_key):
    """Streams file and finds the header row based on config keys."""
    try:
        file_obj.seek(0)
        wb = load_workbook(file_obj, read_only=True, data_only=True)
        ws = wb[sheet_name]
        data = []
        header_row_idx = None
        
        # Scan first 50 rows for headers
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            if i > 50: break
            row_vals = [str(v).strip().lower() if v else "" for v in row]
            if p_key.lower() in row_vals and s_key.lower() in row_vals:
                header_row_idx = i
                headers = [str(v).strip() if v else f"Col_{j}" for j, v in enumerate(row)]
                break
        
        if header_row_idx:
            for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
                if any(v is not None for v in row): data.append(row)
            return pd.DataFrame(data, columns=headers)
    except: return None
    return None

def run_comparison_logic(df1, df2, p_key, s_key):
    """Core comparison engine using dynamic keys from config."""
    # Preserve original order from PRE
    original_order = [c for c in df1.columns if c.lower() not in [p_key.lower(), s_key.lower(), 'k']]
    
    # Create matching index
    df1['K'] = df1[p_key].astype(str).str.strip() + '|' + df1[s_key].astype(str).str.strip()
    df2['K'] = df2[s_key].astype(str).str.strip() # Temporary placeholder logic
    # Find actual POST column names (case insensitive)
    p_key_post = [c for c in df2.columns if c.lower() == p_key.lower()][0]
    s_key_post = [c for c in df2.columns if c.lower() == s_key.lower()][0]
    df2['K'] = df2[p_key_post].astype(str).str.strip() + '|' + df2[s_key_post].astype(str).str.strip()

    df1_idx = df1.drop_duplicates(subset='K').set_index('K')
    df2_idx = df2.drop_duplicates(subset='K').set_index('K')
    
    all_keys = sorted(set(df1_idx.index).union(set(df2_idx.index)))
    
    wb = Workbook()
    ws_det = wb.active
    ws_det.title = "Comparison Results"
    
    # Headers
    headers = [p_key, s_key, 'Status']
    for col in original_order:
        headers += [f"{col}\n(PRE)", f"{col}\n(POST)", f"{col}\nMatch?"]
    
    for c_idx, h in enumerate(headers, 1):
        cell = ws_det.cell(row=1, column=c_idx, value=h)
        cell.fill, cell.font, cell.border, cell.alignment = HEADER_FILL, HEADER_FONT, THIN_BORDER, Alignment(wrap_text=True)

    stats = {"match": 0, "mismatch": 0, "only_pre": 0, "only_post": 0}

    for row_idx, key in enumerate(all_keys, 2):
        k_parts = key.split('|', 1)
        ws_det.cell(row=row_idx, column=1, value=k_parts[0]).border = THIN_BORDER
        ws_det.cell(row=row_idx, column=2, value=k_parts[1]).border = THIN_BORDER
        status_cell = ws_det.cell(row=row_idx, column=3)
        status_cell.border = THIN_BORDER

        if key not in df1_idx.index:
            status_cell.value, status_cell.fill, stats["only_post"] = "ONLY IN POST", YELLOW_FILL, stats["only_post"] + 1
        elif key not in df2_idx.index:
            status_cell.value, status_cell.fill, stats["only_pre"] = "ONLY IN PRE", YELLOW_FILL, stats["only_pre"] + 1
        else:
            status_cell.value = "IN BOTH"
            r1, r2 = df1_idx.loc[key], df2_idx.loc[key]
            has_mismatch, col_ptr = False, 4
            for col in original_order:
                v1, v2 = str(r1.get(col, 'NULL')), str(r2.get(col, 'NULL'))
                c1, c2, cm = ws_det.cell(row=row_idx, column=col_ptr, value=v1), ws_det.cell(row=row_idx, column=col_ptr+1, value=v2), ws_det.cell(row=row_idx, column=col_ptr+2)
                if v1 == v2: cm.value, cm.fill = "✓ MATCH", GREEN_FILL
                else:
                    cm.value, cm.fill, has_mismatch = "✗ MISMATCH", LIGHT_RED_FILL, True
                    c1.fill, c2.fill = RED_FILL, RED_FILL
                for c in [c1, c2, cm]: c.border = THIN_BORDER
                col_ptr += 3
            status_cell.fill, stats["mismatch" if has_mismatch else "match"] = (LIGHT_RED_FILL if has_mismatch else GREEN_FILL), (stats["mismatch" if has_mismatch else "match"] + 1)

    # Summary Sheet
    ws_sum = wb.create_sheet("Summary", 0)
    # ... [Summary and Legend logic as developed previously] ...
    # (Simplified for space, includes matching records, legend, and borders)
    ws_sum.cell(row=1, column=1, value="AUDIT SUMMARY").font = Font(bold=True)
    ws_sum.cell(row=3, column=1, value=f"Report Type: {report_type}")
    ws_sum.cell(row=4, column=1, value=f"Total Keys: {len(all_keys)}")
    ws_sum.cell(row=5, column=1, value=f"Matches: {stats['match']}")
    ws_sum.cell(row=6, column=1, value=f"Mismatches: {stats['mismatch']}")
    
    # Add Legend...
    legend_data = [("Match", GREEN_FILL), ("Mismatch", LIGHT_RED_FILL), ("Cell Error", RED_FILL), ("Unique", YELLOW_FILL)]
    for i, (text, fill) in enumerate(legend_data, 8):
        ws_sum.cell(row=i, column=1, value=text).fill = fill

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- RUN EXECUTION ---
if st.button("🚀 Run Global Audit"):
    if pre_files and post_files:
        config = REPORT_CONFIG[report_type]
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            pre_dict = {f.name: f for f in pre_files}
            post_dict = {f.name: f for f in post_files}
            for fname, fobj in pre_dict.items():
                if fname in post_dict:
                    xl = pd.ExcelFile(fobj, engine='openpyxl')
                    for i, sname in enumerate(xl.sheet_names):
                        if i == 0 or sname.lower().endswith('pivot') or sname == "General Information": continue
                        df_pre = get_data_with_headers(fobj, sname, config['primary_key'], config['secondary_key'])
                        df_post = get_data_with_headers(post_dict[fname], sname, config['primary_key'], config['secondary_key'])
                        if df_pre is not None and df_post is not None:
                            report = run_comparison_logic(df_pre, df_post, config['primary_key'], config['secondary_key'])
                            zf.writestr(f"{fname.split('.')[0]}/{sname}_Audit.xlsx", report)
        st.success("Audit Complete!")
        st.download_button("📥 Download Results", zip_buffer.getvalue(), "Audit_Results.zip")
