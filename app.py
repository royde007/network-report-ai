import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Network Report Auditor", layout="wide")
st.title("📡 Automatic Report Comparison Tool")

# --- STYLES ---
RED_FILL = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
LIGHT_RED_FILL = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
GREEN_FILL = PatternFill(start_color='CCFFCC', end_color='CCFFCC', fill_type='solid')
YELLOW_FILL = PatternFill(start_color='FFFFCC', end_color='FFFFCC', fill_type='solid')
HEADER_FILL = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
HEADER_FONT = Font(color='FFFFFF', bold=True)
BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

# --- INSTRUCTION PANEL ---
with st.sidebar:
    st.header("📋 Audit Instructions")
    st.markdown("""
    1. **Upload Files**: Select **PRE** and **POST** reports.
    2. **Logic**:
        * **Summary Tab**: Statistics and legend.
        * **Detailed Tab**: Side-by-side comparison with RED highlights.
    3. **Key**: Matched via 'Sector Name' + 'Carrier'.
    """)
    st.divider()
    st.caption("v2.9 | PRE vs POST Logic")

col1, col2 = st.columns(2)
with col1:
    pre_files = st.file_uploader("Upload PRE Reports", accept_multiple_files=True)
with col2:
    post_files = st.file_uploader("Upload POST Reports", accept_multiple_files=True)

def streaming_load(file_obj, sheet_name):
    """Streams large files to find 'Sector Name' and 'Carrier' headers."""
    try:
        file_obj.seek(0)
        wb = load_workbook(file_obj, read_only=True, data_only=True)
        ws = wb[sheet_name]
        data = []
        header_row_idx = None
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            if i > 50: break 
            row_vals = [str(v).strip().lower() if v is not None else "" for v in row]
            if 'sector name' in row_vals and 'carrier' in row_vals:
                header_row_idx = i
                headers = [str(v).strip() if v is not None else f"Col_{j}" for j, v in enumerate(row)]
                break
        if header_row_idx:
            for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
                if any(v is not None for v in row): data.append(row)
            return pd.DataFrame(data, columns=headers)
    except Exception as e:
        st.error(f"Error reading {sheet_name}: {e}")
    return None

def create_comparison_report(df1, df2, sname):
    # Standardize Key Columns
    cols_pre = {str(c).strip().lower(): c for c in df1.columns}
    cols_post = {str(c).strip().lower(): c for c in df2.columns}
    sec, car = cols_pre['sector name'], cols_pre['carrier']

    df1['K'] = df1[sec].astype(str).str.strip() + '|' + df1[car].astype(str).str.strip()
    df2['K'] = df2[cols_post['sector name']].astype(str).str.strip() + '|' + df2[cols_post['carrier']].astype(str).str.strip()
    
    all_keys = sorted(set(df1['K']).union(set(df2['K'])))
    other_cols = sorted(set(df1.columns) | set(df2.columns) - {sec, car, 'K'})
    
    wb = Workbook()
    
    # --- DETAILED COMPARISON SHEET ---
    ws_det = wb.active
    ws_det.title = "Comparison Results"
    
    headers = [sec, car, 'Status']
    for col in other_cols:
        headers += [f"{col}\n(PRE)", f"{col}\n(POST)", f"{col}\nMatch?"]
    
    for c_idx, h in enumerate(headers, 1):
        cell = ws_det.cell(row=1, column=c_idx, value=h)
        cell.fill, cell.font, cell.border = HEADER_FILL, HEADER_FONT, BORDER
        ws_det.column_dimensions[cell.column_letter].width = 20

    # Statistics for Summary
    stats = {"match": 0, "mismatch": 0, "only_pre": 0, "only_post": 0}
    
    row_idx = 2
    for key in all_keys:
        r1, r2 = df1[df1['K'] == key], df2[df2['K'] == key]
        k_parts = key.split('|', 1)
        ws_det.cell(row=row_idx, column=1, value=k_parts[0]).border = BORDER
        ws_det.cell(row=row_idx, column=2, value=k_parts[1]).border = BORDER
        status_cell = ws_det.cell(row=row_idx, column=3)
        status_cell.border = BORDER

        if r1.empty:
            status_cell.value, status_cell.fill, stats["only_post"] = "ONLY IN POST", YELLOW_FILL, stats["only_post"] + 1
        elif r2.empty:
            status_cell.value, status_cell.fill, stats["only_pre"] = "ONLY IN PRE", YELLOW_FILL, stats["only_pre"] + 1
        else:
            status_cell.value = "IN BOTH FILES"
            has_mismatch, col_ptr = False, 4
            for col in other_cols:
                v1, v2 = str(r1.iloc[0].get(col, 'NULL')), str(r2.iloc[0].get(col, 'NULL'))
                c1, c2 = ws_det.cell(row=row_idx, column=col_ptr, value=v1), ws_det.cell(row=row_idx, column=col_ptr+1, value=v2)
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
        row_idx += 1

    # --- SUMMARY SHEET ---
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

if pre_files and post_files:
    if st.button("🚀 Run Global Audit"):
        pre_dict = {f.name: f for f in pre_files}
        post_dict = {f.name: f for f in post_files}
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for fname, fobj in pre_dict.items():
                if fname in post_dict:
                    xl = pd.ExcelFile(fobj, engine='openpyxl')
                    for i, sname in enumerate(xl.sheet_names):
                        if i == 0 or sname.lower().endswith('pivot') or sname == "General Information":
                            continue
                        df_pre = streaming_load(fobj, sname)
                        df_post = streaming_load(post_dict[fname], sname)
                        if df_pre is not None and df_post is not None:
                            report = create_comparison_report(df_pre, df_post, sname)
                            zf.writestr(f"{fname.split('.')[0]}/{sname}_Audit.xlsx", report)
                            st.write(f"✅ Processed: {sname}")
        st.success("🏁 Audit Complete!")
        st.download_button("📥 Download Results", zip_buffer.getvalue(), "Network_Audit_Results.zip")
