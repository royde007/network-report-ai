import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Network Report Auditor", layout="wide")
st.title("📡 Sector & Carrier Keyed Comparison Tool")

# --- INSTRUCTION PANEL ---
with st.sidebar:
    st.header("📋 Audit Instructions")
    st.markdown("""
    **Optimization Active:**
    * **Streaming Mode**: Large `.xlsm` files are now streamed to prevent freezing.
    * **Auto-Header**: Scans for 'Sector Name' and 'Carrier'.
    * **Skips**: 1st sheet, 'General Information', and sheets ending in 'Pivot'.
    """)
    st.divider()
    st.caption("v2.7 | Streaming Reader (High Speed)")

col1, col2 = st.columns(2)
with col1:
    pre_files = st.file_uploader("Upload BASELINE (Pre) Reports", accept_multiple_files=True)
with col2:
    post_files = st.file_uploader("Upload CURRENT (Post) Reports", accept_multiple_files=True)

def streaming_load(file_obj, sheet_name):
    """Uses read_only mode to handle large files without getting stuck."""
    try:
        # Reset file pointer
        file_obj.seek(0)
        # Load workbook in read_only mode (extremely fast for large files)
        wb = load_workbook(file_obj, read_only=True, data_only=True)
        ws = wb[sheet_name]
        
        data = []
        header_row_idx = None
        
        # Scan first 50 rows to find headers
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            if i > 50: break 
            row_vals = [str(v).strip().lower() if v is not None else "" for v in row]
            if 'sector name' in row_vals and 'carrier' in row_vals:
                header_row_idx = i
                headers = [str(v).strip() if v is not None else f"Column_{j}" for j, v in enumerate(row)]
                break
        
        if header_row_idx:
            # Extract data starting from row after headers
            for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
                # Only add rows that aren't completely empty
                if any(v is not None for v in row):
                    data.append(row)
            
            df = pd.DataFrame(data, columns=headers)
            # Cleanup
            df = df.dropna(how='all').dropna(axis=1, how='all')
            return df
    except Exception as e:
        st.error(f"Error in {sheet_name}: {e}")
    return None

def compare_dataframes_to_excel(df1, df2):
    # Map headers case-insensitively
    cols_clean1 = {str(c).strip().lower(): c for c in df1.columns}
    cols_clean2 = {str(c).strip().lower(): c for c in df2.columns}
    
    sec_col = cols_clean1['sector name']
    car_col = cols_clean1['carrier']

    df1['Comp_Key'] = df1[sec_col].astype(str).str.strip() + '|' + df1[car_col].astype(str).str.strip()
    df2['Comp_Key'] = df2[cols_clean2['sector name']].astype(str).str.strip() + '|' + df2[cols_clean2['carrier']].astype(str).str.strip()
    
    all_keys = sorted(set(df1['Comp_Key']).union(set(df2['Comp_Key'])))
    other_cols = sorted(set(df1.columns) | set(df2.columns) - {sec_col, car_col, 'Comp_Key'})
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparison Results"
    
    # Formatters
    red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
    light_red_fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
    green_fill = PatternFill(start_color='CCFFCC', end_color='CCFFCC', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFFFCC', end_color='FFFFCC', fill_type='solid')
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    headers = [sec_col, car_col, 'Status']
    for col in other_cols:
        headers += [f"{col} (Pre)", f"{col} (Post)", f"{col} Match?"]
    
    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c_idx, value=h)
        cell.fill, cell.font, cell.border = header_fill, header_font, border

    row_idx = 2
    for key in all_keys:
        r1 = df1[df1['Comp_Key'] == key]
        r2 = df2[df2['Comp_Key'] == key]
        k_parts = key.split('|', 1)
        
        ws.cell(row=row_idx, column=1, value=k_parts[0]).border = border
        ws.cell(row=row_idx, column=2, value=k_parts[1]).border = border
        status_cell = ws.cell(row=row_idx, column=3)
        status_cell.border = border

        if r1.empty:
            status_cell.value, status_cell.fill = "ONLY IN POST", yellow_fill
        elif r2.empty:
            status_cell.value, status_cell.fill = "ONLY IN PRE", yellow_fill
        else:
            status_cell.value = "IN BOTH"
            has_mismatch = False
            col_ptr = 4
            for col in other_cols:
                v1, v2 = str(r1.iloc[0].get(col, 'N/A')), str(r2.iloc[0].get(col, 'N/A'))
                c1, c2 = ws.cell(row=row_idx, column=col_ptr, value=v1), ws.cell(row=row_idx, column=col_ptr+1, value=v2)
                cm = ws.cell(row=row_idx, column=col_ptr+2)
                
                if v1 == v2:
                    cm.value, cm.fill = "✓ MATCH", green_fill
                else:
                    cm.value, cm.fill, has_mismatch = "✗ MISMATCH", light_red_fill, True
                    c1.fill, c2.fill = red_fill, red_fill
                
                for c in [c1, c2, cm]: c.border = border
                col_ptr += 3
            status_cell.fill = light_red_fill if has_mismatch else green_fill
        row_idx += 1

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- EXECUTION ---
if pre_files and post_files:
    if st.button("🚀 Run Comparison"):
        pre_dict = {f.name: f for f in pre_files}
        post_dict = {f.name: f for f in post_files}
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for fname, fobj in pre_dict.items():
                if fname in post_dict:
                    report_name = fname.split('.')[0]
                    # Get sheet names without loading full file
                    xl = pd.ExcelFile(fobj, engine='openpyxl')
                    
                    for i, sname in enumerate(xl.sheet_names):
                        if i == 0 or sname.lower().endswith('pivot') or sname == "General Information":
                            st.write(f"⏩ Skipped: {sname}")
                            continue
                        
                        st.write(f"🔍 Streaming: {sname}...")
                        df_pre = streaming_load(fobj, sname)
                        df_post = streaming_load(post_dict[fname], sname)
                        
                        if df_pre is not None and df_post is not None:
                            result = compare_dataframes_to_excel(df_pre, df_post)
                            zf.writestr(f"{report_name}/{sname}_Comparison.xlsx", result)
                            st.write(f"✅ {sname} Processed.")
                        else:
                            st.warning(f"⚠️ {sname}: Could not locate Sector/Carrier headers.")

        st.success("🏁 Done!")
        st.download_button("📥 Download ZIP", zip_buffer.getvalue(), "Comparison_Results.zip")
