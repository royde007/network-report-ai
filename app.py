import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Network Report Auditor", layout="wide")
st.title("📡 Sector & Carrier Keyed Comparison Tool")

# --- STYLES & CONSTANTS ---
RED_FILL = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
LIGHT_RED_FILL = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
GREEN_FILL = PatternFill(start_color='CCFFCC', end_color='CCFFCC', fill_type='solid')
YELLOW_FILL = PatternFill(start_color='FFFFCC', end_color='FFFFCC', fill_type='solid')
HEADER_FILL = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
HEADER_FONT = Font(color='FFFFFF', bold=True)
THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), 
                     top=Side(style='thin'), bottom=Side(style='thin'))

# --- SIDEBAR INSTRUCTIONS ---
with st.sidebar:
    st.header("📋 Processing Rules")
    st.info("""
    1. **Skip**: 1st Sheet (General Information) is ignored.
    2. **Skip**: Any sheet name ending in 'Pivot' is ignored.
    3. **Key**: Logic searches for **'Sector Name'** and **'Carrier'** columns anywhere in the sheet to create the unique ID.
    4. **Output**: Formatted Excel reports for each valid sheet.
    """)

# --- FILE UPLOADERS ---
col1, col2 = st.columns(2)
with col1:
    pre_files = st.file_uploader("Upload BASELINE (Pre) Reports", accept_multiple_files=True)
with col2:
    post_files = st.file_uploader("Upload CURRENT (Post) Reports", accept_multiple_files=True)

def find_key_columns(columns):
    """Dynamically finds the 'Sector Name' and 'Carrier' columns."""
    cols_clean = [str(c).strip().lower() for c in columns]
    s_idx = None
    c_idx = None
    
    for i, name in enumerate(cols_clean):
        if name == "sector name": s_idx = i
        if name == "carrier": c_idx = i
        
    if s_idx is not None and c_idx is not None:
        return columns[s_idx], columns[c_idx]
    return None, None

def compare_dataframes_to_excel(df1, df2, sheet_name):
    # Standardize column headers
    df1.columns = df1.columns.astype(str).str.strip()
    df2.columns = df2.columns.astype(str).str.strip()

    # Find the Specific Key Columns
    sec_col, car_col = find_key_columns(df1.columns)
    
    if not sec_col or not car_col:
        return "ERROR: 'Sector Name' or 'Carrier' columns not found in this sheet."

    # Create Composite Key
    df1['Comp_Key'] = df1[sec_col].astype(str) + '|' + df1[car_col].astype(str)
    df2['Comp_Key'] = df2[sec_col].astype(str) + '|' + df2[car_col].astype(str)
    
    all_keys = sorted(set(df1['Comp_Key']).union(set(df2['Comp_Key'])))
    # Columns to compare (everything except the keys)
    other_cols = sorted(set(df1.columns) | set(df2.columns) - {sec_col, car_col, 'Comp_Key'})
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparison Results"
    
    # Write Headers
    headers = [sec_col, car_col, 'Status']
    for col in other_cols:
        headers += [f"{col} (Pre)", f"{col} (Post)", f"{col} Match?"]
    
    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c_idx, value=h)
        cell.fill, cell.font, cell.border = HEADER_FILL, HEADER_FONT, THIN_BORDER

    # Comparison Process
    row_idx = 2
    for key in all_keys:
        r1 = df1[df1['Comp_Key'] == key]
        r2 = df2[df2['Comp_Key'] == key]
        k_parts = key.split('|', 1)
        
        ws.cell(row=row_idx, column=1, value=k_parts[0]).border = THIN_BORDER
        ws.cell(row=row_idx, column=2, value=k_parts[1]).border = THIN_BORDER
        status_cell = ws.cell(row=row_idx, column=3)
        status_cell.border = THIN_BORDER

        if r1.empty:
            status_cell.value, status_cell.fill = "ONLY IN POST", YELLOW_FILL
        elif r2.empty:
            status_cell.value, status_cell.fill = "ONLY IN PRE", YELLOW_FILL
        else:
            status_cell.value = "IN BOTH"
            has_mismatch = False
            col_ptr = 4
            for col in other_cols:
                v1 = r1.iloc[0][col] if col in r1.columns else 'N/A'
                v2 = r2.iloc[0][col] if col in r2.columns else 'N/A'
                v1, v2 = str(v1), str(v2)
                
                c1 = ws.cell(row=row_idx, column=col_ptr, value=v1)
                c2 = ws.cell(row=row_idx, column=col_ptr+1, value=v2)
                cm = ws.cell(row=row_idx, column=col_ptr+2)
                
                if v1 == v2:
                    cm.value, cm.fill = "✓ MATCH", GREEN_FILL
                else:
                    cm.value, cm.fill, has_mismatch = "✗ MISMATCH", LIGHT_RED_FILL, True
                    c1.fill, c2.fill = RED_FILL, RED_FILL
                
                for c in [c1, c2, cm]: c.border = THIN_BORDER
                col_ptr += 3
            status_cell.fill = LIGHT_RED_FILL if has_mismatch else GREEN_FILL
        row_idx += 1

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- EXECUTION ---
if pre_files and post_files:
    if st.button("🚀 Run Global Comparison"):
        pre_dict = {f.name: f for f in pre_files}
        post_dict = {f.name: f for f in post_files}
        
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for fname, fobj in pre_dict.items():
                if fname in post_dict:
                    report_name = fname.split('.')[0]
                    st.write(f"📁 Processing: **{report_name}**")
                    
                    pre_sheets = pd.read_excel(fobj, sheet_name=None, engine='calamine')
                    post_sheets = pd.read_excel(post_dict[fname], sheet_name=None, engine='calamine')
                    
                    for i, (sname, df_pre) in enumerate(pre_sheets.items()):
                        # 1. Skip 1st Sheet
                        if i == 0:
                            st.write(f"   ⏩ Skipped 1st Sheet: {sname}")
                            continue
                        # 2. Skip Pivot Sheets
                        if sname.lower().endswith('pivot'):
                            st.write(f"   ⏩ Skipped Pivot: {sname}")
                            continue
                        
                        if sname in post_sheets:
                            st.write(f"   ✅ Analyzing Sheet: {sname}")
                            result = compare_dataframes_to_excel(df_pre, post_sheets[sname], sname)
                            
                            if isinstance(result, str) and "ERROR" in result:
                                st.warning(f"   ⚠️ {sname}: {result}")
                            else:
                                zf.writestr(f"{report_name}/{sname}_Comparison.xlsx", result)

        st.success("🏁 Comparison Finished!")
        st.download_button("📥 Download ZIP", zip_buffer.getvalue(), "Network_Audit_Results.zip")
