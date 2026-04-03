import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Network Report Auditor", layout="wide")
st.title("📡 Sector & Carrier Keyed Comparison Tool")

# --- INSTRUCTION PANEL (Sidebar) ---
with st.sidebar:
    st.header("📋 Audit Instructions")
    st.markdown("""
    **Follow these steps:**
    1. **Upload Files**: Select your **BASELINE** and **CURRENT** reports.
    2. **Wait for Processing**: Large `.xlsm` files can take 30-60 seconds to scan.
    3. **Automated Skip**: 
        * 1st sheet (Splash screen) is ignored.
        * *General Information* is ignored.
        * Sheets ending in **'Pivot'** are ignored.
    4. **Unique Key**: Rows are matched using **'Sector Name'** and **'Carrier'**.
    """)
    st.divider()
    st.caption("v2.6 | Optimized Engine")

# --- FILE UPLOADERS ---
col1, col2 = st.columns(2)
with col1:
    pre_files = st.file_uploader("Upload BASELINE (Pre) Reports", accept_multiple_files=True)
with col2:
    post_files = st.file_uploader("Upload CURRENT (Post) Reports", accept_multiple_files=True)

def fast_header_load(file_obj, sheet_name):
    """Efficiently finds the header row without reloading the whole file."""
    try:
        # Step 1: Read only the first 30 rows to find where headers are
        preview = pd.read_excel(file_obj, sheet_name=sheet_name, nrows=30, engine='calamine', header=None)
        
        header_idx = None
        for i, row in preview.iterrows():
            row_vals = [str(v).strip().lower() for v in row.values]
            if 'sector name' in row_vals and 'carrier' in row_vals:
                header_idx = i
                break
        
        if header_idx is not None:
            # Step 2: Load the full sheet starting from that row
            df = pd.read_excel(file_obj, sheet_name=sheet_name, skiprows=header_idx, engine='calamine')
            df.columns = df.columns.astype(str).str.strip()
            # Drop completely empty rows/cols to save memory
            df = df.dropna(how='all').dropna(axis=1, how='all')
            return df
    except Exception as e:
        st.error(f"Error in {sheet_name}: {e}")
    return None

def compare_dataframes_to_excel(df1, df2):
    # Standardize columns
    cols_clean1 = {str(c).strip().lower(): c for c in df1.columns}
    cols_clean2 = {str(c).strip().lower(): c for c in df2.columns}
    
    sec_col = cols_clean1['sector name']
    car_col = cols_clean1['carrier']

    df1['Comp_Key'] = df1[sec_col].astype(str) + '|' + df1[car_col].astype(str)
    df2['Comp_Key'] = df2[cols_clean2['sector name']].astype(str) + '|' + df2[cols_clean2['carrier']].astype(str)
    
    all_keys = sorted(set(df1['Comp_Key']).union(set(df2['Comp_Key'])))
    other_cols = sorted(set(df1.columns) | set(df2.columns) - {sec_col, car_col, 'Comp_Key'})
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparison Results"
    
    # Styles
    red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
    light_red_fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
    green_fill = PatternFill(start_color='CCFFCC', end_color='CCFFCC', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFFFCC', end_color='FFFFCC', fill_type='solid')
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    headers = [sec_col, car_col, 'Status']
    for col in other_cols:
        headers += [f"{col} (Pre)", f"{col} (Post)", f"{col} Match?"]
    
    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c_idx, value=h)
        cell.fill, cell.font, cell.border = header_fill, header_font, thin_border

    row_idx = 2
    for key in all_keys:
        r1 = df1[df1['Comp_Key'] == key]
        r2 = df2[df2['Comp_Key'] == key]
        k_parts = key.split('|', 1)
        
        ws.cell(row=row_idx, column=1, value=k_parts[0]).border = thin_border
        ws.cell(row=row_idx, column=2, value=k_parts[1]).border = thin_border
        status_cell = ws.cell(row=row_idx, column=3)
        status_cell.border = thin_border

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
                
                for c in [c1, c2, cm]: c.border = thin_border
                col_ptr += 3
            status_cell.fill = light_red_fill if has_mismatch else green_fill
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
                    xl_pre = pd.ExcelFile(fobj, engine='calamine')
                    xl_post = pd.ExcelFile(post_dict[fname], engine='calamine')
                    
                    for i, sname in enumerate(xl_pre.sheet_names):
                        # Skip Logic
                        if i == 0 or sname.lower().endswith('pivot') or sname == "General Information":
                            st.write(f"⏩ Skipped: {sname}")
                            continue
                        
                        st.write(f"🔍 Analyzing: {sname}...")
                        df_pre = fast_header_load(fobj, sname)
                        df_post = fast_header_load(post_dict[fname], sname)
                        
                        if df_pre is not None and df_post is not None:
                            result = compare_dataframes_to_excel(df_pre, df_post)
                            zf.writestr(f"{report_name}/{sname}_Comparison.xlsx", result)
                            st.write(f"✅ {sname} Processed.")
                        else:
                            st.warning(f"⚠️ {sname}: Could not find data headers.")

        st.success("🏁 Comparison Finished!")
        st.download_button("📥 Download Results", zip_buffer.getvalue(), "Network_Audit_Results.zip")
