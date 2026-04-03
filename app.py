import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Network Audit Tool", layout="wide")
st.title("📡 Sector & Carrier Comparison Agent")

# --- INSTRUCTION PANEL (Sidebar) ---
with st.sidebar:
    st.header("📋 Audit Instructions")
    st.info("""
    - **Skips**: 'General Information' and any sheet ending in 'Pivot'.
    - **Matching**: Uses 'Sector Name' + 'Carrier' as a unique key.
    - **Output**: Generates a ZIP with folders containing Summary and Detailed comparisons per sheet.
    """)
    st.divider()
    st.caption("v2.3 | Engine: Calamine")

# --- FILE UPLOADERS ---
col1, col2 = st.columns(2)
with col1:
    pre_files = st.file_uploader("Upload BASELINE (Pre) Reports", accept_multiple_files=True)
with col2:
    post_files = st.file_uploader("Upload CURRENT (Post) Reports", accept_multiple_files=True)

def get_record_values(record, exclude_cols):
    if record.empty: return {}
    return {col: record.iloc[0][col] if pd.notna(record.iloc[0][col]) else 'NULL' 
            for col in record.columns if col not in exclude_cols}

def apply_formatting_detailed(wb_sheet):
    """Applies Red highlighting to mismatched cells in the 'Detailed Comparison' sheet."""
    red_fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
    # Find columns ending with 'Match'
    match_columns = [col_idx for col_idx, cell in enumerate(wb_sheet[1], 1) if cell.value and 'Match' in str(cell.value)]
    
    for col_idx in match_columns:
        for row_idx in range(2, wb_sheet.max_row + 1):
            cell = wb_sheet.cell(row=row_idx, column=col_idx)
            if cell.value == '✗':
                # Highlight BASELINE (Pre) [col-2], CURRENT (Post) [col-1], and Match [col]
                wb_sheet.cell(row=row_idx, column=col_idx-2).fill = red_fill
                wb_sheet.cell(row=row_idx, column=col_idx-1).fill = red_fill
                cell.fill = red_fill

def compare_logic_merged(df_pre, df_post):
    """Merged Logic: Composite Key Alignment + Side-by-Side Comparison."""
    # Create Composite Key
    df_pre['Composite_Key'] = df_pre['Sector Name'].astype(str) + '|' + df_pre['Carrier'].astype(str)
    df_post['Composite_Key'] = df_post['Sector Name'].astype(str) + '|' + df_post['Carrier'].astype(str)
    
    all_keys = sorted(set(df_pre['Composite_Key']).union(set(df_post['Composite_Key'])))
    exclude = ['Sector Name', 'Carrier', 'Composite_Key']
    all_cols = sorted(set(df_pre.columns).union(set(df_post.columns)) - set(exclude))
    
    summary_data = []
    detailed_data = []

    for key in all_keys:
        sector_name, carrier = key.split('|', 1)
        record_pre = df_pre[df_pre['Composite_Key'] == key]
        record_post = df_post[df_post['Composite_Key'] == key]
        
        row_detailed = {'Sector Name': sector_name, 'Carrier': carrier}
        
        if record_pre.empty:
            status, match_s = 'Only in CURRENT (Post)', 'Missing in BASELINE (Pre)'
            vals_pre, vals_post = {}, get_record_values(record_post, exclude)
            mismatches = ["Record missing in Baseline"]
        elif record_post.empty:
            status, match_s = 'Only in BASELINE (Pre)', 'Missing in CURRENT (Post)'
            vals_pre, vals_post = get_record_values(record_pre, exclude), {}
            mismatches = ["Record missing in Current"]
        else:
            status = 'In Both Files'
            vals_pre = get_record_values(record_pre, exclude)
            vals_post = get_record_values(record_post, exclude)
            mismatches = []
            
            for col in all_cols:
                v_pre = vals_pre.get(col, 'N/A')
                v_post = vals_post.get(col, 'N/A')
                row_detailed[f'{col} (BASELINE (Pre))'] = v_pre
                row_detailed[f'{col} (CURRENT (Post))'] = v_post
                match_icon = '✓' if str(v_pre) == str(v_post) else '✗'
                row_detailed[f'{col} Match'] = match_icon
                if match_icon == '✗': 
                    mismatches.append(f"{col}: {v_pre} vs {v_post}")
            
            match_s = 'MATCH' if not mismatches else 'MISMATCH'

        summary_row = {
            'Sector Name': sector_name, 
            'Carrier': carrier, 
            'Status': status,
            'Match_Status': match_s, 
            'Mismatch_Details': '; '.join(mismatches) if mismatches else ''
        }
        summary_data.append(summary_row)
        if status == 'In Both Files':
            detailed_data.append(row_detailed)

    # Create Excel Output
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)
        pd.DataFrame(detailed_data).to_excel(writer, sheet_name='Detailed Comparison', index=False)
    
    output.seek(0)
    wb = load_workbook(output)
    if 'Detailed Comparison' in wb.sheetnames:
        apply_formatting_detailed(wb['Detailed Comparison'])
    
    final_buf = io.BytesIO()
    wb.save(final_buf)
    return final_buf.getvalue()

# --- EXECUTION ---
if pre_files and post_files:
    if st.button("🚀 Run Global Comparison"):
        pre_dict = {f.name: f for f in pre_files}
        post_dict = {f.name: f for f in post_files}
        
        master_zip = io.BytesIO()
        with zipfile.ZipFile(master_zip, "w") as zf:
            for name, file in pre_dict.items():
                if name in post_dict:
                    folder_name = name.split('.')[0]
                    st.write(f"📁 **Analyzing Folder:** {folder_name}")
                    
                    try:
                        pre_sheets = pd.read_excel(file, sheet_name=None, engine='calamine')
                        post_sheets = pd.read_excel(post_dict[name], sheet_name=None, engine='calamine')
                        
                        for sname, df_pre in pre_sheets.items():
                            # SKIP LOGIC: General Information and sheets ending in Pivot
                            if sname == "General Information" or sname.lower().endswith('pivot'):
                                st.write(f"   ⏩ Skipped: {sname}")
                                continue
                                
                            if sname in post_sheets:
                                st.write(f"   ✅ Comparing: {sname}")
                                res = compare_logic_merged(df_pre, post_sheets[sname])
                                zf.writestr(f"{folder_name}/{sname}_Comparison.xlsx", res)
                                
                    except Exception as e:
                        st.error(f"Error reading {name}: {e}")

        st.success("Comparison Complete!")
        st.download_button(
            label="📥 Download Audit Results (ZIP)",
            data=master_zip.getvalue(),
            file_name="Network_Audit_Results.zip",
            mime="application/zip"
        )
