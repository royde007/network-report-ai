import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Network Report Auditor", layout="wide")
st.title("📡 Multi-Sheet Network Comparison Agent")

with st.sidebar:
    st.header("📋 Audit Instructions")
    st.markdown("""
    1. **Upload Files**: Select 'Pre' and 'Post' reports.
    2. **Smart Filtering**: Only sheets ending in **'Pivot'** are skipped.
    3. **General Info Logic**: The first sheet is compared as a Parameter list.
    4. **Data Sheet Logic**: Other sheets are matched row-by-row based on available headers.
    5. **Highlights**: Differences are marked in **Red**.
    """)

col1, col2 = st.columns(2)
with col1:
    pre_files = st.file_uploader("Upload BASELINE (Pre) Reports", accept_multiple_files=True)
with col2:
    post_files = st.file_uploader("Upload CURRENT (Post) Reports", accept_multiple_files=True)

def apply_formatting(output_buffer):
    output_buffer.seek(0)
    wb = load_workbook(output_buffer)
    ws = wb.active
    red_fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
    
    match_cols = [i for i, cell in enumerate(ws[1], 1) if cell.value and 'Match' in str(cell.value)]
    
    for row_idx in range(2, ws.max_row + 1):
        for col_idx in match_cols:
            if ws.cell(row=row_idx, column=col_idx).value == '✗':
                ws.cell(row=row_idx, column=col_idx-2).fill = red_fill
                ws.cell(row=row_idx, column=col_idx-1).fill = red_fill
                ws.cell(row=row_idx, column=col_idx).fill = red_fill
    
    final_output = io.BytesIO()
    wb.save(final_output)
    return final_output.getvalue()

def compare_general_info(df_pre, df_post):
    """Specific logic for the 1st sheet: Parameter-Value comparison."""
    # Assume 1st column is Parameter Name, 2nd is Value
    df1 = df_pre.iloc[:, :2].copy()
    df2 = df_post.iloc[:, :2].copy()
    df1.columns = ['Parameter', 'Value_Pre']
    df2.columns = ['Parameter', 'Value_Post']
    
    merged = pd.merge(df1, df2, on='Parameter', how='outer')
    merged['Match'] = merged.apply(lambda x: '✓' if str(x['Value_Pre']) == str(x['Value_Post']) else '✗', axis=1)
    
    temp_buf = io.BytesIO()
    merged.to_excel(temp_buf, index=False)
    return apply_formatting(temp_buf)

def compare_standard_sheets(df_pre, df_post):
    """Generic row-by-row comparison for data sheets."""
    # Attempt to find a unique key (Sector/Carrier) or use index if not found
    if 'Sector Name' in df_pre.columns and 'Carrier' in df_pre.columns:
        df_pre['Key'] = df_pre['Sector Name'].astype(str) + '|' + df_pre['Carrier'].astype(str)
        df_post['Key'] = df_post['Sector Name'].astype(str) + '|' + df_post['Carrier'].astype(str)
        key_col = 'Key'
    else:
        # Fallback to row index if no specific key is found
        df_pre = df_pre.reset_index().rename(columns={'index': 'Row_ID'})
        df_post = df_post.reset_index().rename(columns={'index': 'Row_ID'})
        key_col = 'Row_ID'

    all_cols = [c for c in df_pre.columns if c not in [key_col, 'Sector Name', 'Carrier']]
    
    results = []
    # Merge on the key to align rows
    merged = pd.merge(df_pre, df_post, on=key_col, how='outer', suffixes=('_pre', '_post'))
    
    for _, row in merged.iterrows():
        res_row = {}
        # Keep identifier columns if they exist
        if 'Sector Name_pre' in row: res_row['Sector Name'] = row['Sector Name_pre']
        if 'Carrier_pre' in row: res_row['Carrier'] = row['Carrier_pre']
        
        for col in all_cols:
            v1 = row.get(f'{col}_pre', 'N/A')
            v2 = row.get(f'{col}_post', 'N/A')
            res_row[f'{col} (Pre)'] = v1
            res_row[f'{col} (Post)'] = v2
            res_row[f'{col} Match'] = '✓' if str(v1) == str(v2) else '✗'
        results.append(res_row)

    temp_buf = io.BytesIO()
    pd.DataFrame(results).to_excel(temp_buf, index=False)
    return apply_formatting(temp_buf)

if pre_files and post_files:
    if st.button("🚀 Run Global Audit"):
        pre_dict = {f.name: f for f in pre_files}
        post_dict = {f.name: f for f in post_files}
        
        master_zip = io.BytesIO()
        with zipfile.ZipFile(master_zip, "w") as zf:
            for fname, fobj in pre_dict.items():
                if fname in post_dict:
                    folder_name = fname.split('.')[0]
                    st.write(f"📁 **Analyzing Folder:** {folder_name}")
                    
                    try:
                        pre_sheets = pd.read_excel(fobj, sheet_name=None, engine='calamine')
                        post_sheets = pd.read_excel(post_dict[fname], sheet_name=None, engine='calamine')
                        
                        for i, (sname, df_pre) in enumerate(pre_sheets.items()):
                            # RULE 1: Skip only if name ends with 'Pivot'
                            if sname.lower().endswith('pivot'):
                                st.write(f"   ⏩ Skipped: {sname} (Pivot sheet)")
                                continue
                                
                            if sname in post_sheets:
                                df_post = post_sheets[sname]
                                
                                # RULE 2: Special Logic for 1st Sheet (General Info)
                                if i == 0:
                                    st.write(f"   ⚙️ Processing: {sname} (General Info Logic)")
                                    result = compare_general_info(df_pre, df_post)
                                else:
                                    st.write(f"   ✅ Comparing: {sname}")
                                    result = compare_standard_sheets(df_pre, df_post)
                                
                                if result:
                                    zf.writestr(f"{folder_name}/{sname}_Audit.xlsx", result)
                                    
                    except Exception as e:
                        st.error(f"Error reading {fname}: {e}")

        st.success("Audit Complete!")
        st.download_button("📥 Download Results (ZIP)", master_zip.getvalue(), "Audit_Results.zip")
