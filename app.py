import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Network Report Auditor", layout="wide")
st.title("📡 Multi-Sheet Network Comparison Agent")

# --- UI Layout ---
api_key = st.sidebar.text_input("OpenAI API Key", type="password")
pre_files = st.file_uploader("Upload BASELINE (Pre) Reports", accept_multiple_files=True)
post_files = st.file_uploader("Upload CURRENT (Post) Reports", accept_multiple_files=True)

def apply_formatting(output_buffer):
    """Applies the Red/Yellow highlighting logic to the generated buffer."""
    output_buffer.seek(0)
    wb = load_workbook(output_buffer)
    ws = wb.active
    red_fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
    
    # Identify match columns based on your existing logic
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

def compare_sheets(df_pre, df_post, is_first_sheet):
    """Logic to compare two dataframes sheet-by-sheet."""
    # SKIP PIVOT TABLES: If mandatory keys are missing, skip the sheet
    if 'Sector Name' not in df_pre.columns or 'Carrier' not in df_pre.columns:
        return None

    # --- SPECIAL MODIFICATION FOR 1st SHEET ---
    if is_first_sheet:
        # Example modification: Only compare specific columns for the first sheet
        # You can add specific column filters or logic here
        st.info("Applying special logic for the first sheet...")
        # df_pre = df_pre[['Sector Name', 'Carrier', 'Specific_Metric']] 
    # ------------------------------------------

    # Standard Comparison Logic
    df1, df2 = df_pre.copy(), df_post.copy()
    df1['Composite_Key'] = df1['Sector Name'].astype(str) + '|' + df1['Carrier'].astype(str)
    df2['Composite_Key'] = df2['Sector Name'].astype(str) + '|' + df2['Carrier'].astype(str)
    
    all_keys = sorted(set(df1['Composite_Key']).union(set(df2['Composite_Key'])))
    all_cols = sorted(set(df1.columns).union(set(df2.columns)) - {'Sector Name', 'Carrier', 'Composite_Key'})
    
    comparison_results = []
    for key in all_keys:
        sector, carrier = key.split('|', 1)
        r1 = df1[df1['Composite_Key'] == key]
        r2 = df2[df2['Composite_Key'] == key]
        
        row = {'Sector Name': sector, 'Carrier': carrier}
        for col in all_cols:
            val1 = r1.iloc[0][col] if not r1.empty and col in r1.columns else 'N/A'
            val2 = r2.iloc[0][col] if not r2.empty and col in r2.columns else 'N/A'
            row[f'{col} (Pre)'] = val1
            row[f'{col} (Post)'] = val2
            row[f'{col} Match'] = '✓' if str(val1) == str(val2) else '✗'
        comparison_results.append(row)

    # Convert to Excel
    temp_buf = io.BytesIO()
    pd.DataFrame(comparison_results).to_excel(temp_buf, index=False)
    return apply_formatting(temp_buf)

if pre_files and post_files:
    if st.button("🚀 Start Multi-Sheet Comparison"):
        pre_dict = {f.name: f for f in pre_files}
        post_dict = {f.name: f for f in post_files}
        
        master_zip = io.BytesIO()
        with zipfile.ZipFile(master_zip, "w") as zf:
            for file_name, file_obj in pre_dict.items():
                if file_name in post_dict:
                    report_folder = file_name.replace(".xlsm", "").replace(".xlsx", "")
                    st.write(f"📁 Processing Report Folder: **{report_folder}**")
                    
                    # Load all sheets
                    pre_sheets = pd.read_excel(file_obj, sheet_name=None)
                    post_sheets = pd.read_excel(post_dict[file_name], sheet_name=None)
                    
                    for i, (sheet_name, df_pre) in enumerate(pre_sheets.items()):
                        if sheet_name in post_sheets:
                            df_post = post_sheets[sheet_name]
                            
                            # Perform comparison
                            is_first = (i == 0)
                            result = compare_sheets(df_pre, df_post, is_first)
                            
                            if result:
                                # Save in folder/sheet format inside ZIP
                                zf.writestr(f"{report_folder}/{sheet_name}_Comparison.xlsx", result)
                                st.write(f"   ✅ Compared Sheet: {sheet_name}")
                            else:
                                st.write(f"   ⏩ Skipped Sheet (Pivot Table/Empty): {sheet_name}")

        st.success("Comparison Complete!")
        st.download_button("📥 Download Final Reports ZIP", master_zip.getvalue(), "Network_Audit_Results.zip")
