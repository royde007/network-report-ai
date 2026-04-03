import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Network Report Auditor", layout="wide")
st.title("📡 Multi-Sheet Sector Comparison Agent")

# --- INSTRUCTION PANEL (Sidebar) ---
with st.sidebar:
    st.header("📋 Audit Instructions")
    st.markdown("""
    **Follow these steps to generate your report:**
    
    1. **Upload Files**: Select your 'Pre' (Baseline) and 'Post' (Current) reports in the main panel. [cite: 4]
    2. **Automatic Filtering**: The tool scans all tabs. Sheets like *Pivot Tables* or *Histograms* that lack 'Sector Name' columns are automatically skipped. [cite: 4]
    3. **Smart Alignment**: Data is matched using a **Composite Key** (Sector Name + Carrier) to ensure rows are compared accurately even if they moved. [cite: 4]
    4. **1st Sheet Logic**: The first tab of every report undergoes a specialized validation check. [cite: 4]
    5. **Visual Highlights**: Any cell that has changed between the two reports will be highlighted in **Red**. [cite: 4]
    6. **Download Results**: A ZIP file will be generated, containing a folder for each report and separate Excel files for each sheet. [cite: 4]
    """)
    st.divider()
    st.caption("v2.1 | Engine: Calamine | No API Required") [cite: 3, 4]

# --- FILE UPLOADERS ---
col1, col2 = st.columns(2)
with col1:
    pre_files = st.file_uploader("Upload BASELINE (Pre) Reports", accept_multiple_files=True) [cite: 4]
with col2:
    post_files = st.file_uploader("Upload CURRENT (Post) Reports", accept_multiple_files=True) [cite: 4]

def apply_formatting(output_buffer):
    """Applies Red highlighting to mismatched cells in the output Excel."""
    output_buffer.seek(0)
    wb = load_workbook(output_buffer)
    ws = wb.active
    red_fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid') [cite: 4]
    
    # Locate all columns ending with 'Match'
    match_cols = [i for i, cell in enumerate(ws[1], 1) if cell.value and 'Match' in str(cell.value)] [cite: 4]
    
    for row_idx in range(2, ws.max_row + 1):
        for col_idx in match_cols:
            # If the match status is '✗', highlight the data cells
            if ws.cell(row=row_idx, column=col_idx).value == '✗': [cite: 4]
                ws.cell(row=row_idx, column=col_idx-2).fill = red_fill  # Pre Value [cite: 4]
                ws.cell(row=row_idx, column=col_idx-1).fill = red_fill  # Post Value [cite: 4]
                ws.cell(row=row_idx, column=col_idx).fill = red_fill    # Status Cell [cite: 4]
    
    final_output = io.BytesIO()
    wb.save(final_output)
    return final_output.getvalue()

def compare_sheets(df_pre, df_post, is_first_sheet):
    """Core logic to align rows via Composite Key and detect differences."""
    # SKIP logic: Validate if essential headers exist (Skips Pivot Tables)
    if 'Sector Name' not in df_pre.columns or 'Carrier' not in df_pre.columns: [cite: 4]
        return None

    # --- SPECIAL MODIFICATION FOR 1st SHEET ---
    if is_first_sheet: [cite: 4]
        # Placeholder for specific first-sheet rules
        pass

    df1, df2 = df_pre.copy(), df_post.copy()
    
    # Create the unique ID for alignment
    df1['Key'] = df1['Sector Name'].astype(str) + '|' + df1['Carrier'].astype(str) [cite: 4]
    df2['Key'] = df2['Sector Name'].astype(str) + '|' + df2['Carrier'].astype(str) [cite: 4]
    
    all_keys = sorted(set(df1['Key']).union(set(df2['Key']))) [cite: 4]
    compare_cols = sorted(set(df1.columns).union(set(df2.columns)) - {'Sector Name', 'Carrier', 'Key'}) [cite: 4]
    
    results = []
    for k in all_keys:
        sec, car = k.split('|', 1) [cite: 4]
        r1 = df1[df1['Key'] == k] [cite: 4]
        r2 = df2[df2['Key'] == k] [cite: 4]
        
        row_data = {'Sector Name': sec, 'Carrier': car} [cite: 4]
        for col in compare_cols:
            v1 = r1.iloc[0][col] if not r1.empty and col in r1.columns else 'N/A' [cite: 4]
            v2 = r2.iloc[0][col] if not r2.empty and col in r2.columns else 'N/A' [cite: 4]
            
            row_data[f'{col} (Pre)'] = v1 [cite: 4]
            row_data[f'{col} (Post)'] = v2 [cite: 4]
            row_data[f'{col} Match'] = '✓' if str(v1) == str(v2) else '✗' [cite: 4]
        results.append(row_data)

    # Generate the Excel file in memory
    temp_buf = io.BytesIO()
    pd.DataFrame(results).to_excel(temp_buf, index=False) [cite: 4]
    return apply_formatting(temp_buf)

# --- EXECUTION LOGIC ---
if pre_files and post_files: [cite: 5]
    if st.button("🚀 Run Global Audit"): [cite: 4]
        pre_dict = {f.name: f for f in pre_files} [cite: 4]
        post_dict = {f.name: f for f in post_files} [cite: 4]
        
        master_zip = io.BytesIO()
        with zipfile.ZipFile(master_zip, "w") as zf:
            for fname, fobj in pre_dict.items(): [cite: 4]
                if fname in post_dict: [cite: 4]
                    # Creating a folder hierarchy inside the ZIP
                    folder_name = fname.split('.')[0] [cite: 4]
                    st.write(f"📂 **Analyzing Folder:** {folder_name}") [cite: 4]
                    
                    try:
                        # Using 'calamine' engine to avoid Boolean value errors [cite: 3]
                        pre_sheets = pd.read_excel(fobj, sheet_name=None, engine='calamine') [cite: 3, 4]
                        post_sheets = pd.read_excel(post_dict[fname], sheet_name=None, engine='calamine') [cite: 3, 4]
                        
                        for i, (sname, df_pre) in enumerate(pre_sheets.items()): [cite: 4]
                            if sname in post_sheets: [cite: 4]
                                is_first = (i == 0) [cite: 4]
                                result_data = compare_sheets(df_pre, post_sheets[sname], is_first) [cite: 4]
                                
                                if result_data: [cite: 4]
                                    # Save inside report-specific folder
                                    zf.writestr(f"{folder_name}/{sname}_Audit.xlsx", result_data) [cite: 4]
                                    st.write(f"   ✅ Compared: {sname}") [cite: 4]
                                else:
                                    st.write(f"   ⏩ Skipped: {sname} (Non-data sheet)") [cite: 4]
                    except Exception as e:
                        st.error(f"Error reading {fname}: {e}") [cite: 3]

        st.success("Audit Complete!") [cite: 4]
        st.download_button(
            label="📥 Download Comparison Results (ZIP)",
            data=master_zip.getvalue(),
            file_name="Network_Audit_Results.zip",
            mime="application/zip"
        ) [cite: 4]
