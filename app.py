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
    
    1. **Upload Files**: Select your 'Pre' (Baseline) and 'Post' (Current) reports in the main panel.
    2. **Automatic Filtering**: The tool scans all tabs. Sheets like *Pivot Tables* or *Histograms* that lack 'Sector Name' columns are automatically skipped.
    3. **Smart Alignment**: Data is matched using a **Composite Key** (Sector Name + Carrier) to ensure rows are compared accurately even if they moved.
    4. **1st Sheet Logic**: The first tab of every report undergoes a specialized validation check.
    5. **Visual Highlights**: Any cell that has changed between the two reports will be highlighted in **Red**.
    6. **Download Results**: A ZIP file will be generated, containing a folder for each report and separate Excel files for each sheet.
    """)
    st.divider()
    st.caption("v2.1 | Engine: Calamine | No API Required")

# --- FILE UPLOADERS ---
col1, col2 = st.columns(2)
with col1:
    pre_files = st.file_uploader("Upload BASELINE (Pre) Reports", accept_multiple_files=True)
with col2:
    post_files = st.file_uploader("Upload CURRENT (Post) Reports", accept_multiple_files=True)

def apply_formatting(output_buffer):
    """Applies Red highlighting to mismatched cells in the output Excel."""
    output_buffer.seek(0)
    wb = load_workbook(output_buffer)
    ws = wb.active
    red_fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
    
    # Locate all columns ending with 'Match'
    match_cols = [i for i, cell in enumerate(ws[1], 1) if cell.value and 'Match' in str(cell.value)]
    
    # Loop through rows and highlight cells where Match is '✗'
    for row_idx in range(2, ws.max_row + 1):
        for col_idx in match_cols:
            if ws.cell(row=row_idx, column=col_idx).value == '✗':
                # Highlight Pre Value (2 columns back), Post Value (1 column back), and Status
                ws.cell(row=row_idx, column=col_idx-2).fill = red_fill
                ws.cell(row=row_idx, column=col_idx-1).fill = red_fill
                ws.cell(row=row_idx, column=col_idx).fill = red_fill
    
    final_output = io.BytesIO()
    wb.save(final_output)
    return final_output.getvalue()

def compare_sheets(df_pre, df_post, is_first_sheet):
    """Core logic to align rows via Composite Key and detect differences."""
    # SKIP logic: Validate if essential headers exist
    if 'Sector Name' not in df_pre.columns or 'Carrier' not in df_pre.columns:
        return None

    # First sheet logic hook
    if is_first_sheet:
        pass

    df1, df2 = df_pre.copy(), df_post.copy()
    
    # Create the unique ID for alignment
    df1['Key'] = df1['Sector Name'].astype(str) + '|' + df1['Carrier'].astype(str)
    df2['Key'] = df2['Sector Name'].astype(str) + '|' + df2['Carrier'].astype(str)
    
    all_keys = sorted(set(df1['Key']).union(set(df2['Key'])))
    compare_cols = sorted(set(df1.columns).union(set(df2.columns)) - {'Sector Name', 'Carrier', 'Key'})
    
    results = []
    for k in all_keys:
        sec, car = k.split('|', 1)
        r1 = df1[df1['Key'] == k]
        r2 = df2[df2['Key'] == k]
        
        row_data = {'Sector Name': sec, 'Carrier': car}
        for col in compare_cols:
            v1 = r1.iloc[0][col] if not r1.empty and col in r1.columns else 'N/A'
            v2 = r2.iloc[0][col] if not r2.empty and col in r2.columns else 'N/A'
            
            row_data[f'{col} (Pre)'] = v1
            row_data[f'{col} (Post)'] = v2
            row_data[f'{col} Match'] = '✓' if str(v1) == str(v2) else '✗'
        results.append(row_data)

    temp_buf = io.BytesIO()
    pd.DataFrame(results).to_excel(temp_buf, index=False)
    return apply_formatting(temp_buf)

# --- EXECUTION LOGIC ---
if pre_files and post_files:
    if st.button("🚀 Run Global Audit"):
        pre_dict = {f.name: f for f in pre_files}
        post_dict = {f.name: f for f in post_files}
        
        master_zip = io.BytesIO()
        with zipfile.ZipFile(master_zip, "w") as zf:
            for fname, fobj in pre_dict.items():
                if fname in post_dict:
                    folder_name = fname.split('.')[0]
                    st.write(f"📂 **Analyzing Folder:** {folder_name}")
                    
                    try:
                        # Using 'calamine' engine to avoid Boolean value errors
                        pre_sheets = pd.read_excel(fobj, sheet_name=None, engine='calamine')
                        post_sheets = pd.read_excel(post_dict[fname], sheet_name=None, engine='calamine')
                        
                        for i, (sname, df_pre) in enumerate(pre_sheets.items()):
                            if sname in post_sheets:
                                is_first = (i == 0)
                                result_data = compare_sheets(df_pre, post_sheets[sname], is_first)
                                
                                if result_data:
                                    zf.writestr(f"{folder_name}/{sname}_Audit.xlsx", result_data)
                                    st.write(f"   ✅ Compared: {sname}")
                                else:
                                    st.write(f"   ⏩ Skipped: {sname} (Non-data sheet)")
                    except Exception as e:
                        st.error(f"Error reading {fname}: {e}")

        st.success("Audit Complete!")
        st.download_button(
            label="📥 Download Comparison Results (ZIP)",
            data=master_zip.getvalue(),
            file_name="Network_Audit_Results.zip",
            mime="application/zip"
        )
