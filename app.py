import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Network Report Analyzer", layout="wide")
st.title("📡 Sector & Carrier Comparison Tool")

# Sidebar for instructions
st.sidebar.header("Instructions")
st.sidebar.info("1. Upload Pre and Post files with matching names.\n2. The tool matches rows using 'Sector Name' + 'Carrier'.\n3. Mismatches are highlighted in Red.")

# File Uploaders
pre_files = st.file_uploader("Upload BASELINE (Pre) Reports", accept_multiple_files=True)
post_files = st.file_uploader("Upload CURRENT (Post) Reports", accept_multiple_files=True)

def process_comparison(file1, file2):
    """Merged Logic: Composite Key Matching + Highlighting"""
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)
    
    # Create composite key as per your script
    df1['Composite_Key'] = df1['Sector Name'].astype(str) + '|' + df1['Carrier'].astype(str)
    df2['Composite_Key'] = df2['Sector Name'].astype(str) + '|' + df2['Carrier'].astype(str)
    
    all_keys = sorted(set(df1['Composite_Key']).union(set(df2['Composite_Key'])))
    all_columns = sorted(set(df1.columns).union(set(df2.columns)) - {'Sector Name', 'Carrier', 'Composite_Key'})
    
    comparison_data = []
    for key in all_keys:
        sector_name, carrier = key.split('|', 1)
        record1 = df1[df1['Composite_Key'] == key]
        record2 = df2[df2['Composite_Key'] == key]
        
        row = {'Sector Name': sector_name, 'Carrier': carrier}
        all_match = True
        
        for col in all_columns:
            val1 = record1.iloc[0][col] if not record1.empty and col in record1.columns else 'N/A'
            val2 = record2.iloc[0][col] if not record2.empty and col in record2.columns else 'N/A'
            
            # Handle Nulls
            val1 = 'NULL' if pd.isna(val1) else val1
            val2 = 'NULL' if pd.isna(val2) else val2
            
            row[f'{col} (Pre)'] = val1
            row[f'{col} (Post)'] = val2
            match_val = '✓' if str(val1) == str(val2) else '✗'
            row[f'{col} Match'] = match_val
            if match_val == '✗': all_match = False
            
        row['Overall_Match'] = 'MATCH' if all_match else 'MISMATCH'
        comparison_data.append(row)

    # Create Excel in Memory
    output = io.BytesIO()
    comp_df = pd.DataFrame(comparison_data)
    comp_df.to_excel(output, index=False, engine='openpyxl')
    
    # Apply your 'apply_highlighting_detailed' logic
    output.seek(0)
    wb = load_workbook(output)
    ws = wb.active
    red_fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
    
    # Find columns with 'Match' in header
    match_cols = [i for i, cell in enumerate(ws[1], 1) if cell.value and 'Match' in str(cell.value)]
    
    for row_idx in range(2, ws.max_row + 1):
        for col_idx in match_cols:
            if ws.cell(row=row_idx, column=col_idx).value == '✗':
                # Highlight Pre, Post, and Match cells
                ws.cell(row=row_idx, column=col_idx-2).fill = red_fill
                ws.cell(row=row_idx, column=col_idx-1).fill = red_fill
                ws.cell(row=row_idx, column=col_idx).fill = red_fill
                
    final_output = io.BytesIO()
    wb.save(final_output)
    return final_output.getvalue()

if pre_files and post_files:
    if st.button("🚀 Run Comparison"):
        pre_dict = {f.name: f for f in pre_files}
        post_dict = {f.name: f for f in post_files}
        
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for name, file in pre_dict.items():
                if name in post_dict:
                    st.write(f"Comparing: {name}...")
                    result = process_comparison(file, post_dict[name])
                    zf.writestr(f"Compared_{name}", result)
        
        st.success("✅ Comparison Complete!")
        st.download_button("📥 Download Results (ZIP)", zip_buffer.getvalue(), "comparison_results.zip")
