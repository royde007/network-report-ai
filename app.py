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
    1. **Upload Files**: Select **PRE** and **POST** reports.
    2. **Requirement**: Filenames must match exactly (e.g., *Site_Data.xlsm* in both boxes).
    3. **Key**: Matched via 'Sector Name' + 'Carrier'.
    """)
    st.divider()
    st.caption("v2.9_Debug | Streamlit Logic")

col1, col2 = st.columns(2)
with col1:
    pre_files = st.file_uploader("Upload PRE Reports", accept_multiple_files=True)
with col2:
    post_files = st.file_uploader("Upload POST Reports", accept_multiple_files=True)

# ... [streaming_load and create_comparison_report functions remain the same as previous version] ...

if st.button("🚀 Run Global Audit"):
    if not pre_files or not post_files:
        st.error("❌ Please upload files in both PRE and POST boxes first.")
    else:
        pre_dict = {f.name: f for f in pre_files}
        post_dict = {f.name: f for f in post_files}
        
        # DEBUG: Show detected files
        st.write(f"Detected PRE files: {list(pre_dict.keys())}")
        st.write(f"Detected POST files: {list(post_dict.keys())}")

        zip_buffer = io.BytesIO()
        processed_count = 0

        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for fname, fobj in pre_dict.items():
                if fname in post_dict:
                    st.info(f"🔄 Starting Audit for: {fname}")
                    xl = pd.ExcelFile(fobj, engine='openpyxl')
                    for i, sname in enumerate(xl.sheet_names):
                        if i == 0 or sname.lower().endswith('pivot') or sname == "General Information":
                            continue
                        
                        df_pre = streaming_load(fobj, sname)
                        df_post = streaming_load(post_dict[fname], sname)
                        
                        if df_pre is not None and df_post is not None:
                            report = create_comparison_report(df_pre, df_post, sname)
                            zf.writestr(f"{fname.split('.')[0]}/{sname}_Audit.xlsx", report)
                            st.write(f"✅ Processed Sheet: {sname}")
                            processed_count += 1
                else:
                    st.warning(f"⚠️ Skipping '{fname}' because a matching file wasn't found in the POST box.")

        if processed_count > 0:
            st.success(f"🏁 Audit Complete! Processed {processed_count} sheets.")
            st.download_button("📥 Download Results", zip_buffer.getvalue(), "Network_Audit_Results.zip")
        else:
            st.error("❌ No sheets were processed. Check if 'Sector Name' and 'Carrier' headers exist.")
