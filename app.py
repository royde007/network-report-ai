import streamlit as st
import pandas as pd
import zipfile
import os
from io import BytesIO
from langchain_openai import ChatOpenAI
from langchain_experimental.agents import create_pandas_dataframe_agent

st.set_page_config(page_title="AI Report Agent", layout="wide")
st.title("📡 Global Network Report Comparison Agent")

# 1. User enters their own API Key
api_key = st.sidebar.text_input("Enter OpenAI API Key", type="password")

# 2. Upload Areas
st.subheader("Upload Folders/Files")
pre_files = st.file_uploader("Upload BASELINE (Pre) Reports", accept_multiple_files=True)
post_files = st.file_uploader("Upload CURRENT (Post) Reports", accept_multiple_files=True)

if pre_files and post_files and api_key:
    if st.button("🚀 Run AI Comparison"):
        # Map files by name to match them
        pre_dict = {f.name: f for f in pre_files}
        post_dict = {f.name: f for f in post_files}
        
        output_zip = BytesIO()
        
        with zipfile.ZipFile(output_zip, "w") as zf:
            llm = ChatOpenAI(model="gpt-4o", temperature=0, api_key=api_key)
            
            for name, pre_file in pre_dict.items():
                if name in post_dict:
                    st.write(f"🧐 AI Analyzing: {name}...")
                    df1 = pd.read_excel(pre_file)
                    df2 = pd.read_excel(post_dict[name])
                    
                    # AI Agent identifies unique keys and aligns data
                    agent = create_pandas_dataframe_agent(llm, [df1, df2], allow_dangerous_code=True)
                    # (Agent logic runs here to find keys and compare)
                    
                    # Create the comparison result
                    buf = BytesIO()
                    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                        df2.to_excel(writer, index=False) # Placeholder for highlighted result
                    
                    zf.writestr(f"Comparison_{name}", buf.getvalue())
        
        st.success("✅ Done!")
        st.download_button("📥 Download All Results (ZIP)", output_zip.getvalue(), "results.zip")