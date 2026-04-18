import streamlit as st
import os
import zipfile
import json
import re
import pandas as pd
from collections import Counter
from io import BytesIO

# --- HELPER FUNCTION ---
def get_color_for_element(config_string, keyword):
    idx = config_string.find(f'"{keyword}"')
    if idx == -1: return "Default/None"
    chunk = config_string[idx : idx + 150]
    match = re.search(r'#[0-9a-fA-F]{6}', chunk)
    return match.group(0).upper() if match else "Default/None"

st.set_page_config(page_title="PBI Governance Tool", layout="wide")
st.title("Power BI Governance Audit Tool")
st.write("Upload your `.pbix` files below to run an automated compliance check. Your data is processed in memory and not stored.")

# 1. File Uploader
uploaded_files = st.file_uploader("Upload .pbix files", type="pbix", accept_multiple_files=True)

if uploaded_files:
    if st.button("Run Audit"):
        st.info(f"Processing {len(uploaded_files)} files...")
        
        # We will store the Excel data in memory using BytesIO
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for uploaded_file in uploaded_files:
                dashboard_name = uploaded_file.name.replace('.pbix', '')
                st.write(f"Analyzing: {dashboard_name}...")
                
                try:
                    # --- THE HACK: OPEN PBIX AS A ZIP FILE IN MEMORY ---
                    with zipfile.ZipFile(uploaded_file, 'r') as pbix_zip:
                        with pbix_zip.open('Report/Layout') as layout_file:
                            content = layout_file.read().decode('utf-16-le')
                            report_data = json.loads(content)
                            
                    # ... [INSERT ALL YOUR PARSING AND AUDIT LOGIC HERE] ...
                    # (Phase 1: Theme, Phase 2: Audit Pages, calculate inconsistencies, etc.)
                    # Make sure to build the dashboard_results list exactly as your original code did.

                    # NOTE: For brevity, I omitted the loop over sections/visuals here. 
                    # Copy that exact logic from your original script.
                    
                    # Dummy data for demonstration so the script runs without the full logic block
                    dashboard_results = [{"Page Name": "Summary", "Visual Count": 10, "Theme Colors": "✅ Pass"}]
                    
                    # Write to Excel Tab
                    if dashboard_results:
                        df = pd.DataFrame(dashboard_results)
                        safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '', dashboard_name)[:31]
                        df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                        st.success(f"Successfully processed {dashboard_name}")

                except Exception as e:
                    st.error(f"❌ Could not process {dashboard_name}: {e}")

        st.success("Audit Complete!")
        
        # 2. Download Button
        output.seek(0)
        st.download_button(
            label="Download Governance Audit Report (Excel)",
            data=output,
            file_name="Governance_Audit_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )