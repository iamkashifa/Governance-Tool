import streamlit as st
import zipfile
import json
import re
import pandas as pd
from collections import Counter
from io import BytesIO

# --- HELPER FUNCTIONS ---
def get_color_for_element(config_string, keyword):
    idx = config_string.find(f'"{keyword}"')
    if idx == -1: return "Default/None"
    chunk = config_string[idx : idx + 150]
    match = re.search(r'#[0-9a-fA-F]{6}', chunk)
    return match.group(0).upper() if match else "Default/None"

# Creates a downloadable Excel template for users
def create_template():
    df = pd.DataFrame({
        "Rule Description": [
            "Logo Max X Position (Pixels)", 
            "Logo Max Y Position (Pixels)", 
            "Slicer Max X Position (Left Zone)", 
            "Slicer Max Y Position (Top Zone)", 
            "Require Titles on Charts (Yes/No)"
        ],
        "Value": [100, 100, 150, 150, "Yes"]
    })
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Governance Rules")
    return output.getvalue()

st.set_page_config(page_title="PBI Governance Tool", layout="wide")
st.title("Dynamic Power BI Governance Audit")
st.write("Upload your simple Excel checklist and your `.pbix` files to run a custom audit.")

# --- TEMPLATE DOWNLOAD ---
st.download_button(
    label="📥 Download Blank Rule Template (Excel)",
    data=create_template(),
    file_name="Governance_Rules_Template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
st.divider()

# --- 1. UPLOAD BOXES ---
col1, col2 = st.columns(2)
with col1:
    rules_file = st.file_uploader("1. Upload Governance Checklist (Excel/CSV)", type=["xlsx", "csv"])
with col2:
    uploaded_files = st.file_uploader("2. Upload .pbix files or Folder", type="pbix", accept_multiple_files=True)

# --- 2. DEFAULT RULES ---
active_rules = {
    "logo_max_x": 100,
    "logo_max_y": 100,
    "slicer_max_x_for_left": 150,
    "slicer_max_y_for_top": 150,
    "require_visual_titles": True
}

if uploaded_files:
    if st.button("Run Dynamic Batch Audit"):
        
        # --- 3. AUTO-CONVERT EXCEL TO JSON/DICT ---
        if rules_file is not None:
            try:
                # Read the simple language file
                if rules_file.name.endswith('.csv'):
                    df_rules = pd.read_csv(rules_file)
                else:
                    df_rules = pd.read_excel(rules_file)
                
                # Convert the two columns into a dictionary
                user_rules = dict(zip(df_rules.iloc[:, 0].str.strip(), df_rules.iloc[:, 1]))
                
                # Map simple language to code logic
                active_rules["logo_max_x"] = int(user_rules.get("Logo Max X Position (Pixels)", 100))
                active_rules["logo_max_y"] = int(user_rules.get("Logo Max Y Position (Pixels)", 100))
                active_rules["slicer_max_x_for_left"] = int(user_rules.get("Slicer Max X Position (Left Zone)", 150))
                active_rules["slicer_max_y_for_top"] = int(user_rules.get("Slicer Max Y Position (Top Zone)", 150))
                
                # Convert "Yes"/"True" text into boolean
                title_val = str(user_rules.get("Require Titles on Charts (Yes/No)", "Yes")).strip().lower()
                active_rules["require_visual_titles"] = title_val in ['yes', 'true', '1', 'y']
                
                st.success("✅ Custom Excel Checklist Translated and Loaded!")
            except Exception as e:
                st.warning(f"Could not parse rules file, using defaults. Error: {e}")
        else:
            st.warning("No custom checklist uploaded. Using standard strict rules.")

        # --- 4. RUN AUDIT WITH ACTIVE RULES ---
        st.info(f"Processing {len(uploaded_files)} files...")
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for uploaded_file in uploaded_files:
                dashboard_name = uploaded_file.name.replace('.pbix', '')
                
                try:
                    with zipfile.ZipFile(uploaded_file, 'r') as pbix_zip:
                        with pbix_zip.open('Report/Layout') as layout_file:
                            content = layout_file.read().decode('utf-16-le')
                            report_data = json.loads(content)
                            
                    pages = report_data.get('sections', [])
                    dashboard_results = [] 
                    
                    for page in pages:
                        page_name = page.get('displayName', 'Unknown Page')
                        visuals = page.get('visualContainers', [])
                        
                        found_logo, found_top_nav, found_slicers = False, False, False
                        slicers_are_consistent = True
                        vis_missing_titles = 0
                        
                        for visual in visuals:
                            y_pos = visual.get('y', 999) 
                            x_pos = visual.get('x', 999)
                            config_string = visual.get('config', '{}')
                            
                            try:
                                config_data = json.loads(config_string)
                                v_type = config_data.get('singleVisual', {}).get('visualType', 'Unknown')
                                
                                # --- DYNAMIC RULE: LOGO ---
                                if v_type == 'image' and x_pos < active_rules["logo_max_x"] and y_pos < active_rules["logo_max_y"]: 
                                    found_logo = True
                                    
                                if v_type in ['actionButton', 'shape'] and y_pos < 100: 
                                    found_top_nav = True
                                    
                                # --- DYNAMIC RULE: SLICERS ---
                                if v_type == 'slicer':
                                    found_slicers = True
                                    if x_pos > active_rules["slicer_max_x_for_left"] and y_pos > active_rules["slicer_max_y_for_top"]: 
                                        slicers_are_consistent = False
                                        
                                # --- DYNAMIC RULE: TITLES ---
                                if active_rules["require_visual_titles"]:
                                    if v_type in ['barChart', 'columnChart', 'lineChart', 'pieChart', 'donutChart', 'tableEx', 'pivotTable']:
                                        if "'title':" not in config_string and '"title":' not in config_string:
                                            vis_missing_titles += 1

                            except Exception:
                                continue 

                        dashboard_results.append({
                            "Page Name": page_name,
                            "Logo Check": "✅ Pass" if found_logo else f"❌ Fail (Not in top {active_rules['logo_max_x']}px)",
                            "Filters Layout": "➖ N/A" if not found_slicers else ("✅ Pass" if slicers_are_consistent else "❌ Scattered"),
                            "Visual Titles": "✅ Pass" if vis_missing_titles == 0 else f"❌ Fail ({vis_missing_titles} missing)",
                        })

                    if dashboard_results:
                        df = pd.DataFrame(dashboard_results)
                        safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '', dashboard_name)[:31]
                        df.to_excel(writer, sheet_name=safe_sheet_name, index=False)

                except Exception as e:
                    st.error(f"❌ Could not process {dashboard_name}: {e}")

        st.success("🎉 Dynamic Batch Audit Complete!")
        output.seek(0)
        st.download_button(
            label="Download Custom Governance Audit Report",
            data=output,
            file_name="Custom_Governance_Audit.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )