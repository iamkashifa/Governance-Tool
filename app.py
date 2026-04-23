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

# ==========================================
# 🎨 1. NEW UI: HERO SECTION & PAGE CONFIG
# ==========================================
st.set_page_config(page_title="PBI Governance Tool", page_icon="📊", layout="wide")

# Optional: If you want an actual image (like a logo or diagram), put the file in your VS code folder and uncomment the line below:
# st.image("your_image_name.png", width=200)

st.title("📊 Dynamic Power BI Governance Audit")
st.markdown("Ensure your enterprise dashboards are clean, consistent, and strictly follow UI/UX brand guidelines—**in seconds, without opening Power BI Desktop.**")
st.divider()

# ==========================================
# 📖 2. NEW UI: HOW IT WORKS (COLLAPSIBLE)
# ==========================================
with st.expander("📖 **How to use this tool & What we check**", expanded=False):
    st.markdown("""
    ### 🛠️ How it works:
    1. **Download the Template:** Click the button below to get the standard Excel rulebook.
    2. **Tweak the Rules:** Change the pixel limits or "Yes/No" rules to fit your department.
    3. **Upload:** Drop your modified Excel file in **Box 1**, and drop your `.pbix` dashboards into **Box 2**.
    4. **Audit:** We unzip the files securely in memory and check every visual against your rules.
    
    ### 🔍 What we are auditing:
    * **Brand Compliance:** Is your logo in the exact top-left corner?
    * **Navigation:** Are buttons/shapes placed at the top for standard navigation?
    * **Filter Zones:** Are slicers floating randomly, or are they anchored Top/Left?
    * **Accessibility:** Do all major charts have standard titles enabled?
    """)
    st.info("💡 *Security Note: Your .pbix files are processed in server memory and instantly deleted. No data is stored.*")

# --- TEMPLATE DOWNLOAD ---
st.download_button(
    label="📥 Download Blank Rule Template (Excel)",
    data=create_template(),
    file_name="Governance_Rules_Template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
st.write("") # Adds a little blank space

# ==========================================
# 📤 3. NEW UI: CLEANER UPLOAD COLUMNS
# ==========================================
st.markdown("### 🚀 Start Your Audit")
col1, col2 = st.columns(2)

with col1:
    st.markdown("#### Step 1: The Rules")
    rules_file = st.file_uploader("Upload Governance Checklist (Excel)", type=["xlsx", "csv"])

with col2:
    st.markdown("#### Step 2: The Dashboards")
    uploaded_files = st.file_uploader("Upload .pbix files or Folder", type="pbix", accept_multiple_files=True)

# --- DEFAULT RULES ---
active_rules = {
    "logo_max_x": 100,
    "logo_max_y": 100,
    "slicer_max_x_for_left": 150,
    "slicer_max_y_for_top": 150,
    "require_visual_titles": True
}

st.divider()

# ==========================================
# ⚙️ 4. CORE ENGINE (Unchanged)
# ==========================================
if uploaded_files:
    # Made the button more prominent
    if st.button("⚡ Run Dynamic Batch Audit", type="primary", use_container_width=True):
        
        # --- AUTO-CONVERT EXCEL TO JSON/DICT ---
        if rules_file is not None:
            try:
                if rules_file.name.endswith('.csv'):
                    df_rules = pd.read_csv(rules_file)
                else:
                    df_rules = pd.read_excel(rules_file)
                
                user_rules = dict(zip(df_rules.iloc[:, 0].str.strip(), df_rules.iloc[:, 1]))
                
                active_rules["logo_max_x"] = int(user_rules.get("Logo Max X Position (Pixels)", 100))
                active_rules["logo_max_y"] = int(user_rules.get("Logo Max Y Position (Pixels)", 100))
                active_rules["slicer_max_x_for_left"] = int(user_rules.get("Slicer Max X Position (Left Zone)", 150))
                active_rules["slicer_max_y_for_top"] = int(user_rules.get("Slicer Max Y Position (Top Zone)", 150))
                
                title_val = str(user_rules.get("Require Titles on Charts (Yes/No)", "Yes")).strip().lower()
                active_rules["require_visual_titles"] = title_val in ['yes', 'true', '1', 'y']
                
                st.success("✅ Custom Excel Checklist Translated and Loaded!")
            except Exception as e:
                st.warning(f"Could not parse rules file, using defaults. Error: {e}")
        else:
            st.warning("No custom checklist uploaded. Using standard strict rules.")

        # --- RUN AUDIT WITH ACTIVE RULES ---
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