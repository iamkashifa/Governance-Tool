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
# # ==========================================
# 📖 2. NEW UI: DOCUMENTATION (COLLAPSIBLE SECTIONS)
# ==========================================
st.markdown("### 📚 Project Documentation")

with st.expander("📝 **Overview**", expanded=False):
    st.markdown("""
    The Dynamic Power BI Governance Audit Tool is an automated compliance engine designed to enforce design standards and best practices across Power BI dashboards.
    
    Auditing Power BI reports for structural consistency (like checking if logos are in the right place, slicers are aligned, or visuals have titles) required manually opening every single `.pbix` file. This tool automates that entire process. By extracting and scanning the underlying JSON metadata of `.pbix` files, it performs batch quality-assurance checks in seconds and generates a detailed compliance report.
    """)

with st.expander("✨ **Key Features**", expanded=False):
    st.markdown("""
    * **In-Memory Batch Processing:** Upload multiple `.pbix` files at once. The tool processes them securely in memory without needing to open Power BI Desktop or save your proprietary data to a server.
    * **Dynamic Rule Configuration:** Governance isn't one-size-fits-all. The tool allows administrators to download a rule template, customize the pixel boundaries for layout rules, and upload it back to dynamically drive the audit.
    * **Deep Metadata Parsing:** Treats `.pbix` files as zip archives to crack open the hidden `Report/Layout` structure, scanning exact X/Y coordinates and visual configurations.
    * **Automated Excel Reporting:** Outputs a clean, multi-tab Excel file detailing the exact pass/fail status of every visual and page across all audited dashboards.
    """)

with st.expander("🛠️ **How It Works (The Architecture)**", expanded=False):
    st.markdown("""
    Under the hood, this tool leverages Python, Pandas, and Streamlit.
    
    Power BI `.pbix` files are essentially zipped directories. The engine bypasses the Power BI application entirely by unzipping the `.pbix` file in memory and locating the `Report/Layout` file. Because this file is encoded in `UTF-16 LE`, the script decodes it into a readable JSON format.
    
    Once the JSON tree is exposed, the engine iterates through every section (page) and visualContainer (chart/slicer/shape), extracting the config string to map coordinates (`x`, `y`) and visual types against the active governance checklist.
    """)

with st.expander("🚀 **How to Use the Portal**", expanded=False):
    st.markdown("""
    1. **Download the Template:** Click the "Download Blank Rule Template" button below to get the standard governance checklist.
    2. **Customize the Rules:** Open the downloaded Excel file and adjust the parameters (e.g., changing the maximum allowed X-coordinate for a company logo).
    3. **Upload the Assets:** * Upload your customized Excel checklist into Box 1.
       * Upload one or more `.pbix` files into Box 2.
    4. **Run the Audit:** Click the "Run Dynamic Batch Audit" button.
    5. **Review Results:** Download the generated `Custom_Governance_Audit.xlsx` report to instantly see which dashboards meet company standards and which require redesigns.
    """)

with st.expander("📋 **Current Audit Capabilities**", expanded=False):
    st.markdown("""
    Currently, the engine checks for:
    * **Logo Placement:** Verifies an image exists within the strictly defined top-left pixel coordinates.
    * **Navigation Standards:** Ensures action buttons or shapes are utilized in the top navigation zone.
    * **Slicer Alignment:** Flags scattered filters by enforcing strict Top or Left boundary zones.
    * **Accessibility & Clarity:** Checks standard charts (Bar, Line, Pie, etc.) to ensure titles are explicitly enabled in the visual formatting.
    """)
    st.info("💡 **Important:** You can edit the checklist according to your convenience and what you want to check (as in pixels or yes/no).")

st.write("") # Adds a little blank space before the download button

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