import streamlit as st
import zipfile
import json
import re
import pandas as pd
from io import BytesIO

# ==========================================
# 🧠 DYNAMIC RULE ENGINE: TEMPLATE GENERATOR
# ==========================================
def create_template():
    # This creates the new "Database Table" style template
    df = pd.DataFrame({
        "Rule Name": ["Logo Top Left X", "Logo Top Left Y", "Slicer Left Zone", "Slicer Top Zone", "Charts Must Have Titles"],
        "Target Visual": ["image", "image", "slicer", "slicer", "barChart"],
        "Property to Check": ["x_position", "y_position", "x_position", "y_position", "title_exists"],
        "Condition": ["Less Than", "Less Than", "Greater Than", "Greater Than", "Equals"],
        "Target Value": [100, 100, 150, 150, "True"],
        "Requirement": ["Must Pass", "Must Pass", "Must Pass", "Must Pass", "Must Pass"]
    })
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Governance Rules")
    return output.getvalue()

# ==========================================
# 🎨 UI: PAGE CONFIG & HERO SECTION
# ==========================================
st.set_page_config(page_title="PBI Governance Engine", page_icon="⚙️", layout="wide")

st.title("⚙️ Dynamic Power BI Rule Engine")
st.markdown("Upload your custom rules matrix and your `.pbix` files. The engine will dynamically adapt to whatever rules you set.")

# Place the template download right at the top so it's easy to grab
st.download_button(
    label="📥 Download Dynamic Rule Template (Excel)",
    data=create_template(),
    file_name="Dynamic_Rules_Template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
st.divider()

# ==========================================
# 📤 1. UPLOADER COLUMNS (Action First)
# ==========================================
st.markdown("### 🚀 Start Your Audit")
col1, col2 = st.columns(2)
with col1:
    rules_file = st.file_uploader("1. Upload Rules Matrix (Excel)", type=["xlsx"])
with col2:
    uploaded_files = st.file_uploader("2. Upload .pbix files or Folder", type="pbix", accept_multiple_files=True)

st.write("") # Small spacing

# ==========================================
# ⚙️ 2. CORE ENGINE (Runs right below uploaders)
# ==========================================
if uploaded_files:
    if st.button("⚡ Run Dynamic Batch Audit", type="primary", use_container_width=True):
        
        # --- 1. LOAD DYNAMIC RULES ---
        if rules_file is not None:
            try:
                df_rules = pd.read_excel(rules_file)
                st.success(f"✅ Loaded {len(df_rules)} dynamic rules from matrix!")
            except Exception as e:
                st.error(f"❌ Failed to load rules: {e}")
                st.stop()
        else:
            st.warning("⚠️ Please upload a Rules Matrix Excel file to use the Dynamic Engine.")
            st.stop()

        # --- 2. RUN AUDIT ---
        st.info(f"🚀 Processing {len(uploaded_files)} files against {len(df_rules)} dynamic rules...")
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
                        
                        # Set up rule trackers for this specific page
                        page_rule_stats = {row['Rule Name']: {"evaluated": 0, "failed": 0} for _, row in df_rules.iterrows()}
                        
                        for visual in visuals:
                            # --- EXTRACT VISUAL PROPERTIES ---
                            y_pos = visual.get('y', 999) 
                            x_pos = visual.get('x', 999)
                            config_string = visual.get('config', '{}')
                            
                            try:
                                config_data = json.loads(config_string)
                                v_type = config_data.get('singleVisual', {}).get('visualType', 'Unknown')
                                
                                has_title = 'true' if ("'title':" in config_string or '"title":' in config_string) else 'false'
                                
                                # --- THE PROPERTY ROUTER ---
                                actual_properties = {
                                    "x_position": x_pos,
                                    "y_position": y_pos,
                                    "title_exists": has_title
                                }
                                
                                # --- EVALUATE VISUAL AGAINST EVERY RULE ---
                                for _, rule in df_rules.iterrows():
                                    target_vis = rule['Target Visual']
                                    prop = rule['Property to Check']
                                    cond = rule['Condition']
                                    target_val = str(rule['Target Value']).strip().lower()
                                    rule_name = rule['Rule Name']
                                    
                                    # Does this rule apply to this visual?
                                    if target_vis.lower() == v_type.lower() or target_vis.lower() == 'all':
                                        page_rule_stats[rule_name]["evaluated"] += 1
                                        actual_val = actual_properties.get(prop)
                                        
                                        passed = True
                                        try:
                                            if cond == 'Less Than':
                                                passed = float(actual_val) < float(target_val)
                                            elif cond == 'Greater Than':
                                                passed = float(actual_val) > float(target_val)
                                            elif cond == 'Equals':
                                                passed = str(actual_val).lower() == target_val
                                        except:
                                            passed = False 
                                            
                                        if not passed:
                                            page_rule_stats[rule_name]["failed"] += 1

                            except Exception:
                                continue 

                        # Compile Results for the Excel Row
                        row_data = {"Page Name": page_name, "Total Visuals": len(visuals)}
                        
                        for rule_name, stats in page_rule_stats.items():
                            if stats["evaluated"] == 0:
                                row_data[rule_name] = "➖ N/A (Visual Not Found)"
                            elif stats["failed"] == 0:
                                row_data[rule_name] = "✅ Pass"
                            else:
                                row_data[rule_name] = f"❌ Fail ({stats['failed']} violations)"
                                
                        dashboard_results.append(row_data)

                    # Write to Excel Tab
                    if dashboard_results:
                        df = pd.DataFrame(dashboard_results)
                        safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '', dashboard_name)[:31]
                        df.to_excel(writer, sheet_name=safe_sheet_name, index=False)

                except Exception as e:
                    st.error(f"❌ Could not process {dashboard_name}: {e}")

        st.success("🎉 Dynamic Batch Audit Complete!")
        output.seek(0)
        st.download_button(
            label="📥 Download Custom Governance Audit Report",
            data=output,
            file_name="Dynamic_Governance_Audit.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.divider()

# ==========================================
# 📚 3. DOCUMENTATION (Moved to the bottom)
# ==========================================
st.markdown("### 📚 Project Documentation")

with st.expander("**Overview**", expanded=False):
    st.markdown("""
    The Dynamic Power BI Governance Audit Tool is an automated compliance engine designed to enforce design standards and best practices across Power BI dashboards.
    
    Auditing Power BI reports for structural consistency (such as checking if logos are in the right place, slicers are aligned, or visuals have titles) previously required manually opening every single `.pbix` file, which is a highly tedious process. This tool automates that entire workflow. By extracting and scanning the underlying JSON metadata of `.pbix` files, it performs batch quality-assurance checks in seconds and generates a detailed compliance report.
    """)

with st.expander("**Key Features**", expanded=False):
    st.markdown("""
    * **In-Memory Batch Processing:** Upload multiple `.pbix` files at once. The tool processes them securely in memory.
    * **Dynamic Rule Configuration:** Download a rule template, add custom rules, adjust pixel boundaries for layout configurations to suit your needs, and upload it back to dynamically drive the audit.
    * **Deep Metadata Parsing:** Treats `.pbix` files as zip archives to unlock the hidden `Report/Layout` structure, scanning exact X/Y coordinates and visual configurations.
    * **Automated Excel Reporting:** Outputs a clean, multi-tab Excel file detailing the exact pass/fail status of every visual, where each tab corresponds to a specific audited dashboard.
    """)

with st.expander("**How It Works (The Architecture)**", expanded=False):
    st.markdown("""
    This tool leverages Python, Pandas, and Streamlit in the backend.
    
    Power BI `.pbix` files are essentially zipped directories. The engine bypasses the Power BI application entirely by unzipping the `.pbix` file in memory and locating the `Report/Layout` file. Because this file is encoded in UTF-16 LE, the script decodes it into a readable JSON format.
    
    Once the JSON tree is exposed, the engine iterates through every section (page) and visualContainer (chart/slicer/shape). It extracts the configuration string to map coordinates (x, y) and visual types against the active governance checklist.
    """)

with st.expander("**How to Use the Tool**", expanded=False):
    st.markdown("""
    1. **Download the Template:** Click the "Download Dynamic Rule Template" button at the top of the page to get the standard governance matrix.
    2. **Customize the Rules:** Open the downloaded Excel file. Instead of hardcoding rules, this app reads your Excel file row by row.
       * **Target Visual:** Which visual type does this rule apply to? (e.g., `image`, `slicer`, `barChart`, or `all`)
       * **Property to Check:** What are we measuring? (Currently supports: `x_position`, `y_position`, `title_exists`)
       * **Condition:** `Less Than`, `Greater Than`, or `Equals`
    3. **Upload the Assets:** * Upload your customized Excel checklist into Box 1.
       * Upload one or more `.pbix` files into Box 2.
    4. **Run the Audit:** Click "Run Dynamic Batch Audit."
    5. **Review Results:** Download the generated `Dynamic_Governance_Audit.xlsx` report to instantly see which dashboards meet company standards.
    """)

with st.expander("**Sample Results of the Existing Checklist**", expanded=False):
    st.markdown("""
    The engine checks for:
    
    * **Logo Placement:** Verifies that an image exists within the strictly defined top-left pixel coordinates.
    * **Navigation Standards:** Ensures action buttons or shapes are utilized in the top navigation zone.
    * **Slicer Alignment:** Flags scattered filters by enforcing strict Top or Left boundary zones.
    * **Accessibility & Clarity:** Checks standard charts (Bar, Line, Pie, etc.) to ensure titles are explicitly enabled in the visual formatting.
    """)
    st.info("💡 **Important:** You can edit the checklist according to your convenience and what you want to check (as in pixels or yes/no).")