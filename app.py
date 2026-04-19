import streamlit as st
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
st.write("Upload your `.pbix` files (or drag and drop an entire folder) below to run an automated compliance check. Your data is processed securely in memory.")

# 1. File Uploader (Accepts multiple files / dragged folders)
uploaded_files = st.file_uploader("Upload .pbix files or Drag & Drop a Folder", type="pbix", accept_multiple_files=True)

if uploaded_files:
    if st.button("Run Batch Audit"):
        st.info(f"Processing {len(uploaded_files)} files...")
        
        # We will store the Excel data in memory
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for uploaded_file in uploaded_files:
                dashboard_name = uploaded_file.name.replace('.pbix', '')
                st.write(f"### Analyzing: {dashboard_name}")
                
                try:
                    # --- OPEN PBIX AS A ZIP FILE IN MEMORY ---
                    with zipfile.ZipFile(uploaded_file, 'r') as pbix_zip:
                        with pbix_zip.open('Report/Layout') as layout_file:
                            content = layout_file.read().decode('utf-16-le')
                            report_data = json.loads(content)
                            
                    # --- PHASE 1: THEME ---
                    approved_colors = {"#FFFFFF", "#000000"} 
                    theme_string = json.dumps(report_data.get('themeCollection', {}))
                    for code in re.findall(r'#[0-9a-fA-F]{6}', theme_string):
                        approved_colors.add(code.upper())

                    # --- PHASE 2: AUDIT PAGES ---
                    pages = report_data.get('sections', [])
                    dashboard_results = [] 
                    
                    for page in pages:
                        page_name = page.get('displayName', 'Unknown Page')
                        visuals = page.get('visualContainers', [])
                        
                        unauthorized_colors = set()
                        found_logo, found_top_nav, found_slicers = False, False, False
                        slicers_are_consistent = True
                        vis_missing_titles, vis_missing_tooltips, charts_with_tooltips = 0, 0, 0
                        page_color_profiles = [] 
                        
                        for visual in visuals:
                            y_pos = visual.get('y', 999) 
                            x_pos = visual.get('x', 999)
                            config_string = visual.get('config', '{}')
                            
                            try:
                                config_data = json.loads(config_string)
                                v_type = config_data.get('singleVisual', {}).get('visualType', 'Unknown')
                                
                                # Color profiling
                                if v_type not in ['shape', 'image', 'actionButton']: 
                                    page_color_profiles.append({
                                        "type": v_type,
                                        "Background": get_color_for_element(config_string, "background"),
                                        "Border": get_color_for_element(config_string, "border"),
                                        "Title": get_color_for_element(config_string, "title"),
                                        "Bars/Data": get_color_for_element(config_string, "fill"),
                                        "Labels": get_color_for_element(config_string, "labels")
                                    })

                                for hex_code in re.findall(r'#[0-9a-fA-F]{6}', config_string):
                                    if hex_code.upper() not in approved_colors:
                                        unauthorized_colors.add(hex_code.upper())
                                        
                                if v_type == 'image' and x_pos < 100 and y_pos < 100: found_logo = True
                                if v_type == 'actionButton' and y_pos < 100: found_top_nav = True
                                if v_type == 'slicer':
                                    found_slicers = True
                                    if x_pos > 150 and y_pos > 150: slicers_are_consistent = False
                                        
                                if v_type in ['barChart', 'columnChart', 'lineChart', 'pieChart', 'donutChart', 'tableEx', 'pivotTable']:
                                    if "'title':" not in config_string and '"title":' not in config_string:
                                        vis_missing_titles += 1
                                        
                                if v_type in ['barChart', 'columnChart', 'lineChart', 'pieChart', 'donutChart', 'scatterChart', 'map', 'treemap']:
                                    charts_with_tooltips += 1
                                    projections = config_data.get('singleVisual', {}).get('projections', {})
                                    if 'tooltips' not in projections or len(projections['tooltips']) == 0:
                                        vis_missing_tooltips += 1

                            except Exception:
                                continue 

                        # Calculate Consistency
                        inconsistencies = 0
                        if page_color_profiles:
                            baseline = {}
                            for category in ["Background", "Border", "Title", "Bars/Data", "Labels"]:
                                all_colors = [p[category] for p in page_color_profiles if p[category] != "Default/None"]
                                baseline[category] = Counter(all_colors).most_common(1)[0][0] if all_colors else "Default/None"
                            
                            for p in page_color_profiles:
                                for category in ["Background", "Border", "Title", "Bars/Data", "Labels"]:
                                    if p[category] != "Default/None" and baseline[category] != "Default/None" and p[category] != baseline[category]:
                                        inconsistencies += 1

                        # Compile Excel Row
                        dashboard_results.append({
                            "Page Name": page_name,
                            "Visual Count": len(visuals),
                            "Theme Colors": "✅ Pass" if not unauthorized_colors else f"❌ Fail ({len(unauthorized_colors)} rogue colors)",
                            "Logo Check": "✅ Pass" if found_logo else "❌ Fail",
                            "Top Nav Check": "✅ Pass" if found_top_nav else "❌ Fail",
                            "Filters Layout": "➖ N/A" if not found_slicers else ("✅ Pass" if slicers_are_consistent else "❌ Fail"),
                            "Visual Titles": "✅ Pass" if vis_missing_titles == 0 else f"❌ Fail ({vis_missing_titles} missing)",
                            "Custom Tooltips": "➖ N/A" if charts_with_tooltips == 0 else ("✅ Pass" if vis_missing_tooltips == 0 else f"❌ Fail ({vis_missing_tooltips} missing)"),
                            "Color Consistency": "➖ N/A" if not page_color_profiles else ("✅ Pass" if inconsistencies == 0 else f"❌ Fail ({inconsistencies} deviations)")
                        })
                        
                        # Print visual feedback to the website screen
                        st.markdown(f"**📄 PAGE: '{page_name}' ({len(visuals)} visuals)**")
                        if unauthorized_colors: st.error(f"🎨 Rogue Colors Found: {len(unauthorized_colors)}")
                        st.write(f"Logo: {'✅ Pass' if found_logo else '❌ Fail'} | Top Nav: {'✅ Pass' if found_top_nav else '❌ Fail'} | Titles: {'✅ Pass' if vis_missing_titles == 0 else '❌ Fail'}")
                        st.divider()

                    # Write to Excel Tab
                    if dashboard_results:
                        df = pd.DataFrame(dashboard_results)
                        safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '', dashboard_name)[:31]
                        df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                        st.success(f"✅ Successfully compiled sheet for: {dashboard_name}")

                except Exception as e:
                    st.error(f"❌ Could not process {dashboard_name}: {e}")

        st.success("🎉 Batch Audit Complete! Click below to download your master report.")
        
        # Download Button
        output.seek(0)
        st.download_button(
            label="Download Master Governance Audit Report (Excel)",
            data=output,
            file_name="Master_Governance_Audit_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )