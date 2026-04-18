import streamlit as st
import zipfile
import json
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="PBI Governance Tool", layout="wide")
st.title("Power BI Governance Audit Tool")
st.write("Upload your `.pbix` files below to run an automated compliance check. Your data is processed securely in memory.")

# File Uploader
uploaded_files = st.file_uploader("Upload .pbix files", type="pbix", accept_multiple_files=True)

if uploaded_files:
    if st.button("Run Audit"):
        st.info(f"Processing {len(uploaded_files)} files...")
        
        # We will store the Excel data in memory using BytesIO
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for uploaded_file in uploaded_files:
                dashboard_name = uploaded_file.name.replace('.pbix', '')
                st.write(f"### Analyzing: {dashboard_name}")
                
                try:
                    # --- OPEN PBIX AS A ZIP FILE IN MEMORY ---
                    with zipfile.ZipFile(uploaded_file, 'r') as pbix_zip:
                        with pbix_zip.open('Report/Layout') as layout_file:
                            # Power BI Layout files are typically encoded in utf-16-le inside the pbix
                            content = layout_file.read().decode('utf-16-le')
                            report_data = json.loads(content)
                            
                    pages = report_data.get('sections', [])
                    dashboard_results = []
                    
                    # --- YOUR NEW AUDIT LOGIC ---
                    for page in pages:
                        page_name = page.get('displayName', 'Unknown Page')
                        visuals = page.get('visualContainers', [])
                        
                        found_logo = False
                        found_top_nav = False
                        found_slicers = False
                        slicers_are_consistent = True
                        visuals_missing_titles = 0
                        
                        for visual in visuals:
                            # Get coordinates
                            y_pos = visual.get('y', 999) 
                            x_pos = visual.get('x', 999)
                            
                            # Extract the hidden config data
                            config_string = visual.get('config', '{}')
                            try:
                                config_data = json.loads(config_string)
                                visual_type = config_data.get('singleVisual', {}).get('visualType', 'Unknown')
                                
                                # --- RULE 1: THE LOGO ---
                                if visual_type == 'image' and x_pos < 100 and y_pos < 100:
                                    found_logo = True
                                    
                                # --- RULE 2: TOP NAVIGATION ---
                                if visual_type in ['actionButton', 'shape'] and y_pos < 100:
                                    found_top_nav = True
                                    
                                # --- RULE 3: CONSISTENT FILTERS ---
                                if visual_type == 'slicer':
                                    found_slicers = True
                                    if x_pos > 150 and y_pos > 150:
                                        slicers_are_consistent = False
                                        
                                # --- RULE 4: VISUAL TITLES ---
                                charts_requiring_titles = ['barChart', 'columnChart', 'lineChart', 'pieChart', 'donutChart', 'tableEx', 'pivotTable']
                                if visual_type in charts_requiring_titles:
                                    if "'title':" not in config_string and '"title":' not in config_string:
                                        visuals_missing_titles += 1

                            except Exception:
                                continue # Skip visuals with broken configs
                                
                        # Print visual feedback to the website screen
                        st.markdown(f"**📄 PAGE: '{page_name}' ({len(visuals)} visuals)**")
                        st.write(f"Logo: {'✅ Present' if found_logo else '❌ Missing'}")
                        st.write(f"Top Nav: {'✅ Present' if found_top_nav else '❌ Missing'}")
                        st.write(f"Filters: {'✅ Consistent' if (found_slicers and slicers_are_consistent) else ('❌ Scattered' if found_slicers else '➖ None')}")
                        st.write(f"Titles: {'✅ All Present' if visuals_missing_titles == 0 else f'❌ {visuals_missing_titles} Missing'}")
                        st.divider()

                        # Save results to our Excel table data
                        dashboard_results.append({
                            "Page Name": page_name,
                            "Total Visuals": len(visuals),
                            "Logo Present": "Yes" if found_logo else "No",
                            "Top Nav Present": "Yes" if found_top_nav else "No",
                            "Slicers Consistent": "Yes" if (found_slicers and slicers_are_consistent) else ("No" if found_slicers else "N/A"),
                            "Missing Titles": visuals_missing_titles
                        })
                    
                    # Write to Excel Tab
                    if dashboard_results:
                        df = pd.DataFrame(dashboard_results)
                        safe_sheet_name = dashboard_name[:31] # Excel limits tab names to 31 chars
                        df.to_excel(writer, sheet_name=safe_sheet_name, index=False)

                except Exception as e:
                    st.error(f"❌ Could not process {dashboard_name}: {e}")

        st.success("Audit Complete! Click below to download your report.")
        
        # Download Button
        output.seek(0)
        st.download_button(
            label="Download Governance Audit Report (Excel)",
            data=output,
            file_name="Governance_Audit_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )