# Power BI Governance Audit Tool

A lightweight, automated web application built with Streamlit that audits Power BI dashboards (`.pbix` files) against standard governance and UI/UX design checklists. 

**[🔗 Try the Live App Here](https://governance-tool-nauez4jeschpdl6fhbkrsm.streamlit.app/)**

---

## What It Does
Manually opening and reviewing dozens of Power BI pages for compliance is incredibly time-consuming. This tool automates the process by analyzing the underlying JSON configuration of `.pbix` files to check for structural consistency without requiring Power BI Desktop to be installed.

Currently, the engine audits every page for:
* **Brand Compliance (Logo):** Verifies that a corporate logo (Image visual) is placed in the top-left corner (X < 100, Y < 100).
* **Standardized Navigation:** Checks for the presence of Navigation elements (Buttons/Shapes) at the top of the canvas.
* **Filter/Slicer Placement:** Ensures slicers are consistently placed in standard zones (Top or Left) and flags scattered slicers.
* **Accessibility (Titles):** Scans all major charts (bars, lines, pies, tables) to ensure standard titles are enabled.

---

## How It Works (Under the Hood)
The tool utilizes a "Zip Hack" to process files securely and rapidly:
1.  **In-Memory Unzipping:** A `.pbix` file is essentially a renamed `.zip` archive. The script unzips the uploaded file directly in server memory.
2.  **JSON Parsing:** It locates and reads the hidden `Report/Layout` file, which is encoded in `utf-16-le`.
3.  **Coordinate Mapping:** It iterates through the `sections` (pages) and `visualContainers` (visuals), parsing the deeply nested configuration strings to extract `x/y` coordinates and visual types.
4.  **Excel Generation:** Results are tabulated using `pandas` and written to a downloadable `.xlsx` file using `openpyxl`.

*Note: User data is processed securely in memory and is immediately discarded after the session. No data is stored on the server.*

---

## How to Use the Web App
1. Navigate to the **[Live App](https://governance-tool-nauez4jeschpdl6fhbkrsm.streamlit.app/)**.
2. Drag and drop one or multiple `.pbix` files into the uploader.
3. Click **Run Audit**.
4. View the page-by-page breakdown on the screen.
5. Click **Download Governance Audit Report (Excel)** to get your tabulated compliance checklist.

---

## Local Setup & Installation

If you wish to clone this repository and run the tool locally on your own machine:

**1. Clone the repository:**
```bash
git clone [https://github.com/iamkashifa/Governance-Tool.git](https://github.com/iamkashifa/Governance-Tool.git)
cd Governance-Tool
2. Install the required dependencies:

Bash
pip install -r requirements.txt
3. Launch the Streamlit app:

Bash
streamlit run app.py
📂 Project Structure
app.py - The core Streamlit web application and .pbix parsing engine.

requirements.txt - Python dependencies (streamlit, pandas, openpyxl).

.gitignore - Prevents raw .pbix and output .xlsx files from being tracked by version control.

Created by Kashifa Bano 
