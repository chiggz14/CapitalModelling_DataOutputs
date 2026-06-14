import streamlit as st
import pandas as pd
import numpy as np
import xlwings as xw
import win32com.client as win32
import time
from openpyxl import load_workbook
import os
import shutil
import pythoncom

# --- --- --- TITLE --- --- --- 
st.title("Capital Modelling - Output File Creator")

with st.expander("About"):
    st.markdown(
        """ 
        This is a tool to refresh output templates for a selected model run from the Igloo Model 
        
        Select the Project Run and Output Templates you would like to be created.
        This will then create an output folder and refresh and save down versions of these templates  
        """
    )





# --- --- --- Variables --- --- ---
# This section stores any manual/hardcoded variables such as file paths and table names which can be updated
file_path = r"\\int.westfieldspecialty.net\Dept\Agency\Actuarial\Capital\Capital Modelling\Other working files\Chirag\5. Python\Excel Files\Igloo Runs.xlsm"
base_path = r"\\int.westfieldspecialty.net\Dept\Agency\Actuarial\Capital\Capital Modelling\Other working files\Chirag\5. Python\Excel Files"
template_path = r"\\int.westfieldspecialty.net\Dept\Agency\Actuarial\Capital\2. Data Outputs\2026 SCR March CiL\0. Output Templates"
table_name = "_Igloo_Tbl_Runs"

with st.expander("Variables"):
    st.markdown(
        """ 
        This section contains the list of variables that are manually stored in the Python script 
        """
    )

    st.markdown(f"""
    **Igloo Run FilePath:** `{file_path}`  
    **Output FilePath:** `{base_path}`  
    **Template FilePath:** `{template_path}`
    """)


# --- --- --- SECTION 1 --- --- --- 
# --- --- --- Igloo Run Selector --- --- ---
# Step 1: This step imports the Igloo Runs from an excel file which contains the list of Igloo Runs. This is used to select the runs to output files in a later step
#HEADER
st.header("**Run Selector:**")
st.markdown("Select the run you want to output:")

@st.cache_data
def load_excel_table(file_path, table_name):
    wb = load_workbook(file_path, data_only=True)

    for ws in wb.worksheets:
        if table_name in ws.tables:
            table = ws.tables[table_name]

            # Get range (e.g. A1:D20)
            data = ws[table.ref]

            # Extract values
            rows = list(data)
            headers = [cell.value for cell in rows[0]]
            values = [[cell.value for cell in row] for row in rows[1:]]

            df = pd.DataFrame(values, columns=headers)
            return df

    raise ValueError(f"Table '{table_name}' not found")


# --- Load into dataframe ---
# This step loads the Igloo Run table into a dataframe for Python to then use
try:
    df = load_excel_table(file_path, table_name)

except Exception as e:
    st.error(str(e))

# Clean column names
df.columns = df.columns.str.strip()

# --- Text input ---
search_text = st.text_input("Search ProjectRunName")

# --- Filter logic ---
# This step allows the user to type a string in the text box to filter the run table list
if search_text:
    filtered_df = df[
        df["ProjectRunName"].astype(str).str.contains(search_text, case=False, na=False)
    ]
else:
    filtered_df = df

# --- Display ---
st.dataframe(filtered_df[["ProjectRunName", "ModelSchema"]])


# Dropdown of unique ProjectRunName values to be used as the run selection
# This step uses the filtered dataframe to allow the user to select the run from the filtered list
selected_project = st.selectbox(
    "Select Project Run",
    filtered_df["ProjectRunName"].dropna().unique()
)

st.write(f"Selected ProjectRunName: {selected_project}")

folder_name = selected_project.split("\\")[-1]
st.session_state.destination_path = os.path.join(base_path, folder_name)





# --- --- --- SECTION 2 --- --- --- 
# --- --- --- File Selector --- --- --- 
#Step 2: This step will allow the user to select which output files would be required to copy from the template folder into the run output folder and then refresh
st.header("Select the files you want to output:")

# --- Checklist items ---
checklist_items = {
    "LCR Form": False,
    "Reconciliation File": False,
    "UW Dashboard": False,
    "Detailed Outputs": False,
    "SBF Template": False
}

# --- Display checkboxes and store values ---
for item in checklist_items:
    checklist_items[item] = st.checkbox(item)





# --- --- --- SECTION 3 --- --- --- 
# --- --- --- CREATE FOLDER --- --- ---
# Step 3: This step creates an output folder for the run name in the required file path location 
def create_project_folder(base_path, project_run_name):
    # --- Extract text after last "\" ---
    folder_name = project_run_name.split("\\")[-1]

    # --- Build full path ---
    new_folder_path = os.path.join(base_path, folder_name)

    # --- Check if folder exists ---
    if os.path.exists(new_folder_path):
        print(f"Folder already exists: {new_folder_path}")
        return None  # Stop execution

    # --- Create folder ---
    os.makedirs(new_folder_path)
    print(f"Created folder: {new_folder_path}")

    return new_folder_path

if st.button("Create Folder"):
    result = create_project_folder(base_path, selected_project)

    if result is None:
        st.warning("Folder already exists")
        st.stop()
    else:
        st.session_state.destination_path = result
        st.success(f"Folder created: {result}")





# --- --- --- SECTION 4 --- --- --- 
# --- --- --- Output File Creator --- --- ---
# Step 4: This step creates the output files from the template folder and pastes in the required destination run folder
def copy_selected_templates(template_path, destination_path, checklist, table_name, macro_ClearSelection, macro_LoadData):
    # Mapping from checklist names → actual file names
    file_mapping = {
        "LCR Form": "LCR.xlsm",
        "Reconciliation File": "Reconciliation.xlsm",
        "UW Dashboard": "UW Dashboard.xlsm",
        "Detailed Outputs": "Detailed Outputs.xlsm",
        "SBF Template": "SBF Template.xlsm"
    }

    copied_files = []

    for item, selected in checklist.items():
        if selected:
            file_name = file_mapping.get(item)

            if not file_name:
                print(f"No mapping for {item}")
                continue

            source = os.path.join(template_path, file_name)
            destination = os.path.join(destination_path, file_name)

            if not os.path.exists(source):
                print(f"Missing file: {source}")
                continue

            shutil.copy2(source, destination)
            copied_files.append(file_name)
            print(f"Workbook Copied")

    # --- Step 2: Open, Save, Close each file ---
    if not copied_files:
        return [], "No files copied"

    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False
    UpdateLinks=0  # 0 = don't prompt, don't update automatically
    
    processed_files = []
    failed_files = []

    try:
        for rng_FilePath_Template_Name in copied_files:
            try:
                
                rng_FilePath_Template = os.path.join(destination_path, rng_FilePath_Template_Name)
                
                print(rng_FilePath_Template)

                wb = excel.Workbooks.Open(rng_FilePath_Template)
                st.success(f"Workbook Opened: {rng_FilePath_Template}")
                
                
                # --- Run Macro ---
                try:
                    print(f"Running macro: {macro_ClearSelection}")
                    # If macro is in this workbook:
                    excel.Application.Run(f"'{wb.Name}'!{macro_ClearSelection}")
                except Exception as e:
                    raise RuntimeError(f"Failed to run macro '{macro_ClearSelection}': {e}")

                # --- Refresh Table ---
                found = False
                for sheet in wb.Worksheets:
                    for table in sheet.ListObjects:
                        if table.Name == table_name:
                            st.success(f"Refreshing table: {table_name}")
                            table.QueryTable.Refresh(False)  # synchronous
                            found = True
                            break
                    if found:
                        break

                if not found:
                    raise ValueError(f"Table '{table_name}' not found.")

                # --- Wait for all queries to finish ---
                excel.CalculateUntilAsyncQueriesDone()
                st.success("Calculate until queries complete")

                # --- Copy B141 from '0. Setup' to '_Lookups'!T5 ---
                ws_setup = wb.Worksheets("0. Setup")
                ws_lookup = wb.Worksheets("_Lookups")

                value_to_copy = selected_project
                st.success(f"Copying value {value_to_copy} to _Lookups!T5")

                ws_lookup.Range("T5").Value = value_to_copy

                # Optional: force calc if these feed PQ parameters
                excel.Calculate()
                st.success("Recalculate File")

                # --- Run Macro ---
                try:
                    print(f"Running macro: {macro_LoadData}")
                    # If macro is in this workbook:
                    excel.Application.Run(f"'{wb.Name}'!{macro_LoadData}")
                except Exception as e:
                    raise RuntimeError(f"Failed to run macro '{macro_LoadData}': {e}")


                st.success("Refreshing Files: 30 Seconds")

                time.sleep(30)



                # Optional: force calc if these feed PQ parameters
                excel.Calculate()
                print("Recalculate File")

                wb.Save()
                st.success("Workbook Saved") 
                wb.Close()
                st.success("Workbook Closed")

                processed_files.append(os.path.basename(rng_FilePath_Template))

            except Exception as e:
                failed_files.append((rng_FilePath_Template, str(e)))

    finally:
        excel.Quit()

    return processed_files, failed_files

st.write(st.session_state)

if st.button("Copy Selected Templates"):

    # Ensure folder was created first
    if "destination_path" not in st.session_state:
        st.error("Please create the folder first")
        st.stop()

    processed, failed = copy_selected_templates(
        template_path,
        st.session_state.destination_path,
        checklist_items,
        table_name="_Igloo_Tbl_Runs",
        macro_ClearSelection="st_ClearExistingSelection",
        macro_LoadData = "st_RefreshDataloadQueries",
    )

    if processed:
        st.success(f"Processed files: {processed}")

    if failed:
        st.error(f"Failed files: {failed}")

    if not processed and not failed:
        st.warning("No files selected")



