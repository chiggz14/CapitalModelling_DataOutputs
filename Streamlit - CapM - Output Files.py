import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import load_workbook
import shutil
import pyodbc


# py -m streamlit run "\\int.westfieldspecialty.net\Dept\Agency\Actuarial\Capital\Capital Modelling\Other working files\Chirag\5. Python\Streamlit - CapM - Output Files.py"
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
base_path = r"\\int.westfieldspecialty.net\Dept\Agency\Actuarial\Capital\Capital Modelling\Other working files\Chirag\5. Python\Excel Files"
template_path = r"\\int.westfieldspecialty.net\Dept\Agency\Actuarial\Capital\2. Data Outputs\2026 SCR March CiL\0. Output Templates"
default_folderpath = r"\\int.westfieldspecialty.net\Dept\Agency\Actuarial\Capital\2. Data Outputs\2027 SCR"
default_database = "ICM 5_8_X"
default_rungroup = "2027 S1200 - 1. SCR"
table_name = "_Igloo_Tbl_Runs"

with st.expander("Variables"):
    st.markdown(
        """ 
        This section contains the list of variables that are manually stored in the Python script 
        """
    )















# --- --- --- SECTION 1 --- --- --- 
# --- --- --- Igloo Run Selector --- --- ---
# Step 1: This step imports the Igloo Runs from an excel file which contains the list of Igloo Runs. This is used to select the runs to output files in a later step
#HEADER
st.header("**Run Selector:**", divider='rainbow')
st.markdown("Select the run you want to output:")

# --- CONNECTION FUNCTION ---
# --- This function sets up the connection to the SQLSERVER SWPISQL11 ---
def get_connection(database=None):
    server = "SWPISQL11"

    # Windows Authentication (change if using SQL auth)
    conn_str = f"""
        DRIVER={{ODBC Driver 17 for SQL Server}};
        SERVER={server};
        DATABASE={database if database else 'master'};
        Trusted_Connection=yes;
    """
    
    return pyodbc.connect(conn_str)

# --- GET DATABASE LIST ---
# --- This function obtains the list of all databases within the SQL SERVER ---
@st.cache_data
def get_databases():
    conn = get_connection("master")
    query = """
        SELECT name
        FROM sys.databases
        WHERE database_id > 4   -- excludes system DBs
        ORDER BY name
    """
    df = pd.read_sql(query, conn)
    conn.close()
    return df["name"].tolist()

# --- GET TABLE DATA ---
# --- This function obtains the list of runs from the WTW.RUNS Table within the selected Database ---
def get_wtw_runs(database):
    
    conn = get_connection(database)
    db_check = pd.read_sql("SELECT DB_NAME() AS db", conn)
    query = f"SELECT * FROM [{database}].[WTW].[RUNS]"
    df_WTWRuns = pd.read_sql(query, conn)
    conn.close()
    return df_WTWRuns


# 1️⃣ Database dropdown
try:
    databases = get_databases()
    database_default = default_database
    selected_db = st.selectbox("Select Database", databases, index=databases.index(database_default) if database_default in databases else 0)
except Exception as e:
    st.error(f"Error fetching databases: {e}")
    st.stop()


# 1️⃣ Output WTW.RUNS Table and Filter Name to selected runs
df_WTWRuns = get_wtw_runs(selected_db)
# --- Text input ---
search_text_WTW = st.text_input("Search ProjectRunName", value=default_rungroup)

# --- Filter logic ---
# This step allows the user to type a string in the text box to filter the run table list
if search_text_WTW:
    filtered_df = df_WTWRuns[
        df_WTWRuns["Name"].astype(str).str.contains(search_text_WTW, case=False, na=False)
    ]
else:
    filtered_df = df_WTWRuns

# --- Display ---
st.dataframe(filtered_df[["Name"]])
# Dropdown of unique ProjectRunName values to be used as the run selection
# This step uses the filtered dataframe to allow the user to select the run from the filtered list
selected_project = st.selectbox(
    "Select Project Run",
    filtered_df["Name"].dropna().unique()
)

st.write(f"Selected ProjectRunName: {selected_project}")
folder_name = selected_project.split("\\")[-1]
st.session_state.destination_path = os.path.join(base_path, folder_name)













