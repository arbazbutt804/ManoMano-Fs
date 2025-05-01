import logging
import pandas as pd
import requests
import streamlit as st

from io import StringIO, BytesIO
from dotenv import load_dotenv

# Set up basic logging configuration
logging.basicConfig(
    level=logging.DEBUG,  # Set logging level to DEBUG, INFO, WARNING, ERROR
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler(),  # Logs to the terminal (good for Streamlit Cloud)
        #logging.FileHandler("app.log")  # Logs to a file (optional)
    ]
)
ASANA_TOKEN = st.secrets("ASANA_TOKEN")

# Initialize session state for keeping track of file paths
if "output_file" not in st.session_state:
    st.session_state.output_file = None

def analyze_idq(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file)
        # Filter for products with review scores above 0.1 but below 3.5
        filtered_df = df[(df['PRODUCT_RATING'] > 0.1) & (df['PRODUCT_RATING'] < 3.5)]
        grouped = filtered_df.groupby('PLATFORM')
        # Create a BytesIO object to save the Excel file
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for name, group in grouped:
                group[['EAN','SKU']].to_excel(writer, sheet_name=name, index=False)
        output.seek(0)
        st.session_state.output_file = output
    except Exception as e:
        logging.error(f"An unexpected error occurred during the initial IDQ analysis: {e}")
        st.error("An unexpected error occurred during the initial IDQ analysis")

def update_excel_with_sku_description():
    try:
        logging.info("Starting to update F1s.xlsx with SKU description.")
        # Load the Excel file from session state
        input_file = st.session_state.output_file
        csv_file = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vS_mN7-KwnH2aN-afhBMbM_1IlBylxwgJByEkQU5M3HJQuSDx8-pk3HwaJ5TOLgNeD0SGcdgHikloFK/pub?gid=788370787&single=true&output=csv'

        # Read the CSV file into a DataFrame
        df_csv = pd.read_csv(csv_file, header=2)

        # Open the Excel file for reading sheet names
        xls = pd.ExcelFile(input_file)
        sheet_names = xls.sheet_names
        logging.info(f"Found sheet names: {sheet_names}")

        # Store dataframes temporarily
        df_dict = {}

        # Read and process each sheet, then store in df_dict
        for sheet in sheet_names:
            logging.info(f"Processing sheet: {sheet}")

            # Read the Excel sheet into a DataFrame
            df_excel = pd.read_excel(input_file, sheet_name=sheet)

            # Log the column names for debugging
            logging.info(f"Columns in {sheet}: {df_excel.columns.tolist()}")

            # Check if 'Seller SKU' exists in df_excel
            if 'SKU' in df_excel.columns:
                df_excel['SKU'] = df_excel['SKU'].astype(str).str.strip()
                df_csv['Sku code'] = df_csv['Sku code'].astype(str).str.strip()
                # Create a lookup column without the F1, F2, F3, etc. suffix
                df_excel['SKU Lookup'] = df_excel['SKU'].str.extract(r'(\d+)')

                # Merge the Excel DataFrame and the CSV DataFrame based on 'SKU Lookup' and 'Sku code'
                logging.info(f"Merging SKU description for sheet {sheet}.")
                merged_df = pd.merge(df_excel, df_csv[['Sku code', 'Sku description']], left_on='SKU Lookup',
                                     right_on='Sku code', how='left')

                # Drop the 'Sku code' and 'SKU Lookup' columns as they're redundant
                merged_df.drop(columns=['Sku code', 'SKU Lookup'], inplace=True)
            else:
                logging.warning(f"'SKU' column not found in {sheet}. Skipping SKU description merge.")
                merged_df = df_excel

            df_dict[sheet] = merged_df
        # Close the read operation
        del xls

        # Create a BytesIO object to store the CSV data
        output = BytesIO()
        with pd.ExcelWriter(output) as writer:
            for sheet, df in df_dict.items():
                logging.info(f"Writing updated data to sheet {sheet}.")
                df.to_excel(writer, sheet_name=sheet, index=False)
        output.seek(0)
        # Store the updated file in session state
        st.session_state.output_file = output
        logging.info("Successfully updated F1s.xlsx with SKU description information. Saved as F1s - Desc Added.xlsx.")

    except Exception as e:
        logging.error(f"An error occurred while updating the Excel file with SKU description: {e}")
        st.error("An error occurred while updating the Excel file with SKU description")


def update_excel_with_f1_to_use():
    try:
        logging.info("Starting to update F1s - Desc Added.xlsx with F1 to Use.")

        # Load the existing Excel file from session state
        input_file = st.session_state.output_file

        # Fetch the CSV file from the URL
        url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRxBqpSTMwezeOji3KXDlrp3855sQHFuYxmKsCIDwILg4iHMEx2BBmp87nwEgI__4g3rM6H65rIp0sF/pub?gid=0&single=true&output=csv"
        response = requests.get(url)
        csv_data = StringIO(response.text)
        df_csv = pd.read_csv(csv_data)

        xls = pd.ExcelFile(input_file)
        sheet_names = xls.sheet_names
        logging.info(f"Found sheet names: {sheet_names}")

        # Store dataframes temporarily
        df_dict = {}

        # Read and process each sheet, then store in df_dict
        for sheet in sheet_names:
            logging.info(f"Processing sheet: {sheet}")

            # Read the Excel sheet into a DataFrame
            df_excel = pd.read_excel(input_file, sheet_name=sheet)

            # Check if 'Seller SKU' exists in df_excel
            if 'SKU' in df_excel.columns:
                # Initialize an empty list to hold F1 to Use values
                f1_to_use_values = []

                for sku in df_excel['SKU']:
                    # Search for the SKU in columns B to P of the CSV DataFrame
                    found_row = df_csv.iloc[:, 1:16][
                        df_csv.iloc[:, 1:16].apply(lambda row: row.astype(str).str.contains(str(sku), na=False).any(),
                                                   axis=1)]  # Search for SKU in columns B to P

                    if not found_row.empty:
                        # Take the last non-empty value from the row
                        last_non_empty_value = found_row.iloc[0, :].dropna().iloc[-1]
                        if last_non_empty_value == sku:
                            f1_to_use_values.append(None)
                        else:
                            f1_to_use_values.append(last_non_empty_value)
                    else:
                        f1_to_use_values.append(None)

                # Add the F1 to Use column to the DataFrame
                df_excel['F1 to Use'] = f1_to_use_values
                df_dict[sheet] = df_excel
            else:
                logging.warning(f"'SKU' column not found in sheet {sheet}. Skipping this sheet.")

        # Close the read operation
        del xls

        # Write the updated data back to a BytesIO object
        output = BytesIO()
        with pd.ExcelWriter(output) as writer:
            for sheet, df in df_dict.items():
                logging.info(f"Writing updated data to sheet {sheet}.")
                df.to_excel(writer, sheet_name=sheet, index=False)
        output.seek(0)  # Reset the pointer of the BytesIO object
        st.session_state.output_file = output
        logging.info(
            "Successfully updated F1s - Desc Added.xlsx with F1 to Use information. Saved as F1s - Desc Added with F1 to Use.xlsx.")
    except Exception as e:
        logging.error(f"An error occurred while updating the Excel file with F1 to Use: {e}")


def update_excel_with_barcodes(uploaded_barcodes):
    try:
        logging.info("Updating filtered_ratings_with_desc_and_F1_to_use.xlsx with Barcodes.")

        input_file = st.session_state.output_file
        df_barcodes = pd.read_csv(uploaded_barcodes, header=3)

        xls = pd.ExcelFile(input_file)
        sheet_names = xls.sheet_names

        df_dict = {}
        for sheet in sheet_names:
            logging.info(f"Processing sheet: {sheet}")
            df_excel = pd.read_excel(input_file, sheet_name=sheet)

            if 'F1 to Use' in df_excel.columns:
                barcode_values = []
                gs1_brand_values = []

                for f1 in df_excel['F1 to Use']:
                    found_row = df_barcodes[df_barcodes['SKU'] == f1]
                    if not found_row.empty:
                        number_value = str(found_row['Number'].iloc[0]).replace('=', '').replace('"', '')
                        barcode_values.append(number_value)

                        gs1_brand_value = found_row['Main Brand'].iloc[0]
                        gs1_brand_values.append(gs1_brand_value)
                    else:
                        barcode_values.append(None)
                        gs1_brand_values.append(None)

                df_excel['Barcode'] = barcode_values
                df_excel['GS1 Brand'] = gs1_brand_values
                df_dict[sheet] = df_excel
            else:
                logging.warning(f"'F1 to Use' column not found in sheet {sheet}. Skipping this sheet.")

        del xls

        # Write the updated data back to a BytesIO object
        output = BytesIO()
        # Open a new Excel writer and write data
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for sheet, df in df_dict.items():
                logging.info(f"Writing updated data to sheet {sheet}.")
                df.to_excel(writer, sheet_name=sheet, index=False)

        logging.info(f"Successfully updated {output} with Barcodes.")
        # Store the output file path in session state so it can be downloaded later
        output.seek(0)  # Reset the pointer of the BytesIO object
        st.session_state.output_file = output

    except Exception as e:
        logging.error(f"An error occurred while updating the Excel file with Barcodes: {e}")
        st.error("An error occurred while updating the Excel file with Barcodes")

def create_asana_tasks_from_excel(send_to_asana=True):
    print("create_asana_tasks_from_excel")
    if not send_to_asana:
        st.info("Task creation in Asana is disabled.")
        return

    # Asana API setup
    url = "https://app.asana.com/api/1.0/tasks?opt_fields="
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "authorization": f"Bearer {ASANA_TOKEN}"
    }

    # Load the updated F1s Excel file
    input_file = st.session_state.output_file
    # Save the DataFrame to an Excel file in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name in pd.ExcelFile(input_file).sheet_names:
            df = pd.read_excel(input_file, sheet_name=sheet_name)
            sheet_data = []

            # Check if 'EAN' column exists in the DataFrame
            if 'EAN' not in df.columns:
                print("The 'EAN' column is missing in the Excel sheet.")
                continue  # Skip processing this sheet if 'EAN' is missing

            for idx, row in df.iterrows():
                new_f1_barcode = row['Barcode']
                # Remove any leading apostrophes if the EAN is a string
                if isinstance(new_f1_barcode, str):
                    new_f1_barcode = new_f1_barcode.lstrip("'")
                # Convert float EAN values to integer and then to string, but only if it's not NaN
                elif isinstance(new_f1_barcode, float) and not pd.isna(new_f1_barcode):
                    new_f1_barcode = str(int(new_f1_barcode))

                if pd.notna(new_f1_barcode) and (isinstance(new_f1_barcode, str) or isinstance(new_f1_barcode, int)):
                    new_f1_barcode = str(new_f1_barcode)

                    # Value is valid, proceed with task creation
                    task_name = f"F1 for {row['SKU']} - {row['Sku description']}"
                    sku_to_f1 = row['SKU']
                    new_f1_sku = row['F1 to Use']
                    existing_f1_ean = row['EAN']
                    new_f1_brand = row['GS1 Brand']
                    sheet_data.append([task_name,sku_to_f1, new_f1_sku, existing_f1_ean,new_f1_barcode, new_f1_brand])
                else:
                    if not pd.notna(row['F1 to Use']):
                        print(
                            f"EAN '{new_f1_barcode}' (data type: {type(new_f1_barcode)}) is not a valid value for SKU {row['SKU']}. Skipping Asana task creation.")
                        if row['SKU'] not in unique_seller_skus:
                            unique_seller_skus.add(row['SKU'])  # Add to the unique set

                            # Add to the list of tasks needing new EANs
                            new_eans_needed.append({
                                'Seller SKU': row['SKU'],
                                'Sku description': row['Sku description']
                            })
            if sheet_data:
                # Create a DataFrame for the Excel file
                df_skus = pd.DataFrame(sheet_data, columns=['Task','SKU to be F1', 'New F1 SKU', 'Existing F1 EAN','New F1 Barcode', 'New F1 Brand'])
                df_skus.to_excel(writer, index=False, sheet_name=sheet_name)
    # After all sheets are written, send the file to Asana
    project_section_map = {
        '1205420991974313': '1210132854403371',  # UK
        '1205436216136678': '1210125800761924',  # ES
        '1205436216136683': '1210133451103514',  # IT
        '1205436216136660': '1210133451103517',  # FR
        '1205436216136667': '1210133451103520'  # DE
    }
    projects = list(project_section_map.keys())
    tags = ['1203197857163437']
    notes_content = (f"<body><b>File attached in this task </b> \n"
                             "\n"
                             "<b>PLEASE TICK EACH ITEM ON YOUR CHECKLIST AS YOU GO</b></body>")
    for project_id in projects:
        section_gid = project_section_map[project_id]
        payload = {
            "data": {
                "projects": [project_id],
                "name": "ManoMano F1s to be completed",
                "html_notes": notes_content,
                "tags": tags  # Use the looked-up tag ID here
            }
        }
        # Create the task on Asana
        response = requests.post(url, json=payload, headers=headers)
        task_data = response.json()
        if 'data' in task_data and 'gid' in task_data['data']:
            task_gid = task_data['data']['gid']
            # Move task to ManoMano section
            #UK ES IT FR DE
            move_url = f"https://app.asana.com/api/1.0/sections/{section_gid}/addTask"
            move_payload = {"data": {"task": task_gid}}
            requests.post(move_url, json=move_payload, headers=headers)

            output.seek(0)
            # Upload the CSV file as an attachment to the task
            headers = {
                "Authorization": f"Bearer {ASANA_TOKEN}"
            }
            upload_url = f"https://app.asana.com/api/1.0/tasks/{task_gid}/attachments"
            files = {'file': (
            'manomano_F1_sku_details.xlsx', output, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
            attach_response = requests.post(upload_url, headers=headers, files=files)

            if attach_response.status_code == 200:
                logging.info(f"Excel file successfully attached to task {task_gid}.")
            else:
                logging.error(f"Failed to upload the Excel file. Response: {attach_response.json()}")

        # if new_eans_needed:
        #     # Create the main task
        #     main_task_payload = {
        #         "data": {
        #             "name": "NEW F1's Needed",
        #             "assignee": "1208716819375873",
        #             "html_notes": "<body><b>Please can the following new F1's be created and added to the F1 Log <a href=\"https://docs.google.com/spreadsheets/d/1JesoDfHewylxsso0luFrY6KDclv3kvNjugnvMjRH2ak/edit#gid=0\" target=\"_blank\">here</a></b></body>",
        #             "followers": ["greg.stephenson@monstergroupuk.co.uk, 1208716819375873,1208388789142367"],
        #             "workspace": "17406368418784"
        #         }
        #     }
        #     main_task_response = requests.post(url, json=main_task_payload, headers=headers)
        #     main_task_data = main_task_response.json()
        #     main_task_gid = main_task_data['data']['gid']
        #
        #     # Create subtasks
        #     subtask_url = f"https://app.asana.com/api/1.0/tasks/{main_task_gid}/subtasks"
        #     for task in new_eans_needed:
        #         subtask_name = f"{task['Seller SKU']} - {task['Sku description']}"
        #         subtask_payload = {
        #             "data": {
        #                 "name": subtask_name
        #             }
        #         }
        #         subtask_response = requests.post(subtask_url, json=subtask_payload, headers=headers)
        #         print(f"Added subtask: {subtask_name}. Response: {subtask_response.json()}")

# Initialize an empty set to store unique seller-skus
unique_seller_skus = set()
# Initialize an empty list to store tasks with missing EANs
new_eans_needed = []
# Prepare the list to store SKU details for CSV
all_skus_data = []

# Country-to-project ID mapping
country_project_map = {
    'UK': '1205420991974313',
    'FR': '1205436216136660',
    'DE': '1205436216136667',
    'IT': '1205436216136683',
    'ES': '1205436216136678'
    # Add other countries here
}
country_section_map = {
    'UK': '1205420991974320',
    'FR': '1205436216136664',
    'DE': '1205436216136669',
    'IT': '1205436216136685',
    'ES': '1205436216136680'
}

# Country-to-tag ID mapping
country_tag_map = {
    'UK': '1205430582965096',
    'FR': '1205436216136698',
    'DE': '1205436216136699',
    'IT': '1205436216136700',
    'ES': '1205436216136701'
    # Add other countries and their tag IDs here
}

def main():
    st.set_page_config(page_title="IDQ File Processor", page_icon="📄")

    st.markdown(
        """
        <h1 style='text-align: center;'>
            🔄 ManoMano F1s
        </h1>
        """,
        unsafe_allow_html=True
    )

    st.markdown("""<style>
        .css-1offfwp {padding-top: 1rem;}
        .css-1v3fvcr {background-color: #f8f9fa !important;}
        .block-container {padding: 10rem !important;}
        .stButton button, .stDownloadButton button {background-color: #4CAF50; color: white; border: none; border-radius: 5px; padding: 10px 20px; font-size: 16px; cursor: pointer;}
        .stButton button:hover, .stDownloadButton button:hover {background-color: #45a049;}
        .stFileUploader {border: 2px dashed #4CAF50 !important; border-radius: 10px;}
        </style>""", unsafe_allow_html=True)
    # File uploader widget for the user to upload their IDQ file
    uploaded_file = st.file_uploader("Upload IDQ Excel file", type="xlsx")
    # File uploader widget for the user to upload their barcodes file
    uploaded_barcodes = st.file_uploader("Upload Barcode CSV file", type="csv")

    if uploaded_file is not None and uploaded_barcodes is not None and st.session_state.output_file is None:
        # When a file is uploaded, run the analysis
        with st.spinner("Processing your files. This may take a few moments..."):
                analyze_idq(uploaded_file)
                update_excel_with_sku_description()
                update_excel_with_f1_to_use()
                update_excel_with_barcodes(uploaded_barcodes)
    # Check if the output file exists and show download button
    if st.session_state.output_file is not None:
        # Use Streamlit columns to place buttons side-by-side
        col1, col2, col3 = st.columns([0.1, 1, 1])
        # Column 1: Download Button
        with col2:
            st.download_button(label="Save File", data=st.session_state.output_file, file_name="F1_Barcodes.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Column 2: Trigger Asana Functionality
        with col3:
            if st.button("Create Asana Tasks"):
                st.info("Starting Asana task creation...")
                create_asana_tasks_from_excel(send_to_asana=True)  # Call your function here
                st.success("Asana tasks created successfully!")

if __name__ == "__main__":
    main()