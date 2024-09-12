from Script.Config import Config_Setup
from Script.Utils import FilesUtil
from Script.Data import ExportOrders
from Script.Utils import ProgressAnimation, DateUtil
from Script.Config import Logger
import logging
import pandas as pd
import xlwings as xw
import time
import threading
import sys

#################### Logging #############################
Logger.init_Logger()
logger = logging.getLogger("Allwyn Script")
logging.getLogger("imported_module").setLevel(logging.WARNING)
##################################################################
logger.debug("Getting Sheets names...")
vodafone_provide_path = Config_Setup.input_sheets_dir + FilesUtil.get_file_fullName_by_keyword_in_name(
    Config_Setup.vodafone_provide, 'Base')
site_status_report_path = Config_Setup.input_sheets_dir + FilesUtil.get_file_fullName(Config_Setup.site_status_report)
allwyn_fault_tracking_path = Config_Setup.input_sheets_dir + FilesUtil.get_file_fullName(
    Config_Setup.allwyn_fault_tracking)
orders = []


def handle_Csl_to_master():
    logger.debug("Exporting orders from Status Report Sheet...")
    orders = ExportOrders.export_orders_from_status_report()
    df_master = pd.read_excel(vodafone_provide_path, sheet_name='Provide Update', skiprows=4)
    stop_event = threading.Event()
    progress_thread = threading.Thread(target=ProgressAnimation.rolling_progress_bar, args=(stop_event,))
    progress_thread.start()
    write_orders_to_master_sheet(orders, df_master, vodafone_provide_path, 'Provide Update')
    stop_event.set()
    progress_thread.join()
    logger.debug("\nData Transfer Completed.")
    generate_final_vodafone_provide_sheet(vodafone_provide_path)


def write_orders_to_master_sheet(orders, df_master, file_path, sheet_name):
    # Use xlwings to avoid Excel corruption
    app = xw.App(visible=False)
    book = app.books.open(file_path)
    sheet = book.sheets[sheet_name]

    # Create a dictionary to map retailer IDs to their row numbers in Excel
    retailer_row_map = {}
    for row in range(6, sheet.api.UsedRange.Rows.Count + 1):  # Assuming data starts from row 6 in Excel
        (retailer_id_cell) = sheet.range(f'A{row}').value  # Replace 'A' with the column that contains 'Retailer ID'
        if retailer_id_cell:
            retailer_row_map[retailer_id_cell] = row

    # Handle invalid date cells across all rows
    handle_invalid_dates(sheet)

    # Loop through each order and update the corresponding row in the Excel sheet
    for order in orders:
        if order.retailer_id in retailer_row_map:
            row_num = retailer_row_map[order.retailer_id]

            # Update the corresponding cells in the Excel sheet
            print(order.date_required)
            sheet.range(f'X{row_num}').value = order.date_required  # Replace 'X' with the appropriate column
            sheet.range(
                f'Y{row_num}').value = order.install_date  # Replace 'Y' with the column for 'Forecasted OLO install Date'
            sheet.range(
                f'Z{row_num}').value = order.first_poll_date  # Replace 'Z' with the column for 'Actual completion date/OLO First Poll Date'

            # Safely handle None values for DateUtil.getTodaysDate()
            today_date = DateUtil.getTodaysDate()

            # Write order body and message only if they are not NaN or empty
            order_body = order.body or ''
            order_message = order.message or ''

            today_date = DateUtil.getTodaysDate()

            sheet.range(f'I{row_num}').value = today_date
            if pd.notna(order_body) and order_body.strip():
                sheet.range(f'M{row_num}').value = order_body
            if pd.notna(order_message) and order_message.strip():
                sheet.range(f'L{row_num}').value = order_message

            # Update the concatenated messages in columns 'K' and 'J' only if body or message are not empty
            if pd.notna(order_body) and pd.notna(order_message) and (order_body.strip() or order_message.strip()):
                if pd.notna(order_body) and order_body.strip() and not pd.notna(order_message):
                    concatenated_message = f"{today_date}: {order_body}"
                    sheet.range(f'K{row_num}').value = concatenated_message
                elif pd.notna(order_message) and order_message.strip() and not pd.notna(order_body):
                    concatenated_message = f"{today_date}: {order_message}"
                    sheet.range(f'K{row_num}').value = concatenated_message
                elif pd.notna(order_message) and order_message.strip() and pd.notna(order_body) and order_body.strip():
                    concatenated_message = f"{today_date}: {order_body} {order_message}"
                    sheet.range(f'K{row_num}').value = concatenated_message

            sheet.range(f'J{row_num}').value = sheet.range(f'K{row_num}').value
            if order.service_activated == 1 or order.service_activated == 'TRUE' or order.service_activated == 'True':
                sheet.range(f'AA{row_num}').value = 'TRUE'
                sheet.range(f'U{row_num}').value = '2.Commissioning'
            elif order.service_activated == 0 or order.service_activated == 'FALSE' or order.service_activated == 'False':
                sheet.range(f'AA{row_num}').value = ''
            else:
                sheet.range(f'AA{row_num}').value = ''  # Replace 'AA' with the column for 'OLO Service Activated'

    # Save and close the workbook without Excel recovery prompts
    book.save()
    book.close()
    app.quit()


def handle_invalid_dates(sheet):
    # Iterate through all rows and check for invalid date serial values in column 'T'
    for row in range(6, sheet.api.UsedRange.Rows.Count + 1):  # Adjust start row as needed
        cell_value = sheet.range(f'T{row}').value  # Adjust column 'T' if necessary
        if isinstance(cell_value, int) and cell_value > 2958465:  # Max valid Excel date serial
            print(f"Invalid date value in cell T{row}: {cell_value}")
            # Handle the invalid date by setting it to None
            sheet.range(f'T{row}').value = None


def generate_final_vodafone_provide_sheet(path):
    # Define the required columns
    required_columns = [
        'Retailer ID', 'SR No.', 'Order batch date', 'Allwyn Site Type (ie Type 1 or 2)',
        'SOGEA / FTTP', 'Store Name', 'City', 'Postcode', 'Updates / Comments',
        'Access Service Id (VF Access Service Id)', 'Connection ID (CSL Service)', 'Site Status',
        'Appointment Slot - AM (9am-1pm) / PM (1pm - 5pm)', 'Site Survey Date', 'Initial OLO requested date',
        'Forecasted OLO install Date', 'Actual completion date OLO First Poll Date',
        'OLO Service Activated', 'Line test (Fault)', 'Scheduled Router Install Date',
        'Completed Router Install Date & Time', 'CSL Router - S/N'
    ]

    # Read the Excel sheet
    try:
        df = pd.read_excel(path, sheet_name='Provide Update', skiprows=4)
        print("Columns in Excel file:", df.columns.tolist())  # Print the actual columns for inspection
    except Exception as e:
        print(f"Error reading the file: {e}")
        return None

    # Replace newline characters in column names for easier handling
    df.columns = df.columns.str.replace('\n', ' ').str.strip()

    # Check if the required columns are present in the DataFrame
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise KeyError(f"The following columns are missing from the DataFrame: {missing_columns}")

    # Filter the DataFrame to include only the required columns
    filtered_df = df.loc[0:, required_columns]

    # Define columns that potentially contain dates (excluding 'Order batch date')
    date_columns = [
        'Site Survey Date', 'Initial OLO requested date', 'Forecasted OLO install Date',
        'Actual completion date OLO First Poll Date',  # Adjusted name without '\n'
        'Scheduled Router Install Date', 'Completed Router Install Date & Time'
    ]

    # Format the date columns in UK format (DD/MM/YYYY) and remove time component
    for col in date_columns:
        if col in filtered_df.columns:
            filtered_df[col] = pd.to_datetime(filtered_df[col], errors='coerce').dt.date
            filtered_df[col] = filtered_df[col].apply(lambda x: x.strftime('%d/%m/%Y') if pd.notna(x) else '')

    # Save the filtered DataFrame to a new Excel file
    output_file_path = Config_Setup.output_folder + f'Vodafone_Provide_Update_{DateUtil.getTodaysDateInSerialFormat()}.xlsx'
    try:
        filtered_df.to_excel(output_file_path, index=False)
    except Exception as e:
        print(f"Error saving the file: {e}")

    # Save to Excel with openpyxl in case further appending or modifications are needed
    with pd.ExcelWriter(output_file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        filtered_df.to_excel(writer, startrow=0, index=False)
