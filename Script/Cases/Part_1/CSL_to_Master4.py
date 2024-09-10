from Script.Config import Config_Setup
from Script.Utils import FilesUtil
from Script.Data import ExportOrders
import pandas as pd
import xlwings as xw
import warnings

# Suppress the specific warnings for invalid date values
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

vodafone_provide_path = Config_Setup.input_sheets_dir + FilesUtil.get_file_fullName(Config_Setup.vodafone_provide)
site_status_report_path = Config_Setup.input_sheets_dir + FilesUtil.get_file_fullName(Config_Setup.site_status_report)
allwyn_fault_tracking_path = Config_Setup.input_sheets_dir + FilesUtil.get_file_fullName(
    Config_Setup.allwyn_fault_tracking)
orders = []


def handle_Csl_to_master():
    orders = ExportOrders.export_orders_from_status_report()
    df_master = pd.read_excel(vodafone_provide_path, sheet_name='Provide Update', skiprows=4)
    write_orders_to_master_sheet(orders, df_master, vodafone_provide_path, 'Provide Update')


def write_orders_to_master_sheet(orders, df_master, file_path, sheet_name):
    # Use xlwings to avoid Excel corruption
    app = xw.App(visible=False)
    book = app.books.open(file_path)
    sheet = book.sheets[sheet_name]

    # Create a dictionary to map retailer IDs to their row numbers in Excel
    retailer_row_map = {}
    for row in range(6, sheet.api.UsedRange.Rows.Count + 1):  # Assuming data starts from row 6 in Excel
        retailer_id_cell = sheet.range(f'A{row}').value  # Replace 'A' with the column that contains 'Retailer ID'
        if retailer_id_cell:
            retailer_row_map[retailer_id_cell] = row

    # Handle invalid date cells across all rows
    handle_invalid_dates(sheet)

    # Loop through each order and update the corresponding row in the Excel sheet
    for order in orders:
        if order.retailer_id in retailer_row_map:
            row_num = retailer_row_map[order.retailer_id]

            # Update the corresponding cells in the Excel sheet
            sheet.range(f'X{row_num}').value = order.date_required  # Replace 'X' with the appropriate column
            sheet.range(
                f'Y{row_num}').value = order.install_date  # Replace 'Y' with the column for 'Forecasted OLO install Date'
            sheet.range(
                f'Z{row_num}').value = order.first_poll_date  # Replace 'Z' with the column for 'Actual completion date/OLO First Poll Date'
            if order.service_activated == 1:
                sheet.range(f'AA{row_num}').value = 'TRUE'
                sheet.range(f'U{row_num}').value = '2.Commissioning'
            elif order.service_activated == 0:
                sheet.range(f'AA{row_num}').value = 'FALSE'
            else:
                sheet.range(f'AA{row_num}').value = ''  # Replace 'AA' with the column for 'OLO Service Activated'

    # Save and close the workbook without Excel recovery prompts
    book.save()
    book.close()
    app.quit()
    print(f"Successfully updated {len(orders)} orders in the master sheet")


def handle_invalid_dates(sheet):
    # Iterate through all rows and check for invalid date serial values in column 'T'
    for row in range(6, sheet.api.UsedRange.Rows.Count + 1):  # Adjust start row as needed
        cell_value = sheet.range(f'T{row}').value  # Adjust column 'T' if necessary
        if isinstance(cell_value, int) and cell_value > 2958465:  # Max valid Excel date serial
            print(f"Invalid date value in cell T{row}: {cell_value}")
            # Handle the invalid date by setting it to None
            sheet.range(f'T{row}').value = None


if __name__ == '__main__':
    handle_Csl_to_master()
