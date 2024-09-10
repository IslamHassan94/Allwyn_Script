from Script.Config import Config_Setup
from Script.Utils import FilesUtil
from Script.Data import ExportOrders
from openpyxl import load_workbook
import pandas as pd

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
    book = load_workbook(file_path)
    sheet = book[sheet_name]

    # Create a dictionary to map retailer IDs to their row numbers in Excel
    retailer_row_map = {}
    for row in range(6, sheet.max_row + 1):  # Assuming data starts from row 6 in Excel
        retailer_id_cell = sheet[f'A{row}'].value  # Replace 'A' with the column that contains 'Retailer ID'
        if retailer_id_cell:
            retailer_row_map[retailer_id_cell] = row

    # Loop through each order and update the corresponding row in the Excel sheet
    for order in orders:
        if order.retailer_id in retailer_row_map:
            row_num = retailer_row_map[order.retailer_id]

            # Update the corresponding cells in the Excel sheet
            sheet[
                f'X{row_num}'].value = order.date_required  # Replace 'D' with the appropriate column for 'Initial OLO requested date'
            sheet[
                f'Y{row_num}'].value = order.install_date  # Replace 'E' with the column for 'Forecasted OLO install Date'
            sheet[
                f'Z{row_num}'].value = order.first_poll_date  # Replace 'F' with the column for 'Actual completion date/OLO First Poll Date'
            sheet[
                f'AA{row_num}'].value = order.service_activated  # Replace 'G' with the column for 'OLO Service Activated'

    # Save the workbook after making the updates
    book.save(file_path)
    print(f"Successfully updated {len(orders)} orders in the master sheet")


if __name__ == '__main__':
    handle_Csl_to_master()
