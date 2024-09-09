from Script.Config import Config_Setup
from Script.Utils import FilesUtil
from Script.Data import ExportOrders
import openpyxl
import pandas as pd

vodafone_provide_path = Config_Setup.input_sheets_dir + FilesUtil.get_file_fullName(Config_Setup.vodafone_provide)
site_status_report_path = Config_Setup.input_sheets_dir + FilesUtil.get_file_fullName(Config_Setup.site_status_report)
allwyn_fault_tracking_path = Config_Setup.input_sheets_dir + FilesUtil.get_file_fullName(
    Config_Setup.allwyn_fault_tracking)
orders = []


def handle_Csl_to_master():
    orders = ExportOrders.export_orders_from_status_report()
    df_master = pd.read_excel(vodafone_provide_path, sheet_name='Provide Update', skiprows=4)
    write_orders_to_master_sheet(orders, df_master)


def write_orders_to_master_sheet(orders, df_master):
    # Loop through each order in the orders list
    for order in orders:
        # Find the row in the master DataFrame where the retailer ID matches
        df_row = df_master[df_master['Retailer ID'] == order.retailer_id]

        # If a matching row is found, update the corresponding cells
        if not df_row.empty:
            idx = df_row.index[0]  # Get the index of the matching row
            df_master.at[idx, 'Initial OLO requested date'] = order.date_required
            df_master.at[idx, 'Forecasted OLO install Date'] = order.install_date
            df_master.at[idx, 'Actual completion date\nOLO First Poll Date'] = order.first_poll_date
            df_master.at[idx, 'OLO Service Activated'] = order.service_activated

    df_master.to_excel('updated_provide_sheet.xlsx')
    # # Save or return the updated DataFrame if needed
    # with pd.ExcelWriter(vodafone_provide_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    #     df_master.to_excel(writer, sheet_name='Provide Update', startrow=0, index=False)

    print(f"Successfully updated {len(orders)} orders in the master sheet")


if __name__ == '__main__':
    handle_Csl_to_master()
