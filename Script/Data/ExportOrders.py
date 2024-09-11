from Script.Config import Config_Setup
from Script.Utils import FilesUtil
from Script.Models.Order import Order
import pandas as pd

site_status_report_path = Config_Setup.input_sheets_dir + FilesUtil.get_file_fullName(Config_Setup.site_status_report)

orders = []


def export_orders_from_status_report():
    # Read the Excel file
    df = pd.read_excel(site_status_report_path)

    # Filter the dataframe to only include the desired columns
    df = df[['Site Reference  ↑', 'Date Required', 'Install Date', 'First Poll Date', 'Service Activated', 'Message',
             'Body']]

    # Remove rows where 'Site Reference ↑' is NaT/NaN/blank
    df = df[df['Site Reference  ↑'].notna()]  # Remove NaT/NaN
    df = df[df['Site Reference  ↑'].str.strip() != '']  # Remove blank or spaces

    # Iterate over the rows of the DataFrame and populate the orders list
    for _, row in df.iterrows():
        order = Order(
            retailer_id=row['Site Reference  ↑'],
            date_required=row['Date Required'],
            install_date=row['Install Date'],
            first_poll_date=row['First Poll Date'],
            service_activated=row['Service Activated'],
            message=row['Message'],
            body=row['Body']
        )
        orders.append(order)

    print(f"Successfully added {len(orders)} orders to the list")
    return orders
