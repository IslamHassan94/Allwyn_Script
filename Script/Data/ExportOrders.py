import math
from Script.Config import Config_Setup
from Script.Utils import FilesUtil, DateUtil
from Script.Models.Order import Order
import pandas as pd
from dateutil import parser

site_status_report_path = Config_Setup.input_sheets_dir + FilesUtil.get_file_fullName(Config_Setup.site_status_report)

orders = []

from datetime import datetime


def export_orders_from_status_report():
    # Read the Excel file
    df = pd.read_excel(site_status_report_path)
    df = df[df['First Poll Date'].notna()]  # Keep rows where 'First Poll Date' is not NaN
    df = df[
        df['First Poll Date'].str.strip() != '']  # Keep rows where 'First Poll Date' is not an empty string or spaces

    # Convert the relevant date columns to datetime and format them
    df['First Poll Date'] = pd.to_datetime(df['First Poll Date'])
    df['Date Required'] = pd.to_datetime(df['Date Required'], errors='coerce')  # Coerce invalid dates to NaT
    df['Install Date'] = pd.to_datetime(df['Install Date'], errors='coerce')  # Coerce invalid dates to NaT

    # Format the dates to the desired format '%d/%m/%y'
    df['First Poll Date'] = df['First Poll Date'].dt.strftime('%d/%m/%y')
    df['Date Required'] = df['Date Required'].dt.strftime('%d/%m/%y')
    df['Install Date'] = df['Install Date'].dt.strftime('%d/%m/%y')

    # Filter the dataframe to only include the desired columns
    df = df[['Site Reference  ↑', 'Date Required', 'Install Date', 'First Poll Date', 'Service Activated', 'Message',
             'Body']]

    # Remove rows where 'Site Reference ↑' is NaT/NaN/blank
    df = df[df['Site Reference  ↑'].notna()]  # Remove NaT/NaN
    df = df[df['Site Reference  ↑'].str.strip() != '']  # Remove blank or spaces

    # Ensure 'Site Reference  ↑' can be converted to an integer and filter out non-numeric values
    df['Site Reference  ↑'] = pd.to_numeric(df['Site Reference  ↑'], errors='coerce')
    df = df.dropna(subset=['Site Reference  ↑'])  # Remove rows with invalid 'Site Reference  ↑' values
    df['Site Reference  ↑'] = df['Site Reference  ↑'].astype(int)

    # Iterate over the rows of the DataFrame and populate the orders list
    for _, row in df.iterrows():
        first_poll_date = row['First Poll Date']
        date_required = row['Date Required']
        install_date = row['Install Date']

        # Create Order object with the formatted dates
        order = Order(
            retailer_id=row['Site Reference  ↑'],
            date_required=date_required,  # Already formatted
            install_date=install_date,  # Already formatted
            first_poll_date=first_poll_date,  # Already formatted
            service_activated=row['Service Activated'],
            message=row['Message'],
            body=row['Body']
        )
        # Check if the first poll date is the same as yesterday
        if pd.notna(first_poll_date):
            if is_same_day((datetime.strptime(first_poll_date, '%d/%m/%y')), DateUtil.get_yesterday_date()):
                orders.append(order)

    print(f"Successfully added {len(orders)} orders to the list")
    for o in orders:
        print(o.date_required)
    return orders


def is_same_day(date1_input, date2_input):
    print(f'First poll: date [{date1_input}] , yesterday : [{date2_input}]')
    try:
        # Ensure inputs are not NaN or None
        if date1_input is None or date2_input is None:
            print("One or both dates are None.")
            return False
        if isinstance(date1_input, float) and math.isnan(date1_input):
            print("First date is NaN.")
            return False
        if isinstance(date2_input, float) and math.isnan(date2_input):
            print("Second date is NaN.")
            return False

        # Ensure inputs are strings
        date1_str = str(date1_input)
        date2_str = str(date2_input)

        # Parse the dates from strings
        date1 = parser.parse(date1_str)
        date2 = parser.parse(date2_str)

        # Compare the year, month, and day parts
        return date1.date() == date2.date()
    except Exception as e:
        print(f"Error parsing dates: {e}")
        return False
