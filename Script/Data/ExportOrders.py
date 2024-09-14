import math
from datetime import timedelta

import pandas as pd
from dateutil import parser

from Script.Config import Config_Setup
from Script.Models.Order import Order
from Script.Utils import FilesUtil

site_status_report_path = Config_Setup.input_sheets_dir + FilesUtil.get_file_fullName(Config_Setup.site_status_report)

orders = []

from datetime import datetime


def export_orders_from_status_report():
    # Read the Excel file
    df = pd.read_excel(site_status_report_path)

    # Keep rows where 'First Poll Date' is not NaN or empty
    df = df[df['First Poll Date'].notna() & df['First Poll Date'].str.strip() != '']

    # Strip leading/trailing spaces in date columns (in case there are hidden characters)
    df['First Poll Date'] = df['First Poll Date'].str.strip()
    df['Date Required'] = df['Date Required'].astype(str).str.strip()  # Convert to str and strip spaces
    df['Install Date'] = df['Install Date'].astype(str).str.strip()

    # Convert the relevant date columns to datetime and handle parsing errors with dayfirst=True
    df['First Poll Date'] = pd.to_datetime(df['First Poll Date'], errors='coerce', dayfirst=True)
    df['Date Required'] = pd.to_datetime(df['Date Required'], errors='coerce', dayfirst=True)
    df['Install Date'] = pd.to_datetime(df['Install Date'], errors='coerce', dayfirst=True)

    # Custom formatting rule: if 'First Poll Date' is '09/12/2024', change it to '12/09/2024'
    df['First Poll Date'] = df['First Poll Date'].apply(
        lambda x: x.strftime('%d/%m/%Y') if pd.notna(x) else ''
    )

    # Specific rule for changing '09/12/2024' to '12/09/2024'
    df['First Poll Date'] = df['First Poll Date'].replace('09/12/2024', '12/09/2024')

    # Check for any NaN in 'Date Required' and print details for debugging
    problematic_rows = df[df['Date Required'].isna()]
    if not problematic_rows.empty:
        print("Found NaN in 'Date Required' after conversion. Here are the raw values:")
        print(problematic_rows[['Site Reference  ↑', 'Date Required']])

    # Filter the dataframe to only include the desired columns
    df = df[['Site Reference  ↑', 'Date Required', 'Install Date', 'First Poll Date', 'Service Activated', 'Message',
             'Body']]

    # Remove rows where 'Site Reference ↑' is NaN or blank
    df = df[df['Site Reference  ↑'].notna() & df['Site Reference  ↑'].str.strip() != '']

    # Ensure 'Site Reference  ↑' can be converted to an integer and filter out non-numeric values
    df['Site Reference  ↑'] = pd.to_numeric(df['Site Reference  ↑'], errors='coerce')
    df = df.dropna(subset=['Site Reference  ↑'])
    df['Site Reference  ↑'] = df['Site Reference  ↑'].astype(int)

    # Prepare the list of orders
    orders = []

    # Iterate over the rows of the DataFrame and populate the orders list
    for _, row in df.iterrows():
        first_poll_date = row['First Poll Date']
        date_required = row['Date Required']
        install_date = row['Install Date']

        # Create Order object with the formatted first_poll_date
        order = Order(
            retailer_id=row['Site Reference  ↑'],
            date_required=date_required,
            install_date=install_date,
            first_poll_date=first_poll_date,  # Store the formatted first_poll_date
            service_activated=row['Service Activated'],
            message=row['Message'],
            body=row['Body']
        )

        # Add the order to the list
        orders.append(order)

    print(f"Successfully added {len(orders)} orders to the list")

    for o in orders:
        print(o.first_poll_date)  # Output the formatted first_poll_date for each order

    return orders


def is_same_day(date1_input, date2_input):
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

        # Parse the dates from strings, assuming day first (dd/mm/yyyy)
        date1 = parser.parse(date1_str, dayfirst=True)
        date2 = parser.parse(date2_str, dayfirst=True)

        # Format both dates as '%d/%m/%Y' for logging
        date1_formatted = date1.strftime('%d/%m/%Y')
        date2_formatted = date2.strftime('%d/%m/%Y')

        # Print the comparison log with formatted dates
        print(f'First poll: date [{date1_formatted}] , yesterday : [{date2_formatted}]')

        # Compare the year, month, and day parts
        return date1.date() == date2.date()

    except Exception as e:
        print(f"Error parsing dates: {e}")
        return False


def filter_True_orders(orders):
    # Get yesterday's date
    yesterday = datetime.now() - timedelta(days=2)

    # Format yesterday's date in 'dd/mm/yyyy' format to ensure consistency
    yesterday_formatted = yesterday.strftime('%d/%m/%Y')

    # Filter the orders list based on the service_activated and first_poll_date criteria
    filtered_orders = [
        order for order in orders
        if ((order.service_activated == 'TRUE' or order.service_activated == 1 or order.service_activated is True)
            and is_same_day(order.first_poll_date, yesterday_formatted))
    ]
    return filtered_orders


def filter_orders_without_true(orders, orders_true):
    # Use list comprehension to filter orders
    filtered_orders = [
        order for order in orders
        if order.retailer_id not in {order_true.retailer_id for order_true in orders_true}
    ]
    return filtered_orders
