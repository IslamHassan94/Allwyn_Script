import math
from Script.Config import Config_Setup
from Script.Utils import FilesUtil, DateUtil
from Script.Models.Order import Order
import pandas as pd
from datetime import datetime
from dateutil import parser
from Script.Models.Serials import Serials

Commissioning_File_path = Config_Setup.input_sheets_dir + FilesUtil.get_file_fullName_by_keyword_in_name(
    Config_Setup.Commissioning_File, 'Commissioning')

print(Commissioning_File_path)

def export_serials_from_Commissioning_File():
    # Read the Excel file into a DataFrame
    df = pd.read_excel(Commissioning_File_path)
    df.to_excel("Com.xlsx", index=False)
    # # Initialize an empty list to hold the serials
    serials = []
    #
    # # Loop through the rows and create Serials objects, then append them to the list
    for _, row in df.iterrows():
        serial = Serials(
            serial_Num=row['Router S/N'],
            retailer_id=row['Retailer ID'],
            compledted_Date=row['Completed Router Install Start Date & Time'] if not pd.isna(row['Completed Router Install Start Date & Time']) else None
        )
        serials.append(serial)
        for o in serials:
            print("SSSSSSSSSSSSSSSSSSSSSSS")
            print(o.serial_Num)
            print(o.retailer_id)
            print(o.compledted_Date)
    return serials

def check_duplicated_serials_flat_from_Commissioning_File():
    # Read the Excel file into a DataFrame
    df = pd.read_excel(Commissioning_File_path)

    # Save the file (optional, for debugging purposes)
    df.to_excel("Com.xlsx", index=False)

    serials = []
    serial_counts = {}

    for _, row in df.iterrows():
        serial_num = row['Router S/N']

        serial = Serials(
            serial_Num=serial_num,
            retailer_id=row['Retailer ID'],
            compledted_Date=row['Completed Router Install Start Date & Time'] if not pd.isna(
                row['Completed Router Install Start Date & Time']) else None
        )

        serials.append(serial)

        # Count occurrences of each serial number
        if serial_num in serial_counts:
            serial_counts[serial_num].append(serial)
        else:
            serial_counts[serial_num] = [serial]

    # Filter out only the serials that have duplicates
    duplicated_serials = [serial_list for serial_list in serial_counts.values() if len(serial_list) > 1]

    # Flatten the list of lists (if needed, depending on how you want the output)
    duplicated_serials_flat = [serial for serial_list in duplicated_serials for serial in serial_list]

    for serial in duplicated_serials_flat:
        print("Duplicated Serial:")
        print("Serial Number:", serial.serial_Num)
        print("Retailer ID:", serial.retailer_id)
        print("Completed Date:", serial.compledted_Date)
        # If there are duplicates, export them to an Excel file
        if duplicated_serials_flat:
            # Create a list of dictionaries to convert into a DataFrame
            data = [{
                'Serial Number': serial.serial_Num,
                'Retailer ID': serial.retailer_id,
                'Completed Date': serial.compledted_Date
            } for serial in duplicated_serials_flat]

            # Convert the list of dictionaries to a DataFrame
            df_duplicates = pd.DataFrame(data)

            # Export the DataFrame to an Excel file
            output_path = "duplicated_serials.xlsx"
            df_duplicates.to_excel(output_path, index=False)

            print(f"Duplicated serials saved to {output_path}")
            return True
        else:
            print("No duplicates found.")
            return False


value = check_duplicated_serials_flat_from_Commissioning_File()
print(value)