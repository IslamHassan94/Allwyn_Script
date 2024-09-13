import pandas as pd
from Script.Models.Serials import Serials
from Script.Config import Config_Setup
from Script.Utils import FilesUtil

# Get the commissioning file path
Commissioning_File_path = Config_Setup.input_sheets_dir + FilesUtil.get_file_fullName_by_keyword_in_name(
    Config_Setup.Commissioning_File, 'Commissioning')


def export_serials_from_Commissioning_File():
    # Read the Excel file into a DataFrame
    df = pd.read_excel(Commissioning_File_path)

    # Save the file (optional, for debugging purposes)
    df.to_excel("Com.xlsx", index=False)

    # Initialize an empty list to hold the serials and a dictionary to count occurrences
    serials = []
    serial_counts = {}

    # Loop through the rows and create Serials objects, then append them to the list
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

    # Print the duplicated serials
    for serial in duplicated_serials_flat:
        print("Duplicated Serial:")
        print("Serial Number:", serial.serial_Num)
        print("Retailer ID:", serial.retailer_id)
        print("Completed Date:", serial.compledted_Date)

    return duplicated_serials_flat


# Example usage
export_serials_from_Commissioning_File()

