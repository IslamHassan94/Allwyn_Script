# config_setup.py

import yaml

# Load configuration from YAML file
config = yaml.safe_load(open("../../config.yml"))

# Extract values from the configuration
input_sheets_dir = config['Input_Sheets']['input_folder']
output_folder = config['output_folder']
vodafone_provide = config['Input_Sheets']['vodafone_provide']
allwyn_fault_tracking = config['Input_Sheets']['allwyn_fault_tracking']
site_status_report = config['Input_Sheets']['site_status_report']
Commissioning_File = config['Input_Sheets']['Commissioning_File']
password = config['password']
password_protection_path = config['password_protection_path']
