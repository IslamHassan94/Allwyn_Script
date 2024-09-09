# config_setup.py

import yaml

# Load configuration from YAML file
config = yaml.safe_load(open("../../config.yml"))

# Extract values from the configuration
export_data = config['steps']['export_data']
export_GDP = config['steps']['export_GDP']
export_Terminalis = config['steps']['export_Terminalis']
export_ErrorProv = config['steps']['export_ErrorProv']
export_Renovaction = config['steps']['export_Renovaction']