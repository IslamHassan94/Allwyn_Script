from Script.Cases.Part_1 import CSL_to_Master
from Script.Utils import FilesUtil, DateUtil
from Script.Config import Logger, Config_Setup
import logging
import os

#################### Logging #############################
Logger.init_Logger()
logger = logging.getLogger("Allwyn Script")
logging.getLogger("imported_module").setLevel(logging.WARNING)
#################################################################

output_file_path = os.path.abspath(
    Config_Setup.password_protection_path) + '\\' + f'Vodafone_Provide_Update_{DateUtil.getTodaysDateInSerialFormat()}.xlsx'

if __name__ == '__main__':
    CSL_to_Master.handle_Csl_to_master()
    CSL_to_Master.generate_final_vodafone_provide_sheet(CSL_to_Master.vodafone_provide_path)
