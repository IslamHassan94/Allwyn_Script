from Script.Config import Config_Setup
from Script.Utils import FilesUtil
from Script.Data import ExportOrders
from Script.Utils import ProgressAnimation, DateUtil
from Script.Config import Logger
import logging
import pandas as pd
import xlwings as xw
import time
import threading
import sys


#################### Logging #############################
Logger.init_Logger()
logger = logging.getLogger("Allwyn Script")
logging.getLogger("imported_module").setLevel(logging.WARNING)
##################################################################
logger.debug("Getting Sheets names...")
Commissioning_File_path = Config_Setup.input_sheets_dir + FilesUtil.get_file_fullName_by_keyword_in_name(
    Config_Setup.Commissioning_File, 'Commissioning')


print(Commissioning_File_path)

def handle_Commissioning_to_master():
    logger.debug("Exporting orders from Status Report Sheet...")

# def get_Duplicates_From_CSL():
