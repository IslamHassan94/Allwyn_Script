import os
import sys

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))

from Script.Config import Logger
import logging
import yaml
from Script.Utils import FilesUtil
from Script.Data import ExportData
from Script.Cases import Dispatcher
from Script.Pages import ClarifyPage
from Script.Models.CaseType import CaseType
from Script.Config.Config_Setup import (
    export_data,
    dispatching_sheet,
    backup_path,
    Informe_Sheet_path,
    tasks_per_agent_gdp,
    tasks_per_agent_terminales,
    tasks_per_agent_renov,
    tasks_per_agent_error_prov,
    gdp_sheet_name,
    terminalis_sheet_name,
    renovatcion_sheet_name,
    errorProv_sheet_name,
    clear_sheets
)

################ Variables #########################
global agents_list, gdp_tasks, term_tasks, renov_tasks, errorProv_tasks, errorProv_agents, errorProv_tasks_pdte, errorProv_tasks_com
agents_list = []
agents_list = []
gdp_tasks = []
term_tasks = []
renov_tasks = []
errorProv_tasks = []
errorProv_agents = []
#################### Logging #############################

Logger.init_Logger()
logger = logging.getLogger("ES_Quality_Dispatching_Scrupt")
logging.getLogger("imported_module").setLevel(logging.WARNING)
################################################################################


if __name__ == '__main__':
    logger.debug("Release : Date 28-08-2024 -- Date : 2:11 PM... ")
    logger.debug('Loading all Agents...')
    agents_list = ExportData.getAgents()

    logger.debug('Get current gdp agents...')
    gdp_agents = [agent for agent in agents_list if agent.is_handling_gdp]

    logger.debug('Get current terminalis agents...')
    terminalis_agents = [agent for agent in agents_list if agent.is_handling_termenalis]

    logger.debug('Get current renovaction agents...')
    renovaction_agents = [agent for agent in agents_list if agent.is_handling_renovaction]

    logger.debug('Get current error provision agents...')
    errorProv_agents = [agent for agent in agents_list if agent.is_handling_error_prov]
    logger.debug('................................................')

    if str(export_data) == 'true':
        if str(clear_sheets) == 'true':
            ExportData.clear_all_sheets()
        ExportData.export_sheets()
    logger.debug('Get current gdp tasks...')
    gdp_tasks = ExportData.loadTasks(gdp_sheet_name)

    logger.debug('Get current terminalis tasks...')
    term_tasks = ExportData.loadTasks(terminalis_sheet_name)

    logger.debug('Get current error provision tasks...')
    errorProv_tasks = ExportData.loadTasks(errorProv_sheet_name)

    logger.debug('Get current renovaction tasks...')
    renov_tasks = ExportData.loadTasks(renovatcion_sheet_name)

    print(f'Comm statis ---> {ExportData.is_erro_prov_comm_found}')
    if ExportData.is_erro_prov_comm_found:
        ClarifyPage.goto_comm_table()
        errorProv_tasks_com, errorProv_agents = ClarifyPage.check_agents_names_in_tasks('com', errorProv_tasks,
                                                                                        errorProv_agents)
    else:
        errorProv_tasks_com = []
    print(f'PDTE statis ---> {ExportData.is_erro_prov_comm_found}')
    if ExportData.is_erro_prov_pdte_found:
        ClarifyPage.goto_pdte_table()
        errorProv_tasks_pdte, errorProv_agents = ClarifyPage.check_agents_names_in_tasks('pdte', errorProv_tasks,
                                                                                         errorProv_agents)
    else:
        errorProv_tasks_pdte = []
    errorProv_tasks = errorProv_tasks_com + errorProv_tasks_pdte

    logger.debug('................................................')

    logger.debug('Dispatching GDP...')
    gdp_agents = Dispatcher.dispatch_Queue(gdp_agents, gdp_tasks, tasks_per_agent_gdp, gdp_sheet_name,
                                           CaseType.GDP_QUEUE)

    logger.debug('Dispatching Terminalis...')
    terminalis_agents = Dispatcher.check_gdp_status(gdp_agents, terminalis_agents)
    terminalis_agents = Dispatcher.dispatch_Queue(terminalis_agents, term_tasks, tasks_per_agent_terminales,
                                                  terminalis_sheet_name, CaseType.TERMINALES_QUEUE)
    logger.debug('Dispatching Renovaction...')
    renovaction_agents = Dispatcher.check_terminales_status(terminalis_agents, renovaction_agents)
    Dispatcher.dispatch_Queue_renov(renovaction_agents, renov_tasks, tasks_per_agent_renov,
                                    renovatcion_sheet_name, CaseType.RENOVACION_QUEUE)

    logger.debug('Dispatching Error Prov...')
    Dispatcher.dispatch_Queue(errorProv_agents, errorProv_tasks, tasks_per_agent_error_prov, errorProv_sheet_name,
                              CaseType.ERROR_PROVISION_QUEUE)
    errorProv_agents = [agent for agent in agents_list if agent.is_handling_error_prov]
    Dispatcher.complete_error_dispatch(errorProv_agents)

    FilesUtil.take_backup(dispatching_sheet, backup_path)
    ExportData.export_informe_sheet()
    FilesUtil.take_backup(Informe_Sheet_path, backup_path)
