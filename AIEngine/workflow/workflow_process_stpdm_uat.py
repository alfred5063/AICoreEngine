# Declare Python libraries needed for this script
#!/usr/bin/python
# FINAL SCRIPT updated as of 29th Dec 2020
# Workflow - STP-DM

# Declare Python libraries needed for this script
import pandas as pd
import json as json
import requests
import shutil
import os, time
import glob
from loading.excel.checkExcelLoading import *
from datetime import datetime 
from connector.connector import *
from pandas.io.json import json_normalize
from utils.Session import session
from utils.guid import get_guid
from utils.audit_trail import audit_log
from utils.notification import send
from utils.logging import logging
from transformation.dm_file_preparation import read_clientfile, convert_to_xlsx
from transformation.dm_fixer import *
from directory.files_listing_operation import getListOfFiles
from directory.directory_setup import prepare_directories_dm
from directory.createfolder import *
from directory.get_filename_from_path import *
from directory.movefile import *
from directory.check_emptyfolder import *
from utils.logging import logging
from connector.connector import MySqlConnector, MSSqlConnector
from connector.dbconfig import *
import pyodbc as db
import win32com.client as win32
import pythoncom

#myjson = {'parent_guid': '2ef57adf-b727-4cb9-8f75-762cfbcb2173', 'guid': '27e18d4f-db47-4b68-9e07-b91734b585c0', 'authentication': {
#  'username': 'alfred.simbun@asia-assistance.com', 'password': 'DZ7viemSKo0ydOTOeF9c8Q=='},
#          'email': 'alfred.simbun@asia-assistance.com', 'stepname': 'Run FIXER File Processing', 'step_parameters': [
#            {'id': 'tr_prop_1_step_2', 'key': 'source', 'value': '\\\\dtisvr2\\DM_Demo\\STP.DM\\DTI-TEST', 'description': ''},
#            {'id': 'tr_prop_2_step_2', 'key': 'destination', 'value': '\\\\dtisvr2\\DM_Demo\\STP.DM\\DTI-TEST', 'description': ''}],
#          'taskId': '2454', 'execution_parameters_att_0': '', 'execution_parameters_txt': '', 'preview': 'False'}

#guid = str(myjson['guid'])
#stepname = str(myjson['stepname'])
#email = str(myjson['email'])
#taskid = myjson['taskId']
#password = str(myjson['authentication']['password'])
#username = str(myjson['authentication']['username'])

#for item in myjson['step_parameters']:
#  if item['key']=='source':
#    MAIN = str(item['value'])
#    INPUT = str(item['value']) + "\\NEWFILES"
#    MAIN_FIXER = str(item['value']) + "\\FIXER"
#  elif item['key']=='destination':
#    DESTINATION = str(item['value']) + "\\OUTPUT"

# Main function starts
def process_STPDM(guid, stepname, email, taskid, username, password, MAIN, INPUT, MAIN_FIXER, DESTINATION):

  current_year = datetime.now().year
  current_timestamp = time.strftime("%Y%m%d_%H%M%S")
  run_date = time.strftime("%Y%m%d")

  try:
    stpdm_base = session.base(taskid, guid, email, stepname, username=username, password=password)
    base_directory = str(Path(MAIN).parent)
    audit_log("STP-DM - Base Directory is [ %s ]" %base_directory, "Completed...", stpdm_base)
    properties = session.data_management(base_directory)
  except Exception as error:
    logging("STP-DM - Error in creating a session base.", error, stpdm_base)
  time.sleep(5)

  # Prepare directories
  FAILED = ''
  ARCHIVED = ''
  RESULTS = ''
  try:
    FAILED, ARCHIVED, RESULTS = prepare_directories_dm(str(DESTINATION), str(run_date), str(current_timestamp), stpdm_base)
    audit_log("STP-DM - Directories created.", "Completed...", stpdm_base)
  except Exception as error:
    logging("STP-DM - Error in preparing directories.", error, stpdm_base)
  time.sleep(5)

  # Create TEMP folder for FIXER
  try:
    createfolder(str(ARCHIVED) + "\TEMP_" + str(guid), stpdm_base)
    TEMP = str(ARCHIVED) + "\TEMP_" + str(guid)
  except Exception as error:
    logging("STP-DM - Error in creating Tempt folder for FIXER.", error, stpdm_base)
  time.sleep(5)

  # Move input files to ARCHIVED and process from ARCHIVED
  #t_end = time.time() + 60 * 15
  #while time.time() < t_end:
  input_flag, new_files = check_emptiness(INPUT, stpdm_base)
  if new_files != '':
    for j in range(len(new_files)):
      head, tail = get_file_name(new_files[j])
      if(".txt" in tail.lower()) != True:
        print("a")
        print(new_files[j].lower())
        pythoncom.CoInitialize()
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(new_files[j])
        print(wb)
        excel.EnableEvents = False
        excel.DisplayAlerts = False
        excel.Visible = False
        wb.Save()
        wb.Close(True)
        excel.Quit()
      else:
        print("b")
        pass
      head, tail = get_file_name(new_files[j])
      shutil.move(os.path.join(head, tail), os.path.join(ARCHIVED, tail))
    input_flag, new_files = check_emptiness(ARCHIVED, stpdm_base)
    if email != '':
      send(stpdm_base, stpdm_base.email, "STP-DM Task Execution Notification.",
           "Hi, <br/><br/><br/>Input file was found. RPA will proceed with Fixer Processing.<br/>Reference Number: " + str(stpdm_base.guid) + "<br/>" +
           "<br/>Downloaded file can be found through the following link: <" + str(ARCHIVED) + "><br/><br/>" +
           "<br/>Regards,<br/>Robotic Process Automation")
      audit_log("STP-DM - Raw files moved to ARCHIVED folder. User has been notified through email.", "Completed...", stpdm_base)

  # Database connection
  print("- Connecting to database")
  conn = MSSqlConnector()
  cur = conn.cursor()
  query = '''SELECT client_code, filename, fixer_filename FROM [dm].[file_association]'''
  cur.execute(query)
  myfixer = pd.DataFrame(cur.fetchall(), columns=['client_code', 'filename', 'fixer_filename'])

  # Input File
  FIXER = ''
  if input_flag == True and new_files != []:
    for j in range(len(myfixer['filename'])):
      j
      myfixer['filename'][j]
      extracted_clist = pd.DataFrame()
      extracted_clist = pd.DataFrame(list(filter(lambda x: str(myfixer['filename'][j]) in x, new_files)), columns = ['Filenames'])

      if extracted_clist.empty != []:
        extracted_clist['Fixer'] = ''
        extracted_clist['Client'] = ''
        extracted_clist['Identifier'] = ''
        for k in range(len(extracted_clist['Filenames'])):
          extracted_clist['Fixer'][k] = str(MAIN_FIXER) + "\\" + str(myfixer['fixer_filename'][j])
          extracted_clist['Client'][k] = str(myfixer['client_code'][j])
          extracted_clist['Identifier'] = myfixer['filename'][j]
        for l in range(len(extracted_clist['Filenames'])):
          print("STP-DM - Processing [ %s ]."% extracted_clist['Filenames'][l])
          audit_log("STP-DM - Processing [ %s ]." % extracted_clist['Filenames'][l], "Completed...", stpdm_base)
          try:
            FIXER = shutil.copy(str(extracted_clist['Fixer'][l]), str(TEMP))
            audit_log("STP-DM - Fixer file moved to TEMP folder for preprocessing.", "Completed...", stpdm_base)
          except Exception as error:
            logging("STP-DM - Error in moving FIXER to TEMP.", error, stpdm_base)
          time.sleep(5)

          t_end = time.time() + 60 * 15
          while time.time() < t_end:
            fix_flag, fix_files = check_emptiness(TEMP, stpdm_base)
            if fix_files != '':
              fix_flag = True
              print("FIXER is found")
              audit_log("STP-DM - Fixer file moved to TEMP folder for preprocessing.", "Completed...", stpdm_base)
              break
            else:
              fix_flag = False

          # Begin FIXER
          try:
            if fix_flag == True:
              head_input, tail_input = get_file_name(extracted_clist['Filenames'][l])
              FILENAME = tail_input.split('.')[0]
              head_fixer, tail_fixer = get_file_name(extracted_clist['Fixer'][l])
              CLIENT = extracted_clist['Client'][l]
              PREFIX = extracted_clist['Identifier'][l].replace(CLIENT, CLIENT + "_MARC")
              PREFIX = PREFIX + current_timestamp + "_"
              read_file(str(ARCHIVED + "\\" + tail_input), FILENAME, RESULTS, FAILED, CLIENT, FIXER, PREFIX, stpdm_base)
              audit_log("STP-DM - FIXER Processing.", "Completed...", stpdm_base)
              print("FIXER Processing Completed")
              os.remove(FIXER)
            else:
              logging("STP-DM - Fixer cannot be loaded.", "Fixer file cannot be found", stpdm_base)
              raise Exception("STP-DM Task Aborted - Fixer cannot be loaded..")
          except Exception as error:
            os.remove(FIXER)
            logging("STP-DM - Error in Fixer Processing.", error, stpdm_base)
          time.sleep(5)
      else:
        logging("STP-DM - Fixer cannot be loaded.", "Fixer file cannot be found", stpdm_base)
        if email != '':
          send(stpdm_base, stpdm_base.email, "STP-DM Task Execution Was Not Completed.",
               "Hi, <br/><br/><br/>Your task was not completed.<br/>Reference Number: " + str(stpdm_base.guid) +
               "<br/>" +
               "<br/>FIXER file cannot be found.<br/><br/>" +
               "<br/>Regards,<br/>Robotic Process Automation")
          audit_log("STP-DM Task Aborted - FIXER file cannot be found.", "Completed...", stpdm_base)
          raise Exception("STP-DM Task Aborted - FIXER file cannot be found.")

    # Count total files
    num_file = len(new_files)
    result_flag, result_files = check_emptiness(RESULTS, stpdm_base)
    failed_flag, failed_files = check_emptiness(FAILED, stpdm_base)
    files_list = open(str(RESULTS) + "\\files_summary_" + str(current_timestamp) + ".txt", "w")
    if result_files != [] and failed_files != []:
      files_list.write("List of Successfully Processed Files: \n")
      files_list.write(str(result_files))
      files_list.write("\n\n\nList of Un-Successfully Processed Files: \n")
      files_list.write(str(failed_files))
      files_list.close()
    else:
      files_list.write("List of Successfully Processed Files: \n")
      files_list.write(str(result_files))
      files_list.close()

    # Clean-up
    try:
      result_flag, result_files = check_emptiness(RESULTS, stpdm_base)
      shutil.rmtree(TEMP, ignore_errors = True)
      audit_log("STP-DM - Succesfully processed files were moved to Results Folder.", "Completed...", stpdm_base)
      audit_log("STP-DM - File clean-up.", "Completed...", stpdm_base)
    except Exception as error:
      logging("STP-DM - Error in Fixer Processing.", error, stpdm_base)

    # Notify user
    if email != '':
      send(stpdm_base, stpdm_base.email, "STP-DM Task Execution Completed.",
           "Hi, <br/><br/><br/>Your task has been completed.<br/>Reference Number: " + str(stpdm_base.guid) +
           "<br/>" +
           "<br/>Results can be found through the following link: <" + str(DESTINATION) + "\\RUN_" + str(run_date) + "><br/><br/>" +
           "<br/>The total number of input files for this session was: " + str(num_file) + 
           "<br/>The total number of successfully processed files for this session is: " + str(len(result_files)) +
           "<br/>The total number of un-successfully processed files for this session is: " + str(len(failed_files)) +
           "<br/>Run Summary can be found through the following link: <" + str(RESULTS) + "\\files_summary_" + str(current_timestamp) + ".txt" + "><br/><br/>" +
           "<br/>Regards,<br/>Robotic Process Automation")
    else:
      pass

    audit_log("STP-DM Task Completion - User has been notified through email.", "Completed...", stpdm_base)

  else:
    logging("STP-DM Task Aborted - Input file cannot be found. Please check your email.", "Input Files cannot be found", stpdm_base)
    if email != '':
      send(stpdm_base, stpdm_base.email, "STP-DM Task Execution Was Not Completed.", "Hi, <br/><br/><br/>Your task was not completed.<br/>Reference Number: " + str(stpdm_base.guid) + "<br/>" +
           "<br/>Either there is no downloaded file from source system OR the user needs to add manually.<br/><br/>"+
           "<br/>Regards,<br/>Robotic Process Automation")
      audit_log("STP-DM Task Aborted - Input file cannot be found. Please check your email.", "Completed...", stpdm_base)
    raise Exception("STP-DM Task Aborted - Input file cannot be found. Please check your email.")

