#!/usr/bin/python
# FINAL SCRIPT updated as of 15th April 2021
# Call Validation of DM 8 Functions Task

# Declare Python libraries needed for this script
import pandas as pd
import json as json
import requests
import shutil
import os, time
import glob
import pyodbc as db
import win32com.client as win32
import pythoncom
import pyautogui
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
from workflow.map_action_workflow import *
from workflow.post_processing_workflow import post_processing_workflow
from extraction.marc.authenticate_marc_access import get_marc_access
from directory.splitexcelfile import split_excel_file
from directory.movefile import move_to_archived
from workflow.dm_import_marc import import_dm_to_marc


# Trigger Phase 0 RPA DM
def trigdm(guid, stepname, email, taskid, mainusername, mainpassword, MAIN, RESULTS, extracted_clist_copy, stpdm_base):

  # Prepare list of IP files as inputs
  flag = ''
  finallist = pd.DataFrame()
  param_stat = ''
  param_filetype = ''
  param_load = ''
  file_type = ''
  end_flag = ''
  try:
    print("Prepare list of IP files as inputs")
    result_flag, result_files = check_emptiness(RESULTS, stpdm_base)
    mylist = pd.DataFrame(list(filter(lambda x: '.xlsx' in x, result_files)), columns = ['Filenames'])
    for y in range(len(extracted_clist_copy['Filenames'])):
      extracted_clist_copy['Filenames'][y]
      head, tail = get_file_name(extracted_clist_copy['Filenames'][y])
      tail = tail.split('.')[0]
      currentmylist = mylist.loc[mylist['Filenames'].astype(str).str.contains(tail)]
      if str(extracted_clist_copy['Status'][y]) == 'True':
        param_stat = 'Active File'
        if str(extracted_clist_copy['OP'][y]) == 'False' and str(extracted_clist_copy['IP'][y]) == 'True':
          file_type = '_IP_'
          param_filetype = 'IP File Type was selected'
          currentmylist = currentmylist.loc[mylist['Filenames'].astype(str).str.contains(file_type)]
          finallist = finallist.append(currentmylist)
        elif str(extracted_clist_copy['OP'][y]) == 'True' and str(extracted_clist_copy['IP'][y]) == 'False':
          file_type = '_OP_'
          param_filetype = 'OP File Type was selected'
          currentmylist = currentmylist.loc[mylist['Filenames'].astype(str).str.contains(file_type)]
          finallist = finallist.append(currentmylist)
        elif str(extracted_clist_copy['OP'][y]) == 'True' and str(extracted_clist_copy['IP'][y]) == 'True':
          file_type = '_IO_'
          param_filetype = 'IO File Type was selected'
          currentmylist = currentmylist.loc[mylist['Filenames'].astype(str).str.contains(file_type)]
          finallist = finallist.append(currentmylist)
        else:
          param_filetype = 'No File Type was selected'
          pass
      else:
        param_stat = 'Inactive File'
        pass

    finallist = finallist.reset_index()
    finallist = finallist.drop(columns = {'index'})

    if finallist.empty != True:
      print("finallist is not empty")
      audit_log("STP-DM - Automatically Triggering Validation of DM 8 Functions.", "Completed...", stpdm_base)

      # Prepare directories
      try:
        phasezero_src = str(RESULTS) + "\\" + str(email) + "\\New"
        phasezero_des = str(RESULTS) + "\\" + str(email) + "\\Result"
        phasezero_arc = str(RESULTS) + "\\" + str(email) + "\\Archived"
        if os.path.exists(phasezero_src) == True and os.path.exists(phasezero_des) == True and os.path.exists(phasezero_arc) == True:
          audit_log("STP-DM - All paths for Validation of DM 8 Functions Task were found.", "Completed...", stpdm_base)
        else:
          print("Prepare directories for Phase 0 Tasks")
          createfolder(str(RESULTS) + "\\" + str(email), stpdm_base)
          createfolder(str(RESULTS) + "\\" + str(email) + "\\New", stpdm_base)
          createfolder(str(RESULTS) + "\\" + str(email) + "\\Result", stpdm_base)
          createfolder(str(RESULTS) + "\\" + str(email) + "\\Archived", stpdm_base)
          audit_log("STP-DM - New, Result and Archived folders Validation of DM 8 Functions Task were created.", "Completed...", stpdm_base)
      except Exception as error:
        logging("STP-DM - Error in creating New, Result and Archived folders for Validation of DM 8 Functions Task.", error, stpdm_base)
      time.sleep(5)

      extracted_clist = mylist.reset_index()
      extracted_clist = extracted_clist.drop(columns = {'index'})
      for f in range(len(extracted_clist['Filenames'])):
        shutil.copy(str(extracted_clist.loc[f, 'Filenames']), str(phasezero_src))
      input_flag, input_files = check_emptiness(phasezero_src, stpdm_base)
      if input_flag == True:
        mylist = pd.DataFrame(list(filter(lambda x: '.xlsx' in x, input_files)), columns = ['Filenames'])
        print("New folder is not empty")
        extracted_clist = mylist.reset_index()
        extracted_clist = extracted_clist.drop(columns = {'index'})
        flag = True
      audit_log("STP-DM - Input files for Validation of DM 8 Functions Task are ready.", "Completed...", stpdm_base)
    else:
      print("finallist is empty")
      flag = False
      audit_log("STP-DM - Input files for Validation of DM 8 Functions Task are missing.", "Completed...", stpdm_base)
  except Exception as error:
    logging("STP-DM - Error in preparing input files for Validation of DM 8 Functions Task.", error, stpdm_base)
  time.sleep(5)

  # Trigger Phase 0 backbones
  if flag == True:
    for file in range(len(extracted_clist['Filenames'])):
      files_to_process = split_excel_file(extracted_clist['Filenames'][file], phasezero_des, row_limit = 400)
      move_to_archived(extracted_clist['Filenames'][file])
      head, tail = get_file_name(extracted_clist['Filenames'][file])
      for f in range(len(files_to_process)):
        file_path = "{0}\\{1}"
        source = file_path.format(head, files_to_process[f])
        try:
          result_df = processing_validation(taskid, guid, step_name=stepname, source=source, destination=phasezero_des, email=email, username=mainusername, password=mainpassword)
          print(result_df)
          #result_df.to_excel(r"C:\Users\alfred.simbun\Desktop\DTI-TEST\OUTPUT\RUN_20210415\RESULTS\RES_20210415_155935\raw_df.xlsx")
          time.sleep(5)
          if result_df.empty != True:
            print("result_df is not empty")
            try:
              result, import_source, content = post_processing_workflow(taskid, guid, result_df, step_name=stepname, source=source, destination=phasezero_des, email=email)
              audit_log("STP-DM - Post Preprocessing for input files for Validation of DM 8 Functions Task was successful.", "Completed...", stpdm_base)
            except Exception as error:
              logging("STP-DM - Error post_processing_workflow.", error, stpdm_base)
            time.sleep(5)
            try:
              if str(result) == 'success':
                print('Import started.')
                result = import_dm_to_marc(taskid, guid, step_name=stepname, source=import_source, destination=phasezero_des, email=email, content=content, username=mainusername, password=mainpassword)
                audit_log("STP-DM - Results from Validation of DM 8 Functions Task successfully imported to MARC. Please check MARC.", "Completed...", stpdm_base)
                print('Import completed.')
              else:
                audit_log("STP-DM - 'result' is not equal to 'success'. Error import_dm_to_marc.", "Completed...", stpdm_base)
                logging("STP-DM - 'result' is not equal to 'success'", "Error import_dm_to_marc", stpdm_base)
                print('Import completed.')
            except Exception as error:
              logging("STP-DM - Error import_dm_to_marc.", error, stpdm_base)
            time.sleep(5)
            end_flag = True
            audit_log("STP-DM - Validation of DM 8 Functions Task successfully triggered. Please check the result.", "Completed...", stpdm_base)
            time.sleep(5)
          else:
            end_flag = False
            print("result_df is empty")
            audit_log("STP-DM - result_df is empty. Validation of DM 8 Functions Task has ended.", "Completed...", stpdm_base)
            time.sleep(5)

          # Email the user
          print("Notify users")
          if end_flag == True:
            send(stpdm_base, stpdm_base.email, "Validation of DM 8 Functions Task successfully triggered through File Processing through FIXER Task.",
                 "<br/>Hi, your task for: <" + str(source) + "><br/>has been completed.<br/>Reference Number: " + str(stpdm_base.guid) + "<br/>" +
                 "<br/>Results can be found through the following link: <" + str(phasezero_des) + "><br/><br/>" +
                 "<br/>Regards,<br/>Robotic Process Automation")
          elif end_flag == False:
            send(stpdm_base, stpdm_base.email, "Validation of DM 8 Functions Task was aborted.",
                 "Hi, <br/><br/><br/>Your task has be aborted. 'result_df' dataframe was empty. Kindly forward this email to dti_rpa_support email.<br/>Reference Number: " + str(stpdm_base.guid) +
                 "<br/>Hi, your task for: <" + str(source) + "><br/>has been ABORTED. 'result_df' dataframe was empty. Kindly forward this email to dti_rpa_support email.<br/>Reference Number: " + str(stpdm_base.guid) + "<br/>" +
                 "<br/>Regards,<br/>Robotic Process Automation")
        except Exception as error:
          pass
        continue
        time.sleep(5)
  else:
    print("STP-DM - Validation of DM 8 Functions Task was not triggered.")
    audit_log("STP-DM - Validation of DM 8 Functions Task was not triggered.", "Completed...", stpdm_base)


