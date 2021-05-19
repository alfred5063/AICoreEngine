#!/usr/bin/python
# FINAL SCRIPT updated as of 12th April 2021
# Workflow - STP-DM

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
import subprocess
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
from transformation.trig_dm import trigdm
from pathlib import Path
from extraction.marc.authenticate_marc_access import get_marc_access
from directory.splitexcelfile import split_excel_file

# Common path
current_path = os.path.dirname(os.path.abspath(__file__))
dti_path = str(Path(current_path).parent)


#myjson = {'parent_guid': '968ff59c-2644-4105-9471-1d6a6d20321a', 'guid': 'de49a5fa-9351-42a4-bf71-bbb8cd3d88c8', 'authentication':
#          {'username': '', 'password': ''}, 'email': 'alfred.simbun@asia-assistance.com', 'stepname': 'Scheduler', 'step_parameters': [
#            {'id': 'tr_prop_1_step_1', 'key': 'source', 'value': 'C:\\Users\\alfred.simbun\\Desktop\\DTI-TEST', 'description': ''},
#            {'id': 'tr_prop_2_step_1', 'key': 'destination', 'value': 'C:\\Users\\alfred.simbun\\Desktop\\DTI-TEST', 'description': ''},
#            {'id': 'tr_prop_3_step_1', 'key': 'gmail', 'value': 'alfred.simbun@asia-assistance.com', 'description': ''},
#            {'id': 'tr_prop_4_step_1', 'key': 'email', 'value': 'alfred.simbun@asia-assistance.com', 'description': ''},
#            {'id': 'tr_prop_5_step_1', 'key': 'password', 'value': '', 'description': ''}],
#          'taskId': '2469',
#          'execution_parameters_att_0': '',
#          'execution_parameters_txt': '',
#          'preview': 'False'}

#guid = str(myjson['guid'])
#stepname = str(myjson['stepname'])
#email = str(myjson['email'])
#taskid = myjson['taskId']
#password = str(myjson['authentication']['password'])
#username = str(myjson['authentication']['username'])

#for item in myjson['step_parameters']:
# if item['key']=='source':
#    MAIN = str(item['value'])
#    INPUT = str(item['value']) + "\\NEWFILES"
#    MAIN_FIXER = str(item['value']) + "\\FIXER"
# elif item['key']=='destination':
#    DESTINATION = str(item['value']) + "\\OUTPUT"
    #DESTINATION = str(item['value']) + "\\00_OUTPUT"


# Main function starts
def process_STPDM(guid, stepname, email, taskid, username, password, MAIN, INPUT, MAIN_FIXER, DESTINATION):

  print("Session")
  try:
    stpdm_base = session.base(taskid, guid, email, stepname, username=username, password=password)
    base_directory = str(Path(MAIN).parent)
    audit_log("STP-DM - Base Directory is [ %s ]" %base_directory, "Completed...", stpdm_base)
    properties = session.data_management(base_directory)
  except Exception as error:
    logging("STP-DM - Error in creating a session base.", error, stpdm_base)
  time.sleep(5)

  print("Date and Timestamp")
  try:
    current_year = datetime.now().year
    current_timestamp = time.strftime("%Y%m%d_%H%M%S")
    run_date = time.strftime("%Y%m%d")
  except Exception as error:
    logging("STP-DM - Error in creating a date and timestamp.", error, stpdm_base)
  time.sleep(5)

  # Prepare directories
  FAILED = ''
  ARCHIVED = ''
  RESULTS = ''
  print("Directories")
  try:
    FAILED, ARCHIVED, RESULTS = prepare_directories_dm(str(DESTINATION), str(run_date), str(current_timestamp), stpdm_base)
    audit_log("STP-DM - Directories created.", "Completed...", stpdm_base)
  except Exception as error:
    logging("STP-DM - Error in preparing directories.", error, stpdm_base)
  time.sleep(5)

  # Create TEMP folder for FIXER
  print("Temp")
  try:
    createfolder(str(ARCHIVED) + "\TEMP_" + str(guid), stpdm_base)
    createfolder(str(ARCHIVED) + "\ORIGINAL", stpdm_base)
    TEMP = str(ARCHIVED) + "\TEMP_" + str(guid)
    ORIGINAL = str(ARCHIVED) + "\ORIGINAL"
  except Exception as error:
    logging("STP-DM - Error in creating Tempt folder for FIXER.", error, stpdm_base)
  time.sleep(5)

  #rename file using batch file
  print("rename file")
  renamepath = MAIN + "\\RENAME"
  print(renamepath)
  inputflag, newfiles = check_emptiness(renamepath, stpdm_base)
  if len(newfiles) > 1: #if only bat file exist means no file needed to rename, proceed to check in newfiles
    batfile = renamepath + "\\RPA-DM_Rename_v1.0.bat"
    subprocess.call(batfile, shell=True)
    inputflag2, newfiles2 = check_emptiness(renamepath, stpdm_base)
    for j in range(len(newfiles2)):
      head, tail = get_file_name(newfiles2[j])
      if(".bat" in tail.lower()) != True:
        try:
          shutil.move(newfiles2[j], INPUT)
        except:
          pass
      else:
        pass
      continue
  else:
    pass

  time.sleep(10)

  # Calculate file size
  def humanbytes(B):
     'Return the given bytes as a human friendly KB, MB, GB, or TB string'
     B = float(B)
     KB = float(1024)
     MB = float(KB ** 2) # 1,048,576
     GB = float(KB ** 3) # 1,073,741,824
     TB = float(KB ** 4) # 1,099,511,627,776

     if B < KB:
        return '{0} {1}'.format(B,'Bytes' if 0 == B > 1 else 'Byte')
     elif KB <= B < MB:
        return '{0:.2f} KB'.format(B/KB)
     elif MB <= B < GB:
        return '{0:.2f} MB'.format(B/MB)
     elif GB <= B < TB:
        return '{0:.2f} GB'.format(B/GB)
     elif TB <= B:
        return '{0:.2f} TB'.format(B/TB)

  # Split bigger file into chunck
  onefile = ''
  input_flag, new_files = check_emptiness(INPUT, stpdm_base)
  extracted_clist = pd.DataFrame(new_files, columns = ['Filenames'])
  for e in range(len(extracted_clist['Filenames'])):
    folder, filename = get_file_name(extracted_clist['Filenames'][e])
    filename
    if '.xlsx' in str(filename.lower()) or '.xls' in str(filename.lower()):
      readxcelfile(extracted_clist['Filenames'][e])
      if str(humanbytes(os.path.getsize(extracted_clist['Filenames'][e])).split(" ")[1]) == 'KB':
        pass
      elif float(humanbytes(os.path.getsize(extracted_clist['Filenames'][e])).split(" ")[0]) <= float(5):
        pass
      else:
        split_excel_file(extracted_clist['Filenames'][e], INPUT, row_limit = 400)
      #head, tail = get_file_name(extracted_clist['Filenames'][e])
      #shutil.move(extracted_clist['Filenames'][e], os.path.join(ARCHIVED, tail))
    else:
      if '.csv' in str(filename.lower()):
        readxcelfile(extracted_clist['Filenames'][e])
        onefile = pd.read_csv(extracted_clist['Filenames'][e], encoding = "ISO-8859-1").astype(str)
      elif ".txt" in str(filename.lower()):
        seperator = find_sep(extracted_clist['Filenames'][e])
        skiprowtxt = skiprow_txt(extracted_clist['Filenames'][e], seperator)
        try:
          onefile = pd.read_csv(extracted_clist['Filenames'][e], sep = seperator, skiprows = skiprowtxt, encoding = "ISO-8859-1", header = None, dtype = str, low_memory = False)
        except:
          onefile = pd.read_csv(extracted_clist['Filenames'][e], sep = seperator, skiprows = skiprowtxt, header = None, encoding = "ISO-8859-1", dtype = str, error_bad_lines = False, low_memory = False)
          pass
      onefile = clean_file(onefile)
      head, tail = get_file_name(extracted_clist['Filenames'][e])
      onefile = onefile.rename(columns = {0:'FirstColumn'})
      onefile.to_excel(head + "\\" + str(tail.split(".")[0]) + ".xlsx", sheet_name = 'raw', index = False)
      split_excel_file(head + "\\" + str(tail.split(".")[0]) + ".xlsx", head, row_limit=400)
      shutil.move(extracted_clist['Filenames'][e], os.path.join(ORIGINAL, tail))
      shutil.move(head + "\\" + str(tail.split(".")[0]) + ".xlsx", os.path.join(ORIGINAL, tail))
    continue

  # Move input files to ARCHIVED and process from ARCHIVED
  print("Input files")
  try:
    input_flag, new_files = check_emptiness(INPUT, stpdm_base)
    if new_files != '':
      for j in range(len(new_files)):
        head, tail = get_file_name(new_files[j])
        if(".txt" in tail.lower()) != True:
          try:
            readxcelfile(new_files[j])
          except:
            pass
        else:
          pass
        if os.path.exists(new_files[j]) == True:
          print(os.path.exists(new_files[j]))
          shutil.move(new_files[j], os.path.join(ARCHIVED, tail))
      input_flag, new_files = check_emptiness(ARCHIVED, stpdm_base)
      if email != '':
        send(stpdm_base, stpdm_base.email, "STP-DM Task Execution Notification.",
             "Hi, <br/><br/><br/>Input file was found. RPA will proceed with Fixer Processing.<br/>Reference Number: " + str(stpdm_base.guid) + "<br/>" +
             "<br/>Downloaded file can be found through the following link: <" + str(ARCHIVED) + "><br/><br/>" +
             "<br/>Regards,<br/>Robotic Process Automation")
        audit_log("STP-DM - Raw files moved to ARCHIVED folder. User has been notified through email.", "Completed...", stpdm_base)
  except Exception as error:
    logging("STP-DM - Error in processing input files.", error, stpdm_base)
  time.sleep(5)

  # Database connection
  print("Connecting to database")
  myfixer = ''
  try:
    conn = MSSqlConnector()
    cur = conn.cursor()
    query = '''SELECT client_code, filename, fixer_filename, op_upload, ip_upload, full_load, status FROM [dm].[file_association]'''
    cur.execute(query)
    myfixer = pd.DataFrame(cur.fetchall(), columns=['client_code', 'filename', 'fixer_filename', 'op_upload', 'ip_upload', 'full_load', 'status'])
  except Exception as error:
    logging("STP-DM - Error in accessing database.", error, stpdm_base)
  time.sleep(5)

  # Input File
  print("Main processing")
  FIXER = ''
  extracted_clist = ''
  extracted_clist_copy = ''
  new_files = pd.DataFrame(new_files)
  new_files = new_files.rename(columns={0: 'Filenames'})

  if input_flag == True and new_files.empty != True:
    print("Main processing - filename")
    j = 0
    for j in range(len(myfixer)):
      extracted_clist = new_files.loc[new_files['Filenames'].astype(str).str.contains("%s" % str(myfixer.loc[j, 'filename']))]
      extracted_clist.insert(1, 'Fixer', "")
      extracted_clist.insert(2, 'Client', "")
      extracted_clist.insert(3, 'Identifier', "")
      extracted_clist.insert(4, 'OP', "")
      extracted_clist.insert(5, 'IP', "")
      extracted_clist.insert(6, 'Status', "")
      extracted_clist.insert(7, 'Full Load', "")
      extracted_clist = extracted_clist.reset_index()
      extracted_clist = extracted_clist.drop(columns = {'index'})
      if extracted_clist.empty != True:
        print("Dataframe is not empty")
        for k in range(len(extracted_clist['Filenames'])):
          extracted_clist.loc[k, 'Fixer'] = str(MAIN_FIXER) + "\\" + str(myfixer.loc[j, 'fixer_filename'])
          extracted_clist.loc[k, 'Client'] = str(myfixer.loc[j, 'client_code'])
          extracted_clist.loc[k, 'Identifier'] = str(myfixer.loc[j, 'filename'])
          extracted_clist.loc[k, 'OP'] = str(myfixer.loc[j, 'op_upload'])
          extracted_clist.loc[k, 'IP'] = str(myfixer.loc[j, 'ip_upload'])
          extracted_clist.loc[k, 'Status'] = str(myfixer.loc[j, 'status'])
          extracted_clist.loc[k, 'Full Load'] = str(myfixer.loc[j, 'full_load'])
        extracted_clist_copy = extracted_clist
        for l in range(len(extracted_clist)):
          print("STP-DM - Processing [ %s ]."% extracted_clist.loc[l, 'Filenames'])
          audit_log("STP-DM - Processing [ %s ]." % extracted_clist.loc[l, 'Filenames'], "Completed...", stpdm_base)
          try:
            if os.path.exists(extracted_clist.loc[l, 'Fixer']) == True and os.path.exists(TEMP) == True:
              FIXER = shutil.copy(str(extracted_clist.loc[l, 'Fixer']), str(TEMP))
              audit_log("STP-DM - Fixer file moved to TEMP folder for preprocessing.", "Completed...", stpdm_base)
            else:
              pass
          except Exception as error:
            logging("STP-DM - Error in moving FIXER to TEMP.", error, stpdm_base)
          time.sleep(5)

          print("Main processing - check empty folder")
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
          print("Main processing - begin FIXER")
          try:
            if fix_flag == True:
              head_input, tail_input = get_file_name(extracted_clist.loc[l, 'Filenames'])
              FILENAME = tail_input.split('.')[0]
              print("Filename is : %s" % FILENAME)
              head_fixer, tail_fixer = get_file_name(extracted_clist.loc[l, 'Fixer'])
              CLIENT = extracted_clist.loc[l, 'Client']
              print("Client is : %s" % CLIENT)
              PREFIX = extracted_clist.loc[l, 'Identifier'].replace(CLIENT, CLIENT + "_MARC")
              PREFIX = PREFIX + current_timestamp + "_"
              print("Prefix is : %s" % PREFIX)
              print("Source is : %s" % str(ARCHIVED + "\\" + tail_input))
              if os.path.exists(str(ARCHIVED + "\\" + tail_input)) == True:
                print("path exists")
              else:
                print("NOT")
              try:
                read_file(str(ARCHIVED + "\\" + tail_input), FILENAME, RESULTS, FAILED, CLIENT, FIXER, PREFIX, stpdm_base)
              except:
                print("Main processing - begin FIXER - FAILED")
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
        pass
      continue

    # Count total files
    print("Count total files")
    num_file = len(new_files)
    result_flag, result_files = check_emptiness(RESULTS, stpdm_base)
    failed_flag, failed_files = check_emptiness(FAILED, stpdm_base)
    files_list = open(str(RESULTS) + "\\files_summary_" + str(current_timestamp) + ".txt", "w")
    if result_files != [] and failed_files != []:
      files_list.write("List of Successfully Processed Files: \n")
      for f in range(len(result_files)):
        head_f, tail_f = get_file_name(str(result_files[f]))
        files_list.write(str(tail_f) + "\n")
      files_list.write("\n\n\nList of Un-Successfully Processed Files: \n\n")
      for i in range(len(failed_files)):
        head_file, tail_file = get_file_name(str(failed_files[i]))
        files_list.write(str(tail_file) + "\n")
      files_list.close()
    else:
      files_list.write("List of Successfully Processed Files: \n")
      for f in range(len(result_files)):
        head_f, tail_f = get_file_name(str(result_files[f]))
        files_list.write(str(tail_f) + "\n")
      files_list.close()

    # Clean-up
    print("Clean-up")
    try:
      shutil.rmtree(TEMP, ignore_errors = True)
      audit_log("STP-DM - Succesfully processed files were moved to Results Folder.", "Completed...", stpdm_base)
      audit_log("STP-DM - File clean-up.", "Completed...", stpdm_base)
    except Exception as error:
      logging("STP-DM - Error in Fixer Processing.", error, stpdm_base)


    # Notify user
    print("Notify users")
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

    # Trigger Validation of DM 8 Functions Task
    # Must trigger the IP first before OP
    try:
      config = read_db_config(dti_path+r'\config.ini', 'marc2')
      mainusername = config['user']
      mainpassword = config['password']
    except Exception as error:
      config = read_db_config(dti_path+r'\config.ini', 'marc2')
      mainusername = config['user']
      mainpassword = config['password']

    try:
      trigdm(guid, stepname, email, taskid, mainusername, mainpassword, MAIN, RESULTS, extracted_clist_copy, stpdm_base)
    except Exception as error:
      logging("STP-DM Validation of DM 8 Functions Task Aborted.", error, stpdm_base)
      pass

  else:
    print("Notify users - aborted task")
    logging("STP-DM Task Aborted - Input file cannot be found. Please check your email.", "Input Files cannot be found", stpdm_base)
    if email != '':
      send(stpdm_base, stpdm_base.email, "STP-DM Task Execution Was Not Completed.", "Hi, <br/><br/><br/>Your task was not completed.<br/>Reference Number: " + str(stpdm_base.guid) + "<br/>" +
           "<br/>Either there is no downloaded file from source system OR the user needs to add manually.<br/><br/>"+
           "<br/>Regards,<br/>Robotic Process Automation")
      audit_log("STP-DM Task Aborted - Input file cannot be found. Please check your email.", "Completed...", stpdm_base)
    raise Exception("STP-DM Task Aborted - Input file cannot be found. Please check your email.")
