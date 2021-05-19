#!/usr/bin/python
# FINAL SCRIPT updated as of 12th Nov 2020
# Workflow - Finance-AR

# Declare Python libraries needed for this script
import shutil
import sys
import os
import glob
import pandas as pd
import json as json
import requests
from datetime import datetime
from pandas.io.json import json_normalize
from utils.Session import session
from utils.guid import get_guid
from utils.notification import send
from utils.audit_trail import audit_log
from utils.logging import logging
from connector.connector import *
from directory.directory_setup import prepare_directories
from directory.movefile import *
from transformation.mapping_sage_to_costing import map_sage_to_costing
from pathlib import Path
from extraction.marc.authenticate_marc_access import get_marc_access
from pandas import ExcelWriter

# UAT - Server
#myjson = {'parent_guid': 'fe955ca2-6e83-4787-a602-0da2a6f4baed', 'guid': 'bc32f70d-41ce-4f38-8837-4f1866be22cb', 'authentication':
#          {'username': '', 'password': ''},
#          'email': 'alfred.simbun@asia-assistance.com', 'stepname': 'Update MAX', 'step_parameters': [
#            {'id': 'tr_prop_1_step_1', 'key': 'source', 'value': '\\\\dtisvr2\\Finance_UAT\\AR\\alfred.simbun@asia-assistance.com\\New\\', 'description': 'source'},
#            {'id': 'tr_prop_2_step_1', 'key': 'destination', 'value': '\\\\dtisvr2\\Finance_UAT\\AR\\alfred.simbun@asia-assistance.com\\Result\\', 'description': 'destination'}],
#          'taskId': '2441', 'execution_parameters_att_0': '', 'execution_parameters_txt': '', 'preview': 'False'}


# UAT - Local Machiene
#myjson = {'parent_guid': 'fe955ca2-6e83-4787-a602-0da2a6f4baed', 'guid': 'bc32f70d-41ce-4f38-8837-4f1866be22cb', 'authentication':
#          {'username': '', 'password': ''},
#          'email': 'nurfazreen.ramly@asia-assistance.com', 'stepname': 'Update MAX', 'step_parameters': [
#            {'id': 'tr_prop_1_step_1', 'key': 'source', 'value': 'C:\\Users\\NURFAZREEN.RAMLY\\Desktop\\RPA-TEST\\New\\', 'description': 'source'},
#            {'id': 'tr_prop_2_step_1', 'key': 'destination', 'value': 'C:\\Users\\NURFAZREEN.RAMLY\\Desktop\\RPA-TEST\\Result\\', 'description': 'destination'}],
#          'taskId': '2441', 'execution_parameters_att_0': '', 'execution_parameters_txt': '', 'preview': 'False'}


# Prod
#myjson = {'parent_guid': '4b3058df-ace4-47c3-a9c5-e57a208829b9', 'guid': '82289b4e-7b57-4aee-842f-9c530e83a65b', 'authentication':
#          {'username': 'nordiana.zainun', 'password': ''},
#          'email': 'alfred.simbun@asia-assistance.com', 'stepname': 'Update MAX', 'step_parameters': [
#            {'id': 'tr_prop_1_step_1', 'key': 'source', 'value': '\\\\10.147.78.80\\Finance\\AR\\\\alfred.simbun@asia-assistance.com\\New\\', 'description': 'source'},
#            {'id': 'tr_prop_2_step_1', 'key': 'destination', 'value': '\\\\10.147.78.80\\Finance\\AR\\\\alfred.simbun@asia-assistance.com\\Result\\', 'description': 'destination'}],
#          'taskId': '3403', 'execution_parameters_att_0': '', 'execution_parameters_txt': '', 'preview': 'False'}

#guid = str(myjson['guid'])
#stepname = str(myjson['stepname'])
#email = str(myjson['email'])
#taskid = myjson['taskId']
#password = str(myjson['authentication']['password'])
#username = str(myjson['authentication']['username'])

#for item in myjson['step_parameters']:
#  if item['key']=='source':
#    source = str(item['value'])
#  elif item['key']=='destination':
#    destination = str(item['value'])

def process_finance_sage_to_max(guid, stepname, email, taskid, username, password, source, destination):
  max_base = session.base(taskid, guid, email, stepname, username=username, password=password)
  base_directory = str(Path(source).parent)
  audit_log("Process sage to max - Base Directory is [ %s ]" %base_directory, "Completed...", max_base)
  properties = session.finance(base_directory)
  if password != '':
    max_base.set_password(get_marc_access(max_base))
  else:
    max_base.password = ''
  try:
    prepare_directories(properties.source, max_base)
    audit_log("Process sage to max", "Completed...", max_base)
    sage_excel = ''
    sage_type = ''
    for file in os.listdir(os.path.join(base_directory, 'New')):
      sage_excel = os.path.join(base_directory, 'New', file)
      if ('sage' in file.lower()):
        sage_type = 'sage'
      elif ('sg' in file.lower()):
        sage_type = 'sgpore'
    audit_log("File Search - Both file is found.", "Completed...", max_base)
  except Exception as error:
    logging("Map SAGE to Costing - Error in preparing directories", error, max_base)

  try:
    result, record_status = map_sage_to_costing(sage_excel, sage_type, max_base)
    try:
      df = pd.DataFrame(record_status)
      now = datetime.now().strftime('%d%m%Y%H%M%S')
      file_name = "result_{0}.xlsx".format(now)
      destination = os.path.join(base_directory, 'Result',file_name)
      writer = ExcelWriter(destination, engine='openpyxl')
      df.to_excel(writer, index=False)
      writer.save()
      writer.close()
      audit_log("SAGE Report - Result Moved to Results Folder.", "Completed...", max_base)
    except Exception as error:
      logging("SAGE Report - Error editing Excelsheet", error, max_base)
  except Exception as error:
    print(error)
    logging("Map SAGE to Costing", error, max_base)

  try:
    move_file_to_archived(sage_excel, max_base)
    audit_log("SAGE Report - File Moved to Archived Folder.", "Completed...", max_base)
  except Exception as error:
    logging("SAGE Report - Move File to Archived Folder.", error, max_base)

  if email != None:
    send(max_base, max_base.email, "RPA Task Execution Completed.", "Hi, <br/><br/><br/>You task has been completed.<br/>Reference Number: "+str(max_base.guid)+"<br/>"+
            "<br/>Result file as following link: <"+str(destination)+"><br/><br/>"+
            "<br/>Regards,<br/>Robotic Process Automation")
    print('Email sent.')
  else:
    print('No email address found!')

  audit_log(max_base.stepname, 'Completed...', max_base)
