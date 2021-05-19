#!/usr/bin/python
# FINAL SCRIPT updated as of 15th April 2021
# Workflow - Finance SOA
# Version 1

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
from transformation.populate_data_soa import *


myjson = {'parent_guid': 'c0e2dcce-4eb2-4d00-903b-6321b70e944b',
         'guid': '6bf1fee8-0c4e-4fd4-8451-766f9788b422',
         'authentication': {'username': 'alfred.simbun', 'password': '!!@#Acgt123'},
        'email': 'iyliaasyqin.abdmajid@asia-assistance.com', 'stepname': 'Populate Hospital SOA File',
        'step_parameters': [{'id': 'tr_prop_1_step_1', 'key': 'source', 'value': '\\\\dtisvr2\\Finance_UAT\\SOA\\iyliaasyqin.abdmajid@asia-assistance.com\\New\\',
                            'description': 'Location of File to be Processed'},
                            {'id': 'tr_prop_2_step_1', 'key': 'destination', 'value': '\\\\dtisvr2\\Finance_UAT\\SOA\\iyliaasyqin.abdmajid@asia-assistance.com\\Result\\',
                            'description': 'Location of the result file'}],
       'taskId': '1383', 'execution_parameters_att_0': '', 'execution_parameters_txt': '', 'preview': 'False'}
guid = "6bf1fee8-0c4e-4fd4-8451-766f9788b422"
stepname = "Populate Hospital SOA File"
email = "iyliaasyqin.abdmajid@asia-assistance.com"
taskid = "1383"

def process_FinanceSOA(myjson, guid, stepname, email, taskid):

  for item in myjson['step_parameters']:
    if item['key']=='source':
      input = str(item['value'])
      print(input)
      source = str(item['value'].split("\\New\\")[0])
    elif item['key']=='destination':
      destination = source

  soarpa_base = session.base(taskid, guid, email, stepname)
  soarpa_properties = session.finance(source,destination)
  try:
    audit_log("Finance SOA - Check if the directory for the file exists, if not create the directory", "Completed...", soarpa_base)
    try:
      prepare_directories(soarpa_properties.source, soarpa_base)
      audit_log("Finance SOA - Directory exists or created", "Completed...", soarpa_base)
    except Exception as error:
      audit_log("Finance SOA - Error in preparing required directories. Please check in the Failed Task List", "Completed...", soarpa_base)
      logging('Finance SOA - Directory exists or created', error, soarpa_base)

    try:
      soa_excel = vt_excel = clear_date_excel = ""
      for file in os.listdir(input):
        if ('soa' in file.lower()):
          soa_excel = str(input) + str(file)
        elif ('vendor' in file.lower()):
          vt_excel = str(input) + str(file)
        elif ('bank' in file.lower()):
          clear_date_excel = str(input) + str(file)

      audit_log("Finance SOA - Searching for Bank Clear Date and Vendor Transaction files in New folder", "Completed...", soarpa_base)
    except Exception as error:
      audit_log("Finance SOA - Error in Searching for Bank Clear Date and Vendor Transaction files in New folder. Please check in the Failed Task List", "Completed...", soarpa_base)
      logging('Finance SOA - Searching for Bank Clear Date and Vendor Transaction files in New folder', error, soarpa_base)

    vendor_transaction(vt_excel, soarpa_base)

    try:
      populate_data_SOA(soa_excel,vt_excel,clear_date_excel, soarpa_base)
      audit_log("Finance SOA - Populating SOA data from MARC", "Completed...", soarpa_base)
    except Exception as error:
      print(error)
      audit_log("Finance SOA - Error in Populating SOA data from MARC. Please check in the Failed Task List", "Completed...", soarpa_base)
      logging('Finance SOA - Populating SOA data from MARC', error, soarpa_base)

    try:
      move_file_to_result(soa_excel, soarpa_base)
      audit_log("Finance SOA - SOA results has been moved to Result folder", "Completed...", soarpa_base)
    except Exception as error:
      audit_log("Finance SOA - Error in moving SOA results to Result folder. Please check in the Failed Task List", "Completed...", soarpa_base)
      logging('Finance SOA - Moving SOA results to Result folder', error, soarpa_base)

    try:
      move_file_to_archived(vt_excel, soarpa_base)
      move_file_to_archived(clear_date_excel, soarpa_base)
      audit_log("Finance SOA - Moving Sample of Bank Clear Date Data and Vendor Transaction files to Archived folder", "Completed...", soarpa_base)
    except Exception as error:
      audit_log("Finance SOA - Error in moving Sample of Bank Clear Date Data and Vendor Transaction files to Archived folder. Please check in the Failed Task List", "Completed...", soarpa_base)
      logging('Finance SOA - Moving Sample of Bank Clear Date Data and Vendor Transaction files to Archived folder', error, soarpa_base)

    if email != None:
      send(soarpa_base, email, "RPA Task Execution Completed.", "Hi, <br/><br/><br/>You task has been completed. <br/>Reference Number: " + str(soarpa_base.guid) + "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
    audit_log("Finance SOA - RPA task execution completion notification. Sending notification through email to the user.", "Completed...", soarpa_base)
  except Exception as error:
    logging("Finance SOA Workflow", error, soarpa_base)

