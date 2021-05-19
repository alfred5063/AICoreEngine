#!/usr/bin/python
# FINAL SCRIPT updated as of 15th June 2020
# Workflow - Finance Credit Control

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
from transformation.populate_mem_letter_creditcontrol import process_cc
from extraction.marc.authenticate_marc_access import get_marc_access

def process_credit_control(myjson, guid, stepname, email, taskid, username, password):

  for item in myjson['step_parameters']:
    if item['key']=='source':
      input = str(item['value'])
      source = str(item['value'].split("\\New\\")[0])
    elif item['key']=='destination':
      destination = source

  creditcontrol_base = session.base(taskid, guid, email, stepname, username, password)
  creditcontrol_properties = session.finance(source, destination)

  if password != '':
    creditcontrol_base.set_password(get_marc_access(creditcontrol_base))
  else:
    creditcontrol_base.password = ''

  try:

    audit_log("Credit Control - Generate Member Letter - Check if the directory for the file exists, if not create the directory", "Completed...", creditcontrol_base)
    try:
      prepare_directories(creditcontrol_properties.source, creditcontrol_base)
      audit_log("Credit Control - Generate Member Letter - Directory exists or created", "Completed...", creditcontrol_base)
    except Exception as error:
      audit_log("Credit Control - Generate Member Letter - Error in preparing required directories. Please check in the Failed Task List", "Completed...", creditcontrol_base)
      logging('Credit Control - Generate Member Letter- Directory exists or created', error, creditcontrol_base)

    try:
      audit_log("Credit Control - File Search - Searching for input files in New folder.", "Completed...", creditcontrol_base)
      excelf = ''
      docf = ''
      for file in os.listdir(input):
        if ('mem' in file.lower()):
          excelf = str(input) + str(file)
        elif ('lod' in file.lower()):
          docf = str(input) + str(file)

      audit_log("Credit Control - Both file are found.", "Completed...", creditcontrol_base)
    except Exception as error:
      audit_log("Credit Control - Error in preparing required directories. Please check in the Failed Task List.", "Completed...", creditcontrol_base)
      logging('Credit Control - Error in preparing required directories. Please check in the Failed Task List.', error, creditcontrol_base)

    try:
      process_cc(excelf, docf, creditcontrol_base)
      audit_log("Credit Control - Successfully processed credit control.", "Completed...", creditcontrol_base)
    except Exception as error:
      audit_log("Credit Control - Error in processing credit control. Please check in the Failed Task List.", "Completed...", creditcontrol_base)
      logging('Credit Control - Error in processing credit control. Please check in the Failed Task List.', error, creditcontrol_base)

    try:
      move_file_to_archived(excelf, creditcontrol_base)
      move_file_to_archived(docf, creditcontrol_base)
      audit_log("Credit Control - All results has been moved to Result folder. Please check.", "Completed...", creditcontrol_base)
    except Exception as error:
      audit_log("Credit Control - Generate Member Letter - Error in moving document results to Result folder. Please check in the Failed Task List", "Completed...", creditcontrol_base)
      logging('Credit Control - Generate Member Letter - Moving Member letter results to Result folder', error, creditcontrol_base)


    if email != None:
      send(creditcontrol_base, email, "RPA Task Execution Completed.", "Hi, <br/><br/><br/>You task has been completed. <br/>Reference Number: " + str(creditcontrol_base.guid) + "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
    audit_log("Generate Member Letter - RPA task execution completion notification. Sending notification through email to the user.", "Completed...", creditcontrol_base)
  except Exception as error:
    logging("Generate Member Letter Workflow", error, creditcontrol_base)

