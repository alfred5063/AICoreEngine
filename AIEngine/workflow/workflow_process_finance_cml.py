#!/usr/bin/python
# FINAL SCRIPT updated as of 16th June 2020
# Workflow - Finance Update CML
# Version 1

# Declare Python libraries needed for this script
import shutil
import sys
import os
import glob
import pandas as pd
import json as json
import requests
from datetime import datetime as dt
from pandas.io.json import json_normalize
from utils.Session import session
from utils.guid import get_guid
from utils.notification import send
from utils.audit_trail import audit_log
from utils.logging import logging
from connector.connector import *
from directory.directory_setup import prepare_directories
from directory.movefile import *
from directory.files_listing_operation import search_cmlfile, getListOfFiles
from loading.excel.checkExcelLoading import *
from transformation.finance_update_cm import *

def process_Finance_CML(myjson, guid, stepname, email, taskid):

  for item in myjson['step_parameters']:
    if item['key']=='source':
      input = str(item['value'])
      source = str(item['value'].split("\\New\\")[0])
    elif item['key']=='destination':
      destination = source
    elif item['key']=='cml_path':
      cml_path = str(item['value'])

  # Process JSON string to a dataframe
  jsondf = pd.DataFrame(json_normalize(myjson['execution_parameters_txt']['parameters']))

  # Set sessions
  cmlrpa_base = session.base(taskid, guid, email, stepname)
  cmlrpa_properties = session.finance(source, destination)

  # Determine the current year and date
  current_year = dt.now().year
  current_date = time.strftime("%Y-%m-%d")

  try:

    audit_log("Finance Update CML - Check if the directory for the file exists, if not create the directory", "Completed...", cmlrpa_base)
    try:
      prepare_directories(cmlrpa_properties.source, cmlrpa_base)
      audit_log("Finance Update CML - Directory exists or created", "Completed...", cmlrpa_base)
    except Exception as error:
      logging('Finance Update CML - Error in preparing required directories. Please check in the Failed Task List', error, cmlrpa_base)

    # Process extracted report from MARC and payment listing from user
    try:
      myclients = search_cmlfile(jsondf, cml_path, current_year)
      userinputs = getListOfFiles(input)
      paymenlisting = userinputs[0]
      marcfile = userinputs[1]

      filestatus_paymenlisting = is_locked(paymenlisting)
      filestatus_marcfile = is_locked(marcfile)

      if str(filestatus_paymenlisting) == 'False' and str(filestatus_marcfile) == 'False':
        audit_log("Finance Update CML - Both Payment Listing and Extracted MARC Report were provided by the user.", "Completed...", cmlrpa_base)
      else:
        audit_log("Finance Update CML - Either Payment Listing OR Extracted MARC Report is missing or locked. Please check. This entire RPA process has been aborted.", "Completed...", cmlrpa_base)
        logging('Finance Update CML.', 'Either Payment Listing OR Extracted MARC Report is missing or locked. Please check. This entire RPA process has been aborted.', cmlrpa_base)
    except Exception as error:
      logging('Finance Update CML - Process extracted report from MARC and payment listing from user', error, cmlrpa_base)

    # Process data
    try:
      preprosessing_cmlinfo(myclients, marcfile, paymenlisting, cmlrpa_base)
      audit_log("Finance Update CML - Processing Payment Listing and MARC Report.", "Completed...", cmlrpa_base)
    except Exception as error:
      audit_log("Finance Update CML - Error in Processing Payment Listing and MARC Report.. Please check in the Failed Task List", "Completed...", cmlrpa_base)
      logging('Finance Update CML - Processing Payment Listing and MARC Report.', error, cmlrpa_base)

    try:
      move_file_to_archived(paymenlisting, cmlrpa_base)
      move_file_to_archived(marcfile, cmlrpa_base)
      audit_log("Finance Update CML - Moving Payment Listing and MARC Report to Archived folder", "Completed...", cmlrpa_base)
    except Exception as error:
      audit_log("Finance Update CML - Error in moving Payment Listing and MARC Report to Archived folder. Please check in the Failed Task List", "Completed...", cmlrpa_base)
      logging('Finance Update CML - Moving Payment Listing and MARC Report to Archived folder', error, cmlrpa_base)

    if email != None:
      send(cmlrpa_base, email, "RPA Task Execution Completed.", "Hi, <br/><br/><br/>You task has been completed. <br/>Reference Number: " + str(cmlrpa_base.guid) + "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
    audit_log("Finance Update CML - RPA task execution completion notification. Sending notification through email to the user.", "Completed...", cmlrpa_base)
  except Exception as error:
    logging("Finance Update CML Workflow", error, cmlrpa_base)

