#!/usr/bin/python
# FINAL SCRIPT updated as of 3rd October 2019
# Workflow - REQ/DCA
# Version 1

# Declare Python libraries needed for this script
import shutil
import os
import pandas as pd
import json as json
import requests
import datetime
from pandas.io.json import json_normalize
from utils.Session import Session
from utils.guid import get_guid
from utils.audit_trail import audit_log
from utils.logging import logging
from connector.connector import *
from directory.directory_setup import prepare_directories
from transformation.mapping_jasper_sage import mapping_sage
from directory.get_filename_from_path import get_file_name
from directory.movefile import move_file_to_archived
from decorators.iterate_source import iterate_source

# Workflow function
now = datetime.datetime.now()
year=now.year
month=now.month
day=now.day

def retrieve_Jasper_file(source, destination, guid, stepname, email, taskid, username, password):
  dm_base = session.base(taskid, guid, email, step_name, username=username, password=password)
  dm_properties = session.finance(source, destination)
  try:

    # Step 1 - Preprocessing (Check if the directory for the file exists, if not create the directory)
    prepare_directories(source)
    audit_log(guid, stepname, email, "Step 1 - Check if the directory for the file exists, if not create the directory.", dm_base)
    print("Step 1 - Check if the directory for the file exists, if not create the directory.")

    #Step 2 - Open Jasper Repository, search the file, download the result as Excel File, and store in input folder
    #browser = initJasper(source)
    #download_Jasper(browser, day, month, year)
    #browser.close()
    audit_log(guid, stepname, email, "Step 2 - Open Jasper Repository, search the file, download the result as Excel File, and store in input folder", dm_base)
    print("Step 2 - Open Jasper Repository, search the file, download the result as Excel File, and store in input folder.")
    filename = 'MY_Adhoc_Accounting_Report_v3_Excel.xls' #Hardcode
    source = source+filename

    #Step 3 - Map the column in Jasper Excel File into Sage Report, create and save the Sage Report.
    mapping_sage(source, destination, day, month, year)
    move_file_to_archived(source)
    audit_log(guid, stepname, email, "Step 3 - Map the column in Jasper Excel File into Sage Report, create and save the Sage Report.", dm_base)
    print("Step 3 - Map the column in Jasper Excel File into Sage Report, create and save the Sage Report.")

  except Exception as error:
    logging("Process ARRPA", error, dm_base)

