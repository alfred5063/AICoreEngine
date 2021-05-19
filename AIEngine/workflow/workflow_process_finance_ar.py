#!/usr/bin/python
# FINAL SCRIPT updated as of 27th May 2020
# Workflow - Finance-AR

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
from transformation.mapping_jasper_sage import mapping_sage
from transformation.mapping_sage_to_costing import *

def process_Finance_AR(myjson, guid, stepname, email, taskid):
  for item in myjson['step_parameters']:
    if item['key']=='source':
      input = str(item['value'])
      source = str(item['value'].split("\\New\\")[0])
    elif item['key']=='destination':
      destination = source

  arrpa_base = session.base(taskid, guid, email, stepname)
  arrpa_properties = session.finance(source, destination)

  try:
    audit_log("Finance AR - Preprocessing - Check if the directory for the file exists, if not create the directory.", "Completed...", arrpa_base)
    prepare_directories(arrpa_properties.source, arrpa_base)
    audit_log("Finance AR - Preprocessing - Directory exists or created.", "Completed...", arrpa_base)

    audit_log("Finance AR - File Search - Searching for Jasper and Sage file in New folder.", "Completed...", arrpa_base)
    template_excel = ''
    jasper_excel = ''
    template_type = ''  
    for file in os.listdir(input):
      if ('sage' in file.lower()):
        template_excel = str(input) + str(file)
        template_type = 'sage'
      elif ('sg' in file.lower()):
        template_excel = str(input) + str(file)
        template_type = 'sgpore'
      else:
        jasper_excel = str(input) + str(file)

    audit_log("Finance AR - File Search - Both file is found.", "Completed...", arrpa_base)
    now = dt.now()

    try:
      mapping_sage(template_type, template_excel, jasper_excel, now, arrpa_base)
    except Exception:
      try:
        mapping_sage(template_type, template_excel, jasper_excel, now, arrpa_base)
      except Exception as error:
        logging("Process Finance AR - Mapping Sage", error, arrpa_base)

    #try:
    #  map_sage_to_costing(source, template_excel, template_type, arrpa_base)
    #except Exception:
    #  try:
    #    map_sage_to_costing(source, template_excel, template_type, arrpa_base)
    #  except Exception as error:
    #    logging("Process Finance AR - Mapping Sage to Costing", error, arrpa_base)

    try:
      move_file_to_result_ar(template_excel, arrpa_base)
      move_file_to_archived(jasper_excel, arrpa_base)
    except Exception:
      try:
        move_file_to_result_ar(template_excel, arrpa_base)
        move_file_to_archived(jasper_excel, arrpa_base)
      except Exception as error:
        logging("Process Finance AR - Move Files to Result and Archived", error, arrpa_base)


    audit_log("Finance AR - Sage Report - Sage Report updated.", "Completed...", arrpa_base)

    if email != None:
      send(arrpa_base, email, "RPA Task Execution Completed.", "Hi, <br/><br/><br/>You task has been completed. <br/>Reference Number: " + str(guid) + "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
    audit_log(arrpa_base.stepname, "Completed", arrpa_base)
  except Exception as error:
    logging("Process Finance AR - Main Workflow", error, arrpa_base)
