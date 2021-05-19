#!/usr/bin/python
# FINAL SCRIPT updated as of 29th April 2020
# Workflow - CBA/MEDICLINIC
# Version 1

# Declare Python libraries needed for this script
from transformation.medi_file_manipulation import file_manipulation, all_excel_move_to_archive, create_path, createfolder, move_file, get_all_File
from transformation.medi_medi import Medi_Mos
from transformation.medi_bdx_manipulation import bdx_automation
from transformation.medi_update_dcm import update_to_dcm
from transformation.medi_generate_dc import generate_dc
from pandas.io.json import json_normalize
from utils.Session import session
from utils.guid import get_guid
from utils.audit_trail import audit_log
from utils.logging import logging
from extraction.marc.authenticate_marc_access import get_marc_access
import pandas as pd
import traceback,os
from utils.notification import send
from datetime import datetime
from openpyxl import *
from directory.directory_setup import prepare_directories
from directory.movefile import *
from transformation.populate_data_medi import *

def process_medMOD(myjson, guid, stepname, email, taskid, password, username, source, input, destination):

  medi_base = session.base(taskid, guid, email, stepname)
  medi_properties = session.finance(source, destination)

  audit_log("MediClinic MOS - Check if the directory for the file exists, if not create the directory", "Completed...", medi_base)
  try:
    prepare_directories(medi_properties.source, medi_base)
    audit_log("MediClinic MOS - Directory exists or created", "Completed...", medi_base)
  except Exception as error:
    audit_log("MediClinic MOS - Error in preparing required directories. Please check in the Failed Task List", "Completed...", medi_base)
    logging('MediClinic MOS - Directory exists or created', error, medi_base)

  try:
    bordlisting = disbursementClaim = ""
    for file in os.listdir(input):
        if ('web' in file.lower()):
          bordlisting = input + file
        elif ('mcl' in file.lower()):
          disbursementClaim = input + file
        elif ('running' in file.lower()):
          disbursementMaster = input + file
    audit_log("MediClinic MOS - Searching for Bord Listing and Disbursement Claim files in New folder", "Completed...", medi_base)
  except Exception as error:
    audit_log("MediClinic MOS - Error in Searching for Bord Listing and Disbursement Claim files in New folder. Please check in the Failed Task List", "Completed...", medi_base)
    logging('MediClinic MOS - Searching for Bord Listing and Disbursement Claim files in New folder', error, medi_base)

  try:
    populate_data_medi(disbursementClaim, bordlisting, disbursementMaster, destination, medi_base)
    audit_log("MediClinic MOS - Successfully performed mapping", "Completed...", medi_base)
  except Exception as error:
    audit_log("MediClinic MOS - Error in mapping file. Please check in the Failed Task List", "Completed...", medi_base)
    logging('MediClinic MOS - Successfully performed mapping', error, medi_base)

  try:
    move_file_to_archived(bordlisting, medi_base)
    move_file_to_archived(disbursementClaim, medi_base)
    move_file_to_archived(disbursementMaster, medi_base)
    audit_log("MediClinic MOS - Moving input files to Archived folder", "Completed...", medi_base)
  except Exception as error:
    audit_log("Finance SOA - Error in moving input files to Archived folder", "Completed...", medi_base)
    logging('MediClinic MOS - Moving input files to Archived folder', error, medi_base)

  if email != None:
    send(medi_base, email, "RPA Task Execution Completed.", "Hi, <br/><br/><br/>You task has been completed. <br/>Reference Number: " + str(medi_base.guid) + "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
  audit_log("Finance SOA - RPA task execution completion notification. Sending notification through email to the user.", "Completed...", medi_base)

  return "Completed"


