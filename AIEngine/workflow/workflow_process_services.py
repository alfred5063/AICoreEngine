# Declare Python libraries needed for this script
#!/usr/bin/python
# FINAL SCRIPT updated as of 12th March 2021
# Workflow - STP-DM

# Declare Python libraries needed for this script
import pandas as pd
import json as json
import shutil
import os, time
import glob
import subprocess
from datetime import datetime 
from connector.connector import *
from pandas.io.json import json_normalize
from utils.Session import session
from utils.guid import get_guid
from utils.audit_trail import audit_log
from utils.notification import send
from utils.logging import logging
from pathlib import Path

# Common path
current_path = os.path.dirname(os.path.abspath(__file__))
dti_path = str(Path(current_path).parent)


# Main function starts
def trigger_service(guid, stepname, email, taskid, username, password):

  print("Session")
  try:
    stpdm_base = session.base(taskid, guid, email, stepname, username=username, password=password)
    audit_log("STP-DM - Base Directory is [ %s ]" %base_directory, "Completed...", stpdm_base)
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

  # Trigger service file
  try:
    batfile = "C:\\Users\\alfred.simbun\\source\\repos\\rpa_engine\\start_rpa_server.bat"
    p = subprocess.Popen(batfile, shell=True, stdout = subprocess.PIPE)
    stdout, stderr = p.communicate()
    if p.returncode == 0:
      print("DTI Start RPA Service")

     # is 0 if success



  else:
    print("Notify users - aborted task")
    logging("STP-DM Task Aborted - Input file cannot be found. Please check your email.", "Input Files cannot be found", stpdm_base)
    if email != '':
      send(stpdm_base, stpdm_base.email, "STP-DM Task Execution Was Not Completed.", "Hi, <br/><br/><br/>Your task was not completed.<br/>Reference Number: " + str(stpdm_base.guid) + "<br/>" +
           "<br/>Either there is no downloaded file from source system OR the user needs to add manually.<br/><br/>"+
           "<br/>Regards,<br/>Robotic Process Automation")
      audit_log("STP-DM Task Aborted - Input file cannot be found. Please check your email.", "Completed...", stpdm_base)
    raise Exception("STP-DM Task Aborted - Input file cannot be found. Please check your email.")
