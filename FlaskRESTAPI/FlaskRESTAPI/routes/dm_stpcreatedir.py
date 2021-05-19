#!/usr/bin/python
# FINAL SCRIPT updated as of 3rd Dec 2020
# Workflow - STP-DM

# Declare Python libraries needed for this script
import json
from flask import current_app as app
from flask import request
from flaskr.response.response import Success, Failure
from flaskr.decorators.login_required import login_required
from workflow.workflow_process_stpdm import process_STPDM
from utils.Session import session
from directory.directory_setup import prepare_directories
from directory.createfolder import *
from utils.audit_trail import audit_log
from utils.logging import logging

@app.route('/process_dmstpfolder', methods = ['POST'])
@login_required
def process_dmstpfolder():
  try:

    myjson = json.loads(request.form['json'])
    guid = str(myjson['guid'])
    stepname = str(myjson['stepname'])
    email = str(myjson['email'])
    taskid = myjson['taskId']
    client = myjson['execution_parameters_txt']['parameters'][0]['Client']
    password = str(myjson['authentication']['password'])
    username = str(myjson['authentication']['username'])

    for item in myjson['step_parameters']:
      if item['key']=='source':
        source = str(item['value']) + "\\" + str(email) + "\\" + str(client) + "\\New"
        main_path = str(item['value'])
      elif item['key']=='destination':
        destination = str(item['value']) + "\\" + str(email) + "\\" + str(client) + "\\Result"

    # Call out main function from workflow_process_stpdm.py
    stpdm_base = session.base(taskid, guid, email, stepname, username=username, password=password)
    try:
      base_directory = str(Path(source).parent)
      audit_log("STP-DM - Base Directory is [ %s ]" %base_directory, "Completed...", stpdm_base)
      properties = session.data_management(base_directory)
    except Exception as error:
      logging("STP-DM - Error in creating a session base.", error, stpdm_base)

    # Prepare directories
    try:
      prepare_directories(properties.source, stpdm_base)
      audit_log("STP-DM - Directories created.", "Completed...", stpdm_base)
    except Exception as error:
      logging("STP-DM - Error in preparing directories.", error, stpdm_base)

    return Success({'message':'Completed.'})
  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )
