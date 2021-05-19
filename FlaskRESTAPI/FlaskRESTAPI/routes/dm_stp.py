#!/usr/bin/python
# FINAL SCRIPT updated as of 6th Jan 2021
# Workflow - STP-DM

# Declare Python libraries needed for this script
import json
from datetime import datetime as dt
import datetime
import time
from flask import current_app as app
from flask import request
from flaskr.response.response import Success, Failure
from flaskr.decorators.login_required import login_required
from workflow.workflow_process_stpdm import process_STPDM

@app.route('/process_dmstp', methods = ['POST'])
@login_required
def process_dmstp():
  try:

    myjson = json.loads(request.form['json'])
    guid = str(myjson['guid'])
    stepname = str(myjson['stepname'])
    #email = str(myjson['email'])
    taskid = myjson['taskId']
    password = str(myjson['authentication']['password'])
    username = str(myjson['authentication']['username'])

    for item in myjson['step_parameters']:
      if item['key']=='source':
        MAIN = str(item['value'])
        INPUT = str(item['value']) + "\\NEWFILES"
        MAIN_FIXER = str(item['value']) + "\\FIXER"
      elif item['key']=='destination':
        DESTINATION = str(item['value']) + "\\OUTPUT"
      elif item['key']=='gmail':
        email = str(item['value'])

    data = process_STPDM(guid, stepname, email, taskid, username, password, MAIN, INPUT, MAIN_FIXER, DESTINATION)
    return Success({'message':'Completed.', 'data': data})
  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )
