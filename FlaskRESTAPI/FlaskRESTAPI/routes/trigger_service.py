#!/usr/bin/python
# FINAL SCRIPT updated as of 15th March 2021
# Workflow - Trigger Service

# Declare Python libraries needed for this script
import json
from flask import current_app as app
from flask import request
from flaskr.response.response import Success, Failure
from flaskr.decorators.login_required import login_required
from workflow.workflow_process_services import trigger_service

@app.route('/process_service', methods = ['POST'])
@login_required
def process_service():
  try:

    myjson = json.loads(request.form['json'])
    password = str(myjson['authentication']['password'])
    username = str(myjson['authentication']['username'])
    guid = str(myjson['guid'])
    stepname = str(myjson['stepname'])
    email = str(myjson['email'])
    taskid = myjson['taskId']

    # Call out main function from workflow_process_msig.py
    data = process_MSIG(myjson, guid, stepname, email, taskid, preview, username, password)

    return Success({'message':'Completed.', 'data': data})
  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )
