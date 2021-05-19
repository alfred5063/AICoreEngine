#!/usr/bin/python
# FINAL SCRIPT updated as of 8th July 2020
# Workflow - MSIG

# Declare Python libraries needed for this script
import json
from flask import current_app as app
from flask import request
from flaskr.response.response import Success, Failure
from flaskr.decorators.login_required import login_required
from workflow.workflow_process_msig import process_MSIG

@app.route('/process_msig', methods = ['POST'])
@login_required
def process_msig():
  try:

    myjson = json.loads(request.form['json'])
    password = str(myjson['authentication']['password'])
    username = str(myjson['authentication']['username'])
    guid = str(myjson['guid'])
    stepname = str(myjson['stepname'])
    email = str(myjson['email'])
    taskid = myjson['taskId']
    preview = str(myjson['preview'])

    for item in myjson['step_parameters']:
      if item['key']=='textfile_path':
        destination = str(item['value'])

    # Call out main function from workflow_process_msig.py
    data = process_MSIG(myjson, destination, guid, stepname, email, taskid, preview, username, password)

    return Success({'message':'Completed.', 'data': data})
  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )
