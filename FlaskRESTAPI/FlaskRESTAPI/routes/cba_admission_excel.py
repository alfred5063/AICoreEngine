#!/usr/bin/python
# FINAL SCRIPT updated as of 3rd January 2020
# Workflow - ADM
# Declare Python libraries needed for this script
import json
from flask import current_app as app
from flask import request
from flaskr.response.response import Success, Failure
from flaskr.decorators.login_required import login_required
from workflow.workflow_process_admission_excel import workflow_process_admission_excel

@app.route('/process_admission_excel', methods = ['POST'])
@login_required
def process_admission_excel():
  try:

    myjson = json.loads(request.form['json'])
    password = str(myjson['authentication']['password'])
    username = str(myjson['authentication']['username'])
    guid = str(myjson['guid'])
    stepname = str(myjson['stepname'])
    email = str(myjson['email'])
    taskid = myjson['taskId']
    
    # Call out main function from workflow_process_msig.py
    data = workflow_process_admission_excel(myjson, guid, stepname, email, taskid, username, password)

    return Success({'message':'Completed.', 'data': data})
  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )


