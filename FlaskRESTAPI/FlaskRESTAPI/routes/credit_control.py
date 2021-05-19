#!/usr/bin/python
# FINAL SCRIPT updated as of 27th May 2020
# Workflow - Finance Credit Control

# Declare Python libraries needed for this script
import json
from flask import current_app as app
from flask import request
from flaskr.response.response import Success, Failure
from flaskr.decorators.login_required import login_required
from workflow.workflow_credit_control import process_credit_control

@app.route('/process_creditcont', methods = ['POST'])
@login_required
def process_creditcont():
  try:
    myjson = json.loads(request.form['json'])
    guid = str(myjson['guid'])
    stepname = str(myjson['stepname'])
    email = str(myjson['email'])
    taskid = myjson['taskId']
    username = myjson['authentication']['username']
    password = myjson['authentication']['password']

    # Call out main function from workflow_process_finance_ar.py
    process_credit_control(myjson, guid, stepname, email, taskid, username, password)

    return Success({'message':'Completed.'})
  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )
