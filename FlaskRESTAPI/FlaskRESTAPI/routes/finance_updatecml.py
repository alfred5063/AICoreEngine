#!/usr/bin/python
# FINAL SCRIPT updated as of 16th June 2020
# Workflow - Finance 
# Version 1

# Declare Python libraries needed for this script
import json
from flask import current_app as app
from flask import request
from flaskr.response.response import Success, Failure
from flaskr.decorators.login_required import login_required
from workflow.workflow_process_finance_cml import process_Finance_CML

@app.route('/process_financecmlupdate', methods = ['POST'])
@login_required
def process_financecmlupdate():
  try:

    myjson = json.loads(request.form['json'])
    guid = str(myjson['guid'])
    stepname = str(myjson['stepname'])
    email = str(myjson['email'])
    taskid = myjson['taskId']

    print(myjson)

    # Call out main function from workflow_process_finance_cml.py
    process_Finance_CML(myjson, guid, stepname, email, taskid)

    return Success({'message':'Completed.'})
  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )
