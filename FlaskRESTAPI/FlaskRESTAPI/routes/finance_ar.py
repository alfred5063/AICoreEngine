#!/usr/bin/python
# FINAL SCRIPT updated as of 19th May 2020
# Final API for Finance AR

# Declare Python libraries needed for this script
import json
from flask import current_app as app
from flask import request
from flaskr.response.response import Success, Failure
from flaskr.decorators.login_required import login_required
from workflow.workflow_process_finance_ar import process_Finance_AR

@app.route('/process_finance_ar', methods = ['POST'])
@login_required
def process_finance_ar():
  try:
    myjson = json.loads(request.form['json'])
    guid = str(myjson['guid'])
    stepname = str(myjson['stepname'])
    email = str(myjson['email'])
    taskid = myjson['taskId']

    # Call out main function from workflow_process_finance_ar.py
    process_Finance_AR(myjson, guid, stepname, email, taskid)

    return Success({'message':'Completed.'})
  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )
