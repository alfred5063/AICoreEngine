#!/usr/bin/python
# FINAL SCRIPT updated as of 25th October 2019
# Workflow - Mediclinic
# Declare Python libraries needed for this script
import json
from flask import current_app as app
from flask import request
from flaskr.response.response import Success, Failure
from flaskr.decorators.login_required import login_required
from workflow.workflow_process_mediclinic import workflow_process_mediclinic

@app.route('/process_mediclinic', methods = ['POST'])
@login_required
def process_mediclinic():
  try:
    myjson = json.loads(request.form['json'])
    guid = str(myjson['guid'])
    stepname = str(myjson['stepname'])
    email = str(myjson['email'])
    taskid = myjson['taskId']
    # Call out main function from workflow_process_mediclinic.py
    data = process_mediclinic(myjson, guid, stepname, email, taskid, password, username, source, destination)

    return Success({'message':'Completed.', 'data': data})
  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )


