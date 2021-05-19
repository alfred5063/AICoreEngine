#!/usr/bin/python
# FINAL SCRIPT updated as of 8th July 2020
# Workflow - REQ/DCA

# Declare Python libraries needed for this script
import json
from flask import current_app as app
from flask import request
from flaskr.response.response import Success, Failure
from flaskr.decorators.login_required import login_required
from workflow.workflow_process_reqdca import process_REQDCA

@app.route('/process_reqdca', methods = ['POST'])
@login_required
def process_reqdca():
  try:

    myjson = json.loads(request.form['json'])
    print(myjson)
    document = str(myjson['execution_parameters_txt']['docType'])
    password = str(myjson['authentication']['password'])
    username = str(myjson['authentication']['username'])
    guid = str(myjson['guid'])
    stepname = str(myjson['stepname'])
    email = str(myjson['email'])
    taskid = str(myjson['taskId'])
    preview = str(myjson['preview'])
    for item in myjson['step_parameters']:
      if item['key']=='dcm_path':
        dcm_path = str(item['value'])
      elif item['key']=='cml_path':
        cml_path = str(item['value'])

    # Call out main function from workflow_process_req.py
    reqdca = process_REQDCA(myjson, document, guid, stepname, email, taskid, preview, username, password, dcm_path, cml_path)

    return Success({'message' : 'Completed.', 'data' : reqdca})
  except Exception as error:
    return Failure({'message': str(error) }, debug = True)

