#!/usr/bin/python
# FINAL SCRIPT updated as of 29th April 2020
# Workflow - Data Migration Req

# Declare Python libraries needed for this script
import json
from flask import current_app as app
from flask import request
from flaskr.response.response import Success, Failure
from flaskr.decorators.login_required import login_required
from workflow.workflow_process_data_migration_req2 import process_data_migration_REQ_incremental_load

@app.route('/process_data_migration_req', methods = ['POST'])
@login_required
def process_data_migration_req():
  try:
    myjson = json.loads(request.form['json'])
    print(myjson)
    guid = str(myjson['guid'])
    stepname = str(myjson['stepname'])
    password = str(myjson['authentication']['password'])
    username = str(myjson['authentication']['username'])
    email = str(myjson['email'])
    taskid = str(myjson['taskId'])
    preview = str(myjson['preview'])
    year = int(myjson['execution_parameters_txt']['parameters'][0]['year'])

    for item in myjson['step_parameters']:
      if item['key']=='req_path':
        migration_folder = str(item['value'])


		# Call out main function from process_data_migration_req.py
    data = process_data_migration_REQ_incremental_load(migration_folder, year, email)

    return Success({'message':'Completed.', 'data': data})
  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )
