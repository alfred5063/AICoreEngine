#!/usr/bin/python
# FINAL SCRIPT updated as of 14th July 2020
# Final API for Finance AR

# Declare Python libraries needed for this script
import json
from flask import current_app as app
from flask import request
from flaskr.response.response import Success, Failure
from flaskr.decorators.login_required import login_required
from workflow.finance_sage_to_max import process_finance_sage_to_max

@app.route('/process_max_update', methods = ['POST'])
@login_required
def process_max_update():
  try:
    myjson = json.loads(request.form['json'])
    guid = str(myjson['guid'])
    stepname = str(myjson['stepname'])
    email = str(myjson['email'])
    taskid = myjson['taskId']
    password = str(myjson['authentication']['password'])
    username = str(myjson['authentication']['username'])

    for item in myjson['step_parameters']:
      if item['key']=='source':
        source = str(item['value'])
      elif item['key']=='destination':
        destination = str(item['value'])

    # Call out main function
    process_finance_sage_to_max(guid, stepname, email, taskid, username, password, source, destination)

    return Success({'message':'Completed.'})
  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )
