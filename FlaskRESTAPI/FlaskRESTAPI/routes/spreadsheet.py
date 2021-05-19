#!/usr/bin/python
# FINAL SCRIPT updated as of 29th April 2020

# Declare Python libraries needed for this script
import json
from flask import current_app as app
from flask import request
from flaskr.response.response import Success, Failure
from flaskr.decorators.login_required import login_required

@app.route('/preprocess_spreadsheet', methods = ['POST'])
@login_required
def preprocess_spreadsheet():
  try:

    myjson = json.loads(request.form['json'])
    guid = str(myjson['guid'])
    stepname = str(myjson['stepname'])
    email = str(myjson['email'])
    taskid = myjson['taskId']
    source = ''
    destination = ''
    for item in myjson['step_parameters']:
      if item['key']=='source':
        source = str(item['value'])
      elif item['key']=='destination':
        predes = source.split("New")
        destination = str(predes[0] + "New\\" + email + "\\Result\\")

    # Call out main function from workflow_process_spreadsheet.py
    #process_spreadsheet(source, destination, guid, stepname, email, taskid)

    return Success({'message':'Completed.'})
  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )
