#!/usr/bin/python
# FINAL SCRIPT updated as of 25th October 2019
# Workflow - Mediclinic
# Declare Python libraries needed for this script
import json
from flask import current_app as app
from flask import request
from flaskr.response.response import Success, Failure
from flaskr.decorators.login_required import login_required
#from workflow.workflow_process_mediclinic_crm import workflow_process_mediclinic_crm
from workflow.mediclinic_crm import processing_crm
import pandas as pd
import os

@app.route('/process_mediclinic_crm', methods = ['POST'])
@login_required
def process_mediclinic_crm():
  try:
    myjson = json.loads(request.form['json'])
    
    guid = str(myjson['guid'])
    step_name = str(myjson['stepname'])
    email = str(myjson['email'])
    taskid = myjson['taskId']

    for item in myjson['step_parameters']:
      if item['key']=='source':
        source = str(item['value'])
      elif item['key']=='dcl_file':
        dcl_file = str(item['value'])

    num_file = 0
    for file in os.listdir(source):
      if 'mcl' in file.lower():
        mcl_file = os.path.join(source,file)
      else:
        bord_file = os.path.join(source,file)
      num_file = num_file + 1

    #for file in os.listdir(source):
    #  raw_data = pd.read_excel(os.path.join(source,file))
    #  try:
    #    if raw_data.iloc[0]['Unnamed: 0'].lower() == 'corporate':
    #      bord_file = os.path.join(source,file)
    #    else:
    #      mcl_file = os.path.join(source,file)
    #  except Exception as error:
    #    mcl_file = os.path.join(source,file)

    # Call out main function from workflow_process_mediclinic.py
    #data = workflow_process_mediclinic_crm(myjson, guid, stepname, email, taskid, username, password)
     
    data = processing_crm(taskid, guid, step_name, bord_file, mcl_file, dcl_file, email, num_file)

    return Success({'message':'Completed.', 'data': data})
  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )



