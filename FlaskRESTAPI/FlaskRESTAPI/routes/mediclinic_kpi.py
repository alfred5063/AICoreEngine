#!/usr/bin/python
# FINAL SCRIPT updated as of 8th July 2020
# Workflow - MSIG

# Declare Python libraries needed for this script
import json
from flask import current_app as app
from flask import request
from flaskr.response.response import Success, Failure
from flaskr.decorators.login_required import login_required
from workflow.mediclinic_kpi import generate_report
from workflow.mediclinic_kpi import generate_staff_productivity_report
import pandas as pd
from pandas.io.json import json_normalize

@app.route('/mediclinic_kpi', methods = ['POST'])
@login_required
def mediclinic_kpi():
  try:

    myjson = json.loads(request.form['json'])
    #print(myjson)
    guid = str(myjson['guid'])
    stepname = str(myjson['stepname'])
    email = str(myjson['email'])
    taskid = myjson['taskId']
    for item in myjson['step_parameters']:
      if item['key']=='source':
        source = str(item['value'])
      
    jsondf = pd.DataFrame(json_normalize(myjson['execution_parameters_txt']["parameters"]))
    claim_status = jsondf['Claim_Status'].tolist()
    start_visit_date = jsondf['Start_Visit_Date'][0]
    end_visit_date = jsondf['End_Visit_Date'][0]
    start_approval_date = jsondf['Start_Approval_Date'][0]
    end_approval_date = jsondf['End_Approval_Date'][0]
    claim_type = jsondf['Claim_Type'].tolist()
    ## if user choose "ALL" then case_type need to include all the element that frontend have.
    if "All" in jsondf['Case_Type']:
      case_type = ['RS','CS','RC','CC','CD','R','E']
    else:
      case_type = jsondf['Case_Type'].tolist()
    #print(claim_status)
    #print(start_visit_date)
    #print(end_visit_date)
    
    # Call out main function from workflow_process_msig.py
    data = generate_report(taskid, guid, stepname, claim_status, start_visit_date, end_visit_date,claim_type, case_type, start_approval_date, end_approval_date,source=source,email=email)
    return Success({'message':'Completed.', 'data': data})
  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )

@app.route('/staff_productivity', methods = ['POST'])
@login_required
def staff_productivity():
  try:

    myjson = json.loads(request.form['json'])
    #print(myjson)
    guid = str(myjson['guid'])
    stepname = str(myjson['stepname'])
    email = str(myjson['email'])
    taskid = myjson['taskId']
    for item in myjson['step_parameters']:
      if item['key']=='source':
        source = str(item['value'])
      
    jsondf = pd.DataFrame(json_normalize(myjson['execution_parameters_txt']["parameters"]))
    start_approval_date = jsondf['Start_Approval_Date'][0]
    end_approval_date = jsondf['End_Approval_Date'][0]
    
    # Call out main function from workflow_process_msig.py
    data = generate_staff_productivity_report(taskid, guid, stepname, start_approval_date, end_approval_date,source=source,email=email)
    return Success({'message':'Completed.', 'data': data})
  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )
