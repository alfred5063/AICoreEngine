from flask import request, current_app as app
from flaskr.response.response import Success, Failure
from workflow.post_processing_workflow import post_processing_workflow
from workflow.map_action_workflow import map_action_workflow, processing_validation
from flaskr.decorators.login_required import login_required
#new validation method
from workflow.dm_import_marc import import_dm_to_marc
from utils.logging import logging
from utils.Session import session
import pandas as pd
import json
from directory.splitexcelfile import split_excel_file
from directory.get_filename_from_path import get_file_name
from directory.movefile import move_to_archived
import os


#myjson = {'parent_guid': 'bf448534-28c8-4d2e-8158-857dc40a8790', 'guid': '6467609c-37c5-4c91-a859-88c5ff338730', 'authentication': {'username': 'norhasdayany.dawood', 'password': 'MzWPEEG7uo9dtDHcv9PG+w=='}, 'email': 'alfred.simbun@asia-assistance.com', 'stepname': 'Final Validation Processing', 'step_parameters': [{'id': 'tr_prop_1_step_801', 'key': 'source', 'value': '\\\\10.147.78.70\\DM_UAT\\alfred.simbun@asia-assistance.com\\New\\AAN-HNEW-30122020.xlsx', 'description': ''}, {'id': 'tr_prop_2_step_801', 'key': 'destination', 'value': '\\\\10.147.78.70\\DM_UAT\\alfred.simbun@asia-assistance.com\\Result\\', 'description': ''}], 'taskId': '351', 'execution_parameters_att_0': '', 'execution_parameters_txt': '', 'preview': 'False'}

@app.route('/map_action', methods=['POST'])
@login_required
def map_action():
  try:
    myjson = json.loads(request.form['json'])
   
    guid = str(myjson['guid'])
    email = str(myjson['email'])
    taskid = myjson['taskId']
    stepname = str(myjson['stepname'])
    print('Start processing step name: '+ stepname)

    destination = None
    username = None
    password = None

    for item in myjson['step_parameters']:
      if item['key']=='source':
        source = str(item['value'])
      elif item['key']=='destination':
        destination = str(item['value'])
     
    try:
      password = str(myjson['authentication']['password'])
    except Exception as error:
      print("Password continue...")

    try:
      username = str(myjson['authentication']['username'])
    except Exception as error:
      print("Username continue...")

    print('source: {0}'.format(source))
    print('destination: {0}'.format(destination))
    files_to_process = split_excel_file(source, destination, row_limit=500)
    move_to_archived(source)
    head, tail = get_file_name(source)

    for file in files_to_process:
      file_path = "{0}\\{1}"
      source = file_path.format(head, file)
  
      try:
        result_df = processing_validation(taskid, guid, step_name=stepname,source=source, destination=destination, email=email, username=username, password=password)
        print("result_df is....")
        print(result_df)
        
        result, import_source, content = post_processing_workflow(taskid, guid, result_df, step_name=stepname, source=source, destination=destination, email=email)
        print('Import started.')
        result = import_dm_to_marc(taskid, guid, step_name=stepname, source=import_source, destination=destination, email=email, content=content, username=username, password=password)
        print('Import completed.')

      except Exception as error:
        print('Raw records issue; Error = {0}'.format(error))
    return Success({ 'message': 'Post processing complete' })
  except Exception as error:
    return Failure({ 'message': str(error)}, debug=True)

@app.route('/rerun_dm_action', methods=['POST'])
@login_required
def rerun_dm_action():
  try:
    myjson = json.loads(request.form['json'])
   
    guid = str(myjson['guid'])
    email = str(myjson['email'])
    taskid = myjson['taskId']
    stepname = str(myjson['stepname'])
    print('Start processing step name: '+ stepname)
    destination = None
    username = None
    password = None

    for item in myjson['step_parameters']:
      if item['key']=='source':
        source = str(item['value'])
      elif item['key']=='destination':
        destination = str(item['value'])
     
    try:
      password = str(myjson['authentication']['password'])
    except Exception as error:
      print("Password continue...")

    try:
      username = str(myjson['authentication']['username'])
    except Exception as error:
      print("Username continue...")

    print('source: {0}'.format(source))
    print('destination: {0}'.format(destination))
    #files_to_process = split_excel_file(source, destination, row_limit=200)
    #rpa_files = '{0}\\{1}\\New\\'.format(source, email)
    raw_source = source + '\\{0}\\New\\'.format(email)
    non_hidden_files = [filename for filename in os.listdir(raw_source) if not filename.startswith('~$') and any(map(filename.endswith, ('.xls', '.xlsx')))]
    result_list = []
    for i, file in enumerate(non_hidden_files):
      source = ""
      source = os.path.join(raw_source, file)
      print('For loop source: {0}'.format(source))
      try:
        result_df = processing_validation(taskid, guid, step_name=stepname,source=source, destination=destination, email=email, username=username, password=password)
        #result_df.replace('nan','', inplace=True)
        #result.rename({"action code":"action_code", "marc":"fields_update", "custom search":"search_criteria"}, axis=1, inplace=True)
        result = post_processing_workflow(taskid, guid, result_df, step_name=stepname, source=source, destination=destination, email=email)
      except Exception as error:
        print('Raw records issue; Error = {0}'.format(error))

    return Success({ 'message': 'Post processing complete' })
  except Exception as error:
    return Failure({ 'message': str(error)}, debug=True)

@app.route('/dm_import_marc', methods=['POST'])
@login_required
def dm_import_marc():
  try:
    myjson = json.loads(request.form['json'])
   
    guid = str(myjson['guid'])
    email = str(myjson['email'])
    taskid = myjson['taskId']
    stepname = str(myjson['stepname'])
    print('Start processing step name: '+ stepname)
    destination = None
    username = None
    password = None

    for item in myjson['step_parameters']:
      if item['key']=='source':
        source = str(item['value'])
      elif item['key']=='destination':
        destination = str(item['value'])
     
    try:
      password = str(myjson['authentication']['password'])
    except Exception as error:
      print("Password continue...")

    try:
      username = str(myjson['authentication']['username'])
    except Exception as error:
      print("Username continue...")

    print('source: {0}'.format(source))
    print('destination: {0}'.format(destination))
    print('Import started.')
    result = import_dm_to_marc(taskid, guid, step_name=stepname, source=source, destination=destination, email=email, username=username, password=password)
    print('Import completed.')
    return Success({'message':'Complete import to marc'})
  except Exception as error:
    return Failure({'message':str(error)}, debug=True)

@app.route('/post_processing', methods=['POST'])
@login_required
def post_processing():
  try:
    myjson = json.loads(request.form['json'])
   
    guid = str(myjson['guid'])
    email = str(myjson['email'])
    taskid = myjson['taskId']
    stepname = str(myjson['stepname'])
    print('Start processing step name: '+ stepname)
    source = None
    destination = None
    for item in myjson['step_parameters']:
      if item['key']=='source':
        source = str(item['value'])
      elif item['key']=='destination':
        destination = str(item['value'])
     
    result = post_processing_workflow(taskid, guid, step_name=stepname, source=source, destination=destination, email=email)
    return Success({ 'message': 'Post processing complete' })
  except Exception as error:
    return Failure({ 'message': str(error)}, debug=True)
