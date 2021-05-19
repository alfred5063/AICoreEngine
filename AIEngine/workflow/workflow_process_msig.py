#!/usr/bin/python
# FINAL SCRIPT updated as of 8th July 2020
# Workflow - MSIG

# Declare Python libraries needed for this script
import pandas as pd
import json as json
import requests
from utils.notification import send
from pandas.io.json import json_normalize
from transformation.process_refnum import process_refnum
from utils.Session import session
from utils.guid import get_guid
from utils.audit_trail import audit_log
from utils.logging import logging
from utils.notification import send
from automagic.auto_updaterefnum import *
from connector.connector import MySqlConnector
from extraction.marc.authenticate_marc_access import get_marc_access

#myjson = {'parent_guid': '7065de5b-f11d-4799-af27-eb5608f31240', 'guid': 'e6da280a-1e58-4116-9504-4078ff5ac711', 'authentication': {'username': 'alfred.simbun', 'password': 'Yt8PKOiv/E5XcLCN3hLISw=='}, 'email': 'nurfazreen.ramly@asia-assistance.com', 'stepname': 'MSIG', 'step_parameters': [{'id': 'tr_prop_1_step_1', 'key': 'Adm No', 'value': 'read_excel', 'description': 'extract info from excel'}, {'id': 'tr_prop_2_step_1', 'key': 'textfile_path', 'value': '\\\\dtisvr2\\CBA_UAT\\Result\\REQDCA_Paths\\Update status', 'description': 'path to text file'}], 'taskId': '247', 'execution_parameters_att_0': [{'Adm. No': 369748, 'CASE ID': 'CL1907002426', 'CH': 'MYCLGDE'}, {'Adm. No': 370677, 'CASE ID': 'CL1907002427', 'CH': 'MYCLMNS'}, {'Adm. No': 370773, 'CASE ID': 'CL1907002428', 'CH': 'MYCLSSU'}, {'Adm. No': 371991, 'CASE ID': 'URC', 'CH': 'N/A'}, {'Adm. No': 372333, 'CASE ID': 'CL1907002429', 'CH': 'MYCLNAR'}, {'Adm. No': 372429, 'CASE ID': 'CL1907002430', 'CH': 'MYCLSPY'}, {'Adm. No': 372443, 'CASE ID': 'CL1907002431', 'CH': 'MYCLGDE'}], 'execution_parameters_txt': '', 'preview': 'False'}
#password = str(myjson['authentication']['password'])
#username = str(myjson['authentication']['username'])
#guid = str(myjson['guid'])
#stepname = str(myjson['stepname'])
#email = str(myjson['email'])
#taskid = myjson['taskId']
#preview = str(myjson['preview'])

#for item in myjson['step_parameters']:
#      if item['key']=='textfile_path':
#        destination = str(item['value'])


#username = ''
#password = ''

# Workflow function
def process_MSIG(myjson, destination, guid, stepname, email, taskid, preview, username, password):

  try:

    # Capture the current login session
    msig_base = session.base(taskid, guid, email, stepname, username = username, password = password)

    # Password encryption
    if password != "":
      msig_base.set_password(get_marc_access(msig_base))
    else:
      msig_base.password = ''
    # Step 1 - Get JSON string from front-end and process Adm. No, Case ID and CH columns
    jsondf = pd.DataFrame(json_normalize(myjson['execution_parameters_att_0']))
    jsondf = jsondf.dropna()
    jsondf = jsondf[pd.notnull(jsondf['Adm. No'])]
    main_df = process_refnum(jsondf, msig_base, preview)
    audit_log("MSIG - Get JSON string from front-end and process Adm. No, Case ID and CH columns.", "Completed...", msig_base)
    #pd.set_option('max_columns',None)
    # Check if a review from user is needed
    if preview == 'True':

      audit_log("MSIG - Review checkbox was selected to run this process.", "Completed...", msig_base)

      # Step 2 - Convert dataframe to JSON string and parse to front-end
      premsig_result = json.dumps(main_df.to_dict(orient = 'index'))
      audit_log("MSIG - Convert dataframe to JSON string and parse to front-end for preview.", "Completed...", msig_base)

      # Return the results for user's review
      return premsig_result
      audit_log("MSIG - JSON string parsed back to front-end for preview.", "Completed...", msig_base)

    elif preview == 'False':

      audit_log("MSIG - Review checkbox was not selected to run this process.", "Completed...", msig_base)

      # Step 3 - Updating client's Reference Number in MARC using RPA
      refnum_result = initBrowser(main_df, msig_base)
      current_time = time.strftime("%Y%m%d_%H%M%S")
      textfilename = '\\refnum_update_status_{0}.txt'.format(current_time)
      refnum_text = refnum_result.to_csv(destination+textfilename, header=True, index=False, sep='|', mode='a')
      result_removenull = refnum_result.fillna('')
      #successful_list = result_removenull["Updated Case ID"].to_list()
      #unsuccessful_list = result_removenull["Not Updated Case ID"].to_list()
      if any(a != '' for a in unsuccessful_list) == True:
        unsuccessful_list = list(filter(('').__ne__, unsuccessful_list))
      else:
        unsuccessful_list = "None"
      audit_log("MSIG - Updating client's Reference Number in MARC using RPA.", "Completed...", msig_base)

      # Summary to display the final result for this workflow
      print("- Provide workflow summary")
      audit_log("MSIG Summary", "Completed...", msig_base)

      # Summary displayed on the front-end
      if(email != None):
        send(msig_base, email, "RPA Task Execution Completed.",
           "Hi, <br/><br/><br/>Your task has been completed.<br/>Reference Number: " + str(guid) +
           "<br/>" +
           "<br/>Run Summary can be found through the following link: <" + str(destination) + "><br/><br/>" +
           "<br/>Regards,<br/>Robotic Process Automation")
        
      audit_log("MSIG - Email notification sent to user", "Completed...", msig_base)
      return "Completed"

  except Exception as error:
    logging("MSIG - process_MSIG", error, msig_base)
    print(error)
