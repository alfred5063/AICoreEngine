#!/usr/bin/python
# FINAL SCRIPT updated as of 20th July 2020
# Workflow - REQ/DCA

# Declare Python libraries needed for this script
import pandas as pd
import json as json
import requests
import os, time
from loading.excel.checkExcelLoading import *
from datetime import datetime as dt
from connector.connector import *
from pandas.io.json import json_normalize
from utils.Session import session
from utils.guid import get_guid
from utils.audit_trail import audit_log
from utils.notification import send
from utils.logging import logging
from transformation.reqdca_generate_preview import extract_info
from transformation.reqdca_edit_doc import update_doc
from transformation.reqdca_update_marc import update_in_marc
from transformation.reqdca_update_cm import *
from transformation.reqdca_update_dcm import *
from transformation.reqdca_excel_path_mapping import *
from transformation.reqdca_checkvalidity import check_validity
from extraction.marc.authenticate_marc_access import get_marc_access
from transformation.reqdca_update_cmf import update_cmf

#myjson =  {'parent_guid': 'b3bb7843-4bae-45bb-9168-b2010e347927', 'guid': '55822366-b2cb-468a-b876-00fb3e2a32d2',
#           'authentication': {'username': 'alfred.simbun', 'password': 'XqtGVHqBw39b/y1i1UqpDg=='},
#           'email': 'nurfazreen.ramly@asia-assistance.com',
#           'stepname': 'DCA', 'step_parameters': [{'id': 'tr_prop_1_step_1', 'key': 'docType',
#                                                   'value': 'DCA', 'description': 'DCA'},
#                                                  {'id': 'tr_prop_2_step_1', 'key': 'dcm_path', 'value': '\\\\dtisvr2\\CBA_UAT\\Result\\REQDCA_Paths\\1. DISBURSEMENT CLAIMS INVOICE LOG FILE', 'description': 'Path to DCM files'},
#                                                  {'id': 'tr_prop_3_step_1', 'key': 'cml_path', 'value': '\\\\dtisvr2\\CBA_UAT\\Result\\REQDCA_Paths\\MASTER', 'description': 'Path to CML files'}],
#           'taskId': '223', 'execution_parameters_att_0': '', 'execution_parameters_txt':
#           {'docType': 'REQ', 'parameters': [{'Case': '344169',
#                                              'Amount': '-299',
#                                              'Reason': 'Over Billing',
#                                              'Remarks': 'Cashless',
#                                              'Type': 'Admission'}]},
#                                              'preview': 'False'}

#document = str(myjson['execution_parameters_txt']['docType'])
#password = ''
#username = str(myjson['authentication']['username'])
#guid = str(myjson['guid'])
#stepname = str(myjson['stepname'])
#email = str(myjson['email'])
#taskid = str(myjson['taskId'])
#preview = str(myjson['preview'])
#for item in myjson['step_parameters']:
#  if item['key']=='dcm_path':
#    dcm_path = str(item['value'])
#  elif item['key']=='cml_path':
#    cml_path = str(item['value'])

def update_dcmfile(newmain_df, document, doc_template, email, DCMfile, reqdca_base):
  try:
    print("- Check if file is not locked.")
    filestatus = is_locked(DCMfile)
    if str(filestatus) == 'False':
      update_dcm(newmain_df, document, doc_template, email, DCMfile, reqdca_base)
      print("- Updating DCM document completed.")
      audit_log("REQ/DCA - Updating DCM document completed.", "Completed...", reqdca_base)
      send(reqdca_base, email, "REQDCA - Updating DCM document completed.", "Hi, <br/><br/><br/>DCM document has been updated. Please check. <br/>Reference Number: " + str(reqdca_base.guid) + "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
      audit_log("REQ/DCA - User has been informed through email.", "Completed...", reqdca_base)
      print("- User has been informed through email.")
      return 1
    else:
      print("- Updating DCM document cannot be done. Trying to open locked file.")
      curtask = pd.DataFrame(os.popen('tasklist').readlines()).to_json(orient = 'index')
      if "excel.exe" in curtask:
        os.system("taskkill /f /im  excel.exe")
        update_dcm(newmain_df, document, doc_template, email, DCMfile, reqdca_base)
        print("- Updating DCM document completed.")
        audit_log("REQ/DCA - Updating DCM document completed.", "Completed...", reqdca_base)
        send(reqdca_base, email, "REQDCA - Updating DCM document completed.", "Hi, <br/><br/><br/>DCM document has been updated. Please check. <br/>Reference Number: " + str(reqdca_base.guid) + "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
        audit_log("REQ/DCA - User has been informed through email.", "Completed...", reqdca_base)
        print("- User has been informed through email.")
        return 1
      else:
        print("- Updating DCM document: %s cannot be done. File is currently still LOCKED / USED." % DCMfile)
        audit_log("REQ/DCA - Updating DCM document: %s cannot be done. File is currently still LOCKED / USED." % DCMfile, "Completed...", reqdca_base)
        send(reqdca_base, reqdca_base.email, "REQ/DCA - Updating DCM document: %s cannot be done. File is currently still LOCKED / USED." % DCMfile, "Hi, <br/><br/><br/>. <br/>Reference Number: " + str(reqdca_base.guid)+"<br/><br/><br/>Regards,<br/>Robotic Process Automation")
        audit_log("REQ/DCA - User has been informed through email.", "Completed...", reqdca_base)
        print("- User has been informed through email.")
        return 1
  except Exception as error:
    print(error)



def process_REQDCA(myjson, document, guid, stepname, email, taskid, preview, username, password, dcm_path, cml_path):

  
  current_year = dt.now().year
  #for testing
  #current_year = '2020'

  if document == 'REQ':
    doc_template = 'CBA-OPS Disbursement'
    DCMfile = dcm_path + "\%s" % str(current_year) + "\Ops disbursement_REQ_%s.xlsx" % str(current_year)
  elif document == 'DCA':
    doc_template = 'CBA-Disbursement Claims Adjustment'
    DCMfile = dcm_path + "\%s" % str(current_year) + "\Disbursement Claims Master %s.xlsx" % str(current_year)

  reqdca_base = session.base(taskid, guid, email, stepname, username = username, password = password)

  if password != '':
    reqdca_base.set_password(get_marc_access(reqdca_base))
  else:
    reqdca_base.password = ''

  try:

    jsondf = pd.DataFrame(json_normalize(myjson['execution_parameters_txt']['parameters']))
    jsondf = jsondf.rename(columns = {'Case': 'Case ID'})
    audit_log("REQ/DCA - Get JSON string from front-end and process Case ID, Amount and Reason.", "Completed...", reqdca_base)

    if preview == 'True':
      jsondf_val = check_validity(jsondf, reqdca_base, document)
      #for testing
      #jsondf_val = maindf_details
      jsondf_invalid = jsondf_val.loc[jsondf_val['Validity'] == "Valid Case ID"].set_index('Case ID').reset_index()

      if jsondf_invalid.empty != True:

        dcamain_df = extract_info(jsondf_val, reqdca_base, document, DCMfile)
        dcamain_df = dcamain_df.drop_duplicates()
        print(dcamain_df)
        prereqdca_result = dcamain_df.to_json(orient = 'index')
      else:
        i = 0
        invalidid = jsondf_val['Case ID']
        invalid_df = pd.DataFrame()
        for i in range(len(invalidid)):
          invalid_df.loc[i, 'Case ID'] = "Case ID [ %s ] cannot be found in MARC. Please check." % invalidid[i]
          invalid_df.loc[i, 'Amount'] = int(jsondf_val['Amount'][i])
          invalid_df.loc[i, 'Remarks'] = jsondf_val['Remarks'][i]
          invalid_df.loc[i, 'Type'] = jsondf_val['Type'][i]
          invalid_df.loc[i, 'Reason'] = jsondf_val['Reason'][i]
          invalid_df.loc[i, 'Validity'] = jsondf_val['Validity'][i]
          invalid_df.loc[i, 'Validity'] = jsondf_val['Client'][i]
          invalid_df.loc[i, 'Validity'] = jsondf_val['Arrangement'][i]
          prereqdca_result = invalid_df.to_json(orient = 'index')
      audit_log("REQ/DCA - Convert dataframe to JSON string and parse to front-end.", "Completed...", reqdca_base)
      print(prereqdca_result)
      return prereqdca_result
    elif preview == 'False':
      jsondf_val = check_validity(jsondf, reqdca_base, document)
      print(jsondf_val)
      jsondf_invalid = jsondf_val.loc[jsondf_val['Validity'] == "Valid Case ID"].set_index('Case ID').reset_index()
      if jsondf_invalid.empty != True:

        dcamain_df = extract_info(jsondf_val, reqdca_base, document, DCMfile)

        dcamain_df = dcamain_df.loc[dcamain_df['Validity'] == "Valid Case ID"].set_index('Case ID').reset_index()
        audit_log("REQ/DCA - Preparing dataframe with valid Case IDs.", "Completed...", reqdca_base)

        if document == 'DCA':
          try:
            maindf_cm = update_cmf(dcamain_df, current_year, cml_path)
            audit_log("REQ/DCA - Updating file paths for accessing and processing Client Master Files.", "Completed...", reqdca_base)
          except Exception as error:
            logging("REQDCA - DCA - Updating file paths.", error, reqdca_base)
            pass

          # Step 4 - Process the ones with paths
          try:
            withpath_df = maindf_cm.loc[maindf_cm['Path'] != "Path not found"]
            rownumwithpath = withpath_df['Case ID'].count()
            if rownumwithpath > 0:
              newmain_df = update_cm(withpath_df, reqdca_base)
            else:
              rownumwithpath = rownumwithpath
          except Exception as error:
            logging("REQDCA - DCA - Process the ones with paths.", error, reqdca_base)
            pass

          # Step 5 - Process the ones without paths
          try:
            nopath_df = maindf_cm.loc[maindf_cm['Path'] == "Path not found"]
            rownumnopath = nopath_df['Case ID'].count()
            if rownumnopath > 0:
              d = 0
              nopath_df = nopath_df.reset_index().drop(columns = ['index'])
              for d in range(len(nopath_df['Case ID'])):
                audit_log("REQ/DCA - Case ID [ %s ] cannot be found in any CML file for the current Year [ %s ]. Please check." % nopath_df['Case ID'][d], current_year, "Completed...", reqdca_base)
            else:
              audit_log("REQ/DCA - All Case IDs can be found in associated CML filepath.", "Completed...", reqdca_base)
          except Exception as error:
            logging("REQDCA - DCA - Process the ones without paths.", error, reqdca_base)
            pass

          if rownumwithpath != 0:
            print("Step 6 - Prepare dataframe to update DCM and MARC")
            try:
              dcamain_df = dcamain_df.astype(object)
              newmain_df = newmain_df.astype(object)
              newmain_df = pd.merge(left = dcamain_df, right = newmain_df, left_on = dcamain_df['Member Ref. ID'], right_on = newmain_df['Member Ref. ID'])
              newmain_df = newmain_df.drop(columns = ['key_0', 'Amount', 'Sub Case ID_y', 'Member Ref. ID_y', 'Type_y', 'Remarks_y', 'OB Registered Date_y', 'Remarks_y'])
              newmain_df = newmain_df.drop_duplicates(subset=['Case ID', 'Disbursement Claims', 'Total', 'Sub Case ID_x'])
            except Exception as error:
              logging("REQDCA - DCA - Prepare dataframe to update DCM and MARC.", error, reqdca_base)
              pass

            # Step 7 - Updating Disbursement Client Master files.
            try:
              update_dcmfile(newmain_df, document, doc_template, email, DCMfile, reqdca_base)
            except Exception as error:
              logging("REQDCA - DCA - Updating Disbursement Client Master files.", error, reqdca_base)
              pass

            # Step 8 - Created document and updated in SQL database"
            try:
              update_doc(newmain_df, doc_template, email, reqdca_base, document)
            except Exception as error:
              logging("REQDCA - DCA - Created document and updated in SQL database.", error, reqdca_base)
              pass

            # Step 9 - Created document and updated MARC."
            try:
              update_in_marc(newmain_df, doc_template, email, reqdca_base, document)
            except Exception as error:
              logging("REQDCA - DCA - Created document and updated MARC..", error, reqdca_base)
              pass
          else:
            audit_log("REQDCA - DCA - CML filepath cannot be found based on Case ID. No CML file has been updated.", "Alert...", reqdca_base)
            audit_log("REQDCA - DCA - Disbursement Client Master is not updated.", "Alert...", reqdca_base)
            audit_log("REQDCA - DCA - MARC document is not updated.", "Alert...", reqdca_base)
            pass

          # Step 10 - Sending notification through email to the user."
          try:
            if(email != None):
              send(reqdca_base, email, "REQ/DCA - RPA Task Execution Completed.", "Hi, <br/><br/><br/>You task has been completed. <br/>Reference Number: " + str(reqdca_base.guid) + "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
            audit_log("REQ/DCA - RPA task execution completion notification. Sending notification through email to the user.", "Completed...", reqdca_base)
          except Exception as error:
            logging("REQDCA - REQ - Sending notification through email to the user.", error, reqdca_base)
            pass

        else:
          # Document is REQ.
          # Step 3 - Prepare dataframe to update DCM and MARC
          try:
            newmain_df = dcamain_df
          except Exception as error:
            logging("REQDCA - REQ - Prepare dataframe to update DCM and MARC.", error, reqdca_base)
            pass

          # Step 4 - Updating Disbursement Client Master files."
          try:
            update_dcmfile(newmain_df, document, doc_template, email, DCMfile, reqdca_base)
          except Exception as error:
            logging("REQDCA - REQ - Updating Disbursement Client Master files.", error, reqdca_base)
            pass

          # Step 5 - Created document and updated in SQL database."
          try:
            update_doc(newmain_df, doc_template, email, reqdca_base, document)
          except Exception as error:
            logging("REQDCA - REQ - Created document and updated in SQL database.", error, reqdca_base)
            pass

          # Step 6 - Created document and updated MARC."
          try:
            update_in_marc(newmain_df, doc_template, email, reqdca_base, document)
          except Exception as error:
            logging("REQDCA - REQ - Created document and updated MARC.", error, reqdca_base)
            pass

          # Step 7 - Sending notification through email to the user."
          try:
            if(email != None):
              send(reqdca_base, email, "REQ/DCA - RPA Task Execution Completed.", "Hi, <br/><br/><br/>You task has been completed. <br/>Reference Number: " + str(reqdca_base.guid) + "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
              audit_log("REQ/DCA - RPA task execution completion notification. Sending notification through email to the user.", "Completed...", reqdca_base)
            else:
              audit_log("REQ/DCA - RPA task execution completion notification. No email was provided for this workflow. Notification is not sent.", "Completed...", reqdca_base)
          except Exception as error:
            logging("REQDCA - REQ - Sending notification through email to the user.", error, reqdca_base)
            pass
      
        audit_log(stepname, "Completed", reqdca_base)

      else:
        try:
          if(email != None):
            send(reqdca_base, email, "REQ/DCA - RPA Task Execution Stopped.", "Hi, <br/><br/><br/>You task has been stopped. This due to non existance of your Case ID in MARC.<br/>Reference Number: " + str(reqdca_base.guid) + "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
            audit_log("REQ/DCA - RPA task execution termination notification. Sending notification through email to the user.", "Completed...", reqdca_base)
          else:
            audit_log("REQ/DCA - RPA task execution termination notification. No email was provided for this workflow. Notification is not sent.", "Completed...", reqdca_base)
        except Exception as error:
          logging("REQ/DCA - Sending task termination notification through email to the user.", error, reqdca_base)
          pass

    print("done")
    return "Completed"

  except Exception as error:
    logging("REQDCA - process_REQDCA", error, reqdca_base)

