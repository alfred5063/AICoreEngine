#!/usr/bin/python
# FINAL SCRIPT updated as of 15th April 2021

# Declare Python libraries needed for this script
from utils.Session import session
from utils.audit_trail import audit_log, update_audit, audit_logs_insert, cs_audit_log
from output.get_action_codes_from_db import get_action_codes_from_db,convert_action_code_to_actions
from selenium_interface.accessMarc import update_marc, login_marc, logout_marc
from utils.notification import send
from utils.logging import *
from output.save_output import save_process_output
from extraction.rpa.get_rpa_db import get_process_dm_validation, get_process_dm_validation_bykey
from extraction.excel.readexcel import readExcel
import pandas as pd
from transformation.df_preprocessing import preprocessing, fix_dates
from extraction.marc.query_marc_df import sp_query_as_df, query_as_df
from extraction.marc.marc_query_dict import query_by_policy_num, get_uniqueue_list_from_multiple_columns, parameter_mapper, get_member_by_NRIC, sp_get_member_by_NRIC, query_principle_by_policy_num,excel_db_mapper
import numpy as np
import math
from datetime import datetime
from loading.excel.createexcel import create_excel
import json
from automagic.marc import *
from extraction.marc.authenticate_marc_access import get_marc_access
from regex import sub, findall
from transformation.dm_generate_matrix_code import generate_matrix_code_by_row
from analytics.marc_analytics import analytic_dm_by_row, analytic_dm, analytic_dm_record

#@iterate_source
def map_action_workflow(taskid, guid, result_df, step_name, source=None, destination=None, email=None, username=None, password=None, debug=False):

  dm_base = session.base(taskid, guid,email,step_name,username=username, password=password, filename=source)
  dm_properties = session.data_management(source, destination)
  print('source: {0}; destination: {1}'.format(dm_properties.source, dm_properties.destination))
  audit_logs = []
  audit_logs.append(cs_audit_log.audit_log('Selenium Result',result, dm_base))

  try:

    if debug==False:
      if password != "":
        dm_base.set_password(get_marc_access(dm_base))

    result, table = update_marc(result_df, dm_base, dm_base.username, dm_base.password)
    audit_logs.append(cs_audit_log.audit_log('Selenium Result',result, dm_base))
    audit_logs.append(cs_audit_log.audit_log('Mapping Action','Completed...', dm_base))
    audit_logs_insert(audit_logs)
  except Exception as error:
    logging('map_action_workflow', error, dm_base)
  return pd.DataFrame(table)


def processing_validation(taskid, guid, step_name=None, source=None, destination=None, email=None, username=None, password=None):
  dm_base = session.base(taskid, guid,email,step_name, username=username, password=password, filename=source)
  dm_properties = session.data_management(source, destination)
  audit_logs = []
  audit_logs.append(cs_audit_log.audit_log('Queue','Starting...{0}'.format('processing validation'), dm_base))
  rows = []
 
  #retrieve the final processing from the validation at rpa table
  try:
    #read excel file
    raw_df = readExcel(dm_properties, dm_base, columnsname=None, dtype=str, keep_default_na=False)

    #convert N/A to Not Available
    raw_df.replace('N/A.','Not Available', inplace=True)
    raw_df = session.dataframe(raw_df)
    raw_df = restructure_marc_format(raw_df.dataframe)
    
    #convert lowercase
    raw_df.columns = map(str.lower, raw_df.columns)
    #raw_df.to_excel(r"C:\Users\alfred.simbun\Desktop\DTI-TEST\OUTPUT\RUN_20210415\RESULTS\RES_20210415_141332\raw_df.xlsx")
    
    #preprocessing to remove unnecessary value
    df = preprocessing(raw_df, dm_base)
    # We can start to remove some records here to MANUAL
    #df.to_excel(r"C:\Users\alfred.simbun\Desktop\DTI-TEST\OUTPUT\RUN_20210415\RESULTS\RES_20210415_141332\df.xlsx")

    parameter_list = get_uniqueue_list_from_multiple_columns(df, parameter_mapper["policy num"], dm_base)
    params = [parameter_list, "policy num"]
    #get_member_by_NRIC should be rename - this function is use to retrieve
    #membership records base on policy number
    marc_result, marc_column = sp_get_member_by_NRIC(params, dm_base)
    #marc_result.to_excel(r"C:\Users\alfred.simbun\Desktop\DTI-TEST\OUTPUT\RUN_20210415\RESULTS\RES_20210415_141332\membership_record.xlsx")

    marc_info = excel_db_mapper["Principal"]
    marc_column = marc_info['marc_column']
    #limit = len(params[1])
    query, prepared_params = marc_info['query'](parameter_list, dm_base)
    print('query: {0} & parameters: {1}'.format(query, prepared_params[0]))
    #get principal value from MARC
    columns = ['imppPrincipalId','Principal Name MARC','Principal IC','Principal Other IC', 'Principal Ext Ref','mmNRIC', 'mmOtherIc', 'imppId','mmId','imppEmployeeId',
               'mmFullname','mmDOB','mmGender','imppRelationship','imppExternalRefId', 'imppVIP', 'imppFirstJoinDate','imppPlanAttachDate','imppPlanExpiryDate',
               'imppCancelDate', 'inptPolPolicyNum', 'inptPolOwner', 'ippInptpolicyId','inptPlanExtCode', 'inptPlanName', 'ippInptPlanId', 'imppInptPolicyPlanId',
               'imppCreatedDate', 'imppSpecialConditions', 'imppAnnualIndUtil', 'imppImportBatchId', 'imppExtensionExpiryDate']
    marc_result_principal = sp_query_as_df(query, parameter_list, columns, dm_base)

    #To sum up all the matrix code validation
    matrixCode = []

    #ensure the df ic number is 12 digit
    raw_df['nric'] = raw_df['nric'].apply(lambda x: "{:0>12}".format(str(x)) if (x!='') else '')
    raw_df['principal nric'] = raw_df['principal nric'].apply(lambda x: "{:0>12}".format(str(x)) if (x!='') else '')
    try:
      raw_df['phone'] = raw_df['phone'].apply(lambda x: x.zfill(len(str(x))+1) if x!='' else '')
    except:
      raw_df['principals phone number'] = raw_df['principals phone number'].apply(lambda x: x.zfill(len(str(x))+1) if x!='' else '')

    print('nric: {0}'.format(raw_df['nric']))
    print('principal nric: {0}'.format(raw_df['principal nric']))

    rpa_records = raw_df.to_dict(orient="records")
    
    #login to MARC system
    audit_logs.append(cs_audit_log.audit_log('Login to marc system','Completed...', dm_base))
    #remember to remove this for encryption use

    try:
      if password != "":
        dm_base.set_password(get_marc_access(dm_base))
    except:
      pass

    global browser
    browser, flag = initBrowser(dm_base, username=dm_base.username, password=dm_base.password)
    
    #declare marc site which need to be update in looping all the record for one time retrieve
    id, url = get_rpa_application_url('marc_search_membership')
    marc_search_page = get_xpath_list(id)
    id2, url2 = get_rpa_application_url('marc_setup_member_policy')
    marc_update_member = get_xpath_list(id2)
    update_member_details_id, update_member_details_url = get_rpa_application_url('marc_update_member_details')
    marc_update_member_details = get_xpath_list(update_member_details_id)
    update_member_details_edit_id, update_member_details_edit_url = get_rpa_application_url('marc_update_member_details_edit') 
    marc_update_member_details_edit = get_xpath_list(update_member_details_edit_id)

    if flag == True:
      audit_log('Read excel row', 'Completed...', dm_base)
      excel_record_count = 0
      message = 'Total transaction processing: {0}'.format(excel_record_count)
      print(message)
      audit_logs.append(cs_audit_log.audit_log('Total transaction processing',message, dm_base))

      analytic_dm_records = []

      for excel in rpa_records:
        excel
        try:
          #converting date format compatiable with MARC
          keys = list(excel)
          for key in keys:
            key
            try:
              if len(findall("^[0-9]{4}-[0-9]{2}-[0-9]{2}$", excel[key])) > 0:
                temp = findall("^([0-9]{4})-([0-9]{2})-([0-9]{2})$", excel[key])[0]
                excel[key] = "%s/%s/%s" % (temp[2], temp[1], temp[0])
            except Exception as error:
                excel[key] = None
            excel[sub("[^a-zA-Z0-9]", " ", key)] = excel.pop(key)
            
          update_audit('Read excel row', 'Completed...Row no# {0}, record no. {1}'.format(excel["no"], excel_record_count), dm_base)

          #retrieve action code
          excel = generate_matrix_code_by_row(dm_base, dm_properties, excel, marc_result, marc_result_principal)
          
          #trigger marc update
          print('Processing automagica with row {0}'.format(excel["no"]))
          excel = automagica_marc_update(excel, dm_base, dm_properties, browser, marc_search_page, url, marc_update_member, url2, marc_update_member_details,
                                         update_member_details_url, marc_update_member_details_edit, update_member_details_edit_url)
          analytic_dm_records.append(analytic_dm_record(excel, dm_base))
          rows.append(excel)
        except Exception as error:
          print("Error for row {0}, message: {1}".format(excel["no"], error))
          continue
        excel_record_count = excel_record_count + 1
        message = 'Total transaction processing: {0}'.format(excel_record_count)
        print(message)
        audit_logs.append(cs_audit_log.audit_log('Total transaction processing',message, dm_base))
      analytic_dm(analytic_dm_records)
    else:
      logging('logging marc issue', 'password/username invalid', dm_base)

    #print('Check audit_logs {0}'.format(audit_logs))
    #audit_logs_insert(audit_logs)
    browser.close()
    audit_logs_insert(audit_logs) 
  except Exception as error:
    logging('processing_validation', error, dm_base)
    print('processing_validation error: {0}'.format(error))
    try:
      browser.close()
    except:
      print('Accessing Marc issue')
  return pd.DataFrame(rows)

def automagica_marc_update(excel, dm_base, dm_properties, browser, marc_search_page, url, marc_update_member, url2, marc_update_member_details,
                           update_member_details_url, marc_update_member_details_edit, update_member_details_edit_url):
  print('Verify action code for update marc, process data {0} started.'.format(excel["no"]))
  #audit_log('automagica_marc_update', 'Verify action code for update marc, process data {0} started.'.format(excel["no"]), dm_base)
  audit_logs = []
  audit_logs.append(cs_audit_log.audit_log('automagica_marc_update','Verify action code for update marc, process data {0} started.'.format(excel["no"]), dm_base))
  try:
    try:
      actions, params, search = convert_action_code_to_actions(excel["action_code"], excel["import type"], dm_base)        
    except:
      excel["error"] = "Could not convert action code. "
      excel["record_status"] = "Unsuccessful"
      if not "import type" in excel:
         excel["error"] = excel["error"] + "Missing import type. "
      if not "action_code" in excel:
         excel["error"] = excel["error"] + "Missing action code. "

    excel["action"] = actions
    excel["fields_update"] = params
    excel["search_criteria"] = search

    actions = actions.split(";")
    flag=False
    flags = []

    imm_id = ""
    for action in actions:
       try:
         if action == "alert":
           excel["status"] = "Completed action %s. " % action
           excel["record_status"] = "Unsuccessful"
           flag = False
         elif action == "import":
           excel["status"] = "Completed action %s. " % action
           excel["record_status"] = "Successful"
           flag = False
         else:
           if action == "update_fields":
             browser, excel, id = search_member(dm_base, browser, excel["search_criteria"], excel, marc_search_page, url, marc_update_member)
             #audit_logs.append(search_audit)
             imm_id = id
             browser, excel, flag, id = update_member(dm_base, browser,excel, marc_update_member, url2)
             #audit_logs.append(member_audit)
             imm_id = id
             flags.append(flag)
           elif action == "update_member":
             #update member
             print("imm id: {0}".format(imm_id))
             if imm_id == "":
                browser, excel, immid = search_member(dm_base, browser, excel["search_criteria"], excel, marc_search_page, url, marc_update_member)
                #audit_logs.append(search_audit)
                imm_id = immid
             browser, excel, flag = update_member_details(dm_base,browser, excel, marc_update_member_details, update_member_details_url, imm_id, marc_update_member_details_edit)
             #audit_logs.append(member_details_audit)
             flags.append(flag)

           #need to check flag for final checking
           if all(flags):
              excel["record_status"] = "Successful"
              print("Record no {0} has been successfully update.".format(excel["no"]))
           else:
             excel["error"] = excel["error"] + "<One of the action failed.>"
             excel["record_status"] = "Unsuccessful"
              

       except Exception as error:
         flag = False
         excel["record_status"] = "Unsuccessful"
         excel["error"] = excel["error"]+"<{0}>".format(error)
         logging('automagica_marc_update',error,dm_base)

    print('Verify action code for update marc, process data {0} ended.'.format(excel["no"]))
    audit_logs.append(cs_audit_log.audit_log('automagica_marc_update','Verify action code for update marc, process data {0} ended.'.format(excel["no"]), dm_base))
  except Exception as error:
    logging('automagica_marc_update','automagica_marc_update error: {0}'.format(error),dm_base)
    print('automagica_marc_update error: {0}'.format(error))
  audit_logs_insert(audit_logs) 
  return excel


def restructure_marc_format(result_df):
  #convert all date into marc system compatiable
  convert_datetime = ['DOB', 'Policy Eff Date','Policy Eff DATE','Policy End Date','Previous Policy End Date', 'Previous Policy END DATE',
                      'First Join Date','Plan Attach Date','Plan Expiry Date', 'Plan Expiry DATE', 'Date Received by AAN',
                      'Cancellation Date']
  # Using for loop to convert format that MARC able to read
  for item in convert_datetime:
    try:
      data = result_df[item]
      init = 0
      for i in data:
        #to convert date into marc
        try:
          temp = findall("^([0-9]{4})-([0-9]{2})-([0-9]{2}) ([0-9]{2}):([0-9]{2}):([0-9]{2})$", i)
          if len(temp) >= 1:
            temp = temp[0]
            date = str("{2}-{1}-{0}".format(temp[2], temp[1], temp[0]))
            #date = datetime.strptime(date, '%Y-%m-%d')
            result_df.at[init, item] =date
          else:
            result_df.at[init, item] =i
          init = init+1
        except Exception as error:
          print(error)
    except:
      pass
    continue

    #print('Result Datetime format for {1}:\n{0}'.format(result_df[item], item))
  return result_df



