#!/usr/bin/python
# FINAL SCRIPT updated as of 15th April 2021

# Declare Python libraries needed for this script
from utils.Session import session
from utils.audit_trail import *
from utils.logging import logging
from output.save_output import save_process_output
from output.get_action_codes_from_db import get_action_codes_from_db
from selenium_interface.accessMarc import update_marc
from loading.excel.createexcel import create_excel_dm
from directory.movefile import move_file_to_archived, move_file_to_result_dm
from directory.get_filename_from_path import get_file_name
from output.get_action_codes_from_db import get_matrix_lookup
from analytics.marc_analytics import analytic_dm
import pandas as pd
from datetime import datetime
from dateutil.parser import parse
import numpy as np
from regex import sub, findall
import shutil
from utils.notification import send
import numpy as np

#@iterate_source

#taskid
#guid
#result_df
#step_name=stepname
#source=source
#destination=phasezero_des
#email=email


def post_processing_workflow(taskid, guid, result_df, step_name, source=None, destination=None, email=None):
  dm_base = session.base(taskid, guid,email,step_name, filename=source)
  dm_properties = session.data_management(source, destination)
  print('destination: {0}'.format(destination))
  audit_logs = []
  audit_logs.append(cs_audit_log.audit_log('Queue','Starting...', dm_base))
  #audit_log('Queue', 'Starting...', dm_base)
  try:
    result_df = remove_column(result_df, dm_base)
    #result_df = restructure_marc_format(result_df, dm_base)
    result_df = restructure_order(result_df)

    head, filename = get_file_name(dm_properties.source)
    #copy a result file
    copy_path = head+"\\result_"+filename
    shutil.copy(head+"\\"+filename, copy_path)

    columns_index = [9, 29, 30, 32, 36, 37, 38, 45, 46]
    columns_text_index = [6, 7, 8, 15, 23, 33, 39, 40, 41]

    result_df.replace('nan','', inplace=True)
    #ensure the df ic number is 12 digit
    result_df['nric'] = result_df['nric'].apply(lambda x: "{:0>12}".format(str(x)) if (x!='') else '')
    result_df['principal nric'] = result_df['principal nric'].apply(lambda x: "{:0>12}".format(str(x)) if (x!='') else '')
    #result_df['phone'] = result_df['phone'].apply(lambda x: x.zfill(len(str(x))+1) if(x!='') else '')

    result_df['nric'] = result_df['nric'].apply(lambda x: np.nan if (x=='000000000000') else x)
    result_df['principal nric'] = result_df['principal nric'].apply(lambda x: np.nan if (x=='000000000000') else x)

    #create new sheet with same file and move to result folder
    create_excel_dm(dm_properties, result_df, dm_base, columns_index, columns_text_index)

    result_path = []
    result_path = move_file_to_result_dm(copy_path, dm_base)
    archieved_path = move_file_to_archived(dm_properties.source, dm_base)

    
    #if dm_base.email != None:
    content = ("<br/>Reference Number: " + str(guid) + "<br/>"+
               "File has been archieved at following link: <"+archieved_path+"><br/>"+
               "Result file as following link: <"+result_path+"><br/>"+
               "<br/>To confirm whether the process folder is empty, check following path: <"+head+"><br/>"+
               "You may rerun this either reupload the file or rerun button for a full run.<br/>")
    
    #audit_log('Task Complete Process', 'Completed', dm_base)
    audit_logs.append(cs_audit_log.audit_log('Task Complete Process','Completed', dm_base))
    #audit_log('Post Processing', 'Completed...', dm_base)
    audit_logs.append(cs_audit_log.audit_log('Post Processing','Completed...', dm_base))
  except Exception as error:
    logging('post_processing_workflow', error, dm_base)
  #update_audit(guid, filename=filename, email=email)
  audit_logs_insert(audit_logs) 
  return 'success',result_path,content

def restructure_order(result):
  result = result[['no', 'import type',	'member full name',
                   'address 1',	'address 2',	'address 3',	'address 4',	'gender',
                   'dob',	'nric',	'otheric',	'external ref id  aka client', 	'internal ref id  aka aan',
                   'employee id',	'marital status',	'race',	'phone',	'vip', 'special condition',
                   'relationship',	'principal int ref id  aka aan', 	'principal ext ref id  aka client',
                   'principal name',	'principal nric',	'principal other ic',	'program id',	'policy type',
                   'policy num',	'policy eff date',	'policy end date',	'previous policy num',	'previous policy end date',
                   'policy owner',	'external plan code',	'internal plan code id',	'first join date',	'plan attach date',
                   'plan expiry date',	'insurer branch',	'insurer agency',	'insurer agency code',	'insurer mco fees',
                   'ima service', 'ima limit',	'date received by aan',	'cancellation date',
                   'action_code',	'fields_update',	'action',	'search_criteria', 'error',	'record_status']]
  return result

def remove_column(result, dm_base):
  try:
    result = result.drop(['status'], axis=1)
  except Exception as error:
    logging('Remove column to generate excel error', error, dm_base)
    print(error)

  result = result.rename(columns={
  #  #'first join date dm': 'first join date',
  #  #                              'plan attach date dm': 'plan attach date',
  #  #                              'plan expiry date dm': 'plan expiry date',
                                  'external ref id  aka client ':'external ref id  aka client',
                                  'internal ref id  aka aan ': 'internal ref id  aka aan',
                                  'principal int ref id  aka aan ':'principal int ref id  aka aan',
                                  'principal ext ref id  aka client ': 'principal ext ref id  aka client',
                                  'ima service ':'ima service'})
 
  return result

def convert_employee_id_to_str(result_df):
  result_df['employee id'] = str(result_df['employee id'])
  return result_df

def restructure_marc_format(result_df, dm_base):
  #convert all date into marc system compatiable
  convert_datetime = ['dob', 'policy eff date','policy end date','previous policy end date',
                      'first join date','plan attach date','plan expiry date','date received by aan',
                      'cancellation date']
  # Using for loop to convert format that MARC able to read
  for item in convert_datetime:
    data = result_df[item]
    init = 0
    for i in data:
      #to convert date into marc 
      try:
        temp = findall("^([0-9]{4})-([0-9]{2})-([0-9]{2})$", i)
        if len(temp) >= 1:
          temp = temp[0]
          date = "{0}/{1}/{2}".format(temp[2], temp[1], temp[0])
          result_df.at[init, item] =date
        else:
          result_df.at[init, item] = ''
        init = init+1
      except Exception as error:
        logging('Restructure date format for marc error', error, dm_base)
        print(error)

    #print('Result Datetime format for {1}:\n{0}'.format(result_df[item], item))
  return result_df


