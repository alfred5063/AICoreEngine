import pandas as pd
import numpy as np
from connector.connector import MSSqlConnector
from utils.audit_trail import audit_log
from utils.logging import logging
from loading.query.query_as_df import mssql_get_df_by_query, mssql_get_df_by_query_without_param
from loading.query.query_as_df import convert_query_as_df
from directory.get_filename_from_path import get_file_name
#from extraction.encryption.encryption import decrypt_json_string

def get_action_codes_from_db(base, properties):
  try:
    audit_log('get_action_codes_from_db' , 'load result as dataframe. get action code from process_output table.', base)
    head, filename = get_file_name(properties.source)
    connector = MSSqlConnector
    query = '''SELECT * FROM dm.process_outputs where task_id = %s and filename = %s and guid = %s'''
    params = (base.taskid, base.filename, base.guid)
    result_df = mssql_get_df_by_query(query, params, base, connector)
    #task_info = decrypt_json_string(result_df['task_info'])
    task_info = result_df['task_info']
    task_list = task_info.to_numpy()
    code_list, first_df = convert_df_to_action_codes(task_list, base)
    df_with_codes = get_raw_df_with_action_codes(first_df, code_list, base)
    return df_with_codes
  except Exception as error:
    logging('get_action_codes_from_db', error, base)
    raise error

def get_matrix_lookup(base):
  try:
     audit_log('get_matrix_lookup', 'load matrix lookup table.', base)
     connector = MSSqlConnector
     query = '''SELECT [action_code], [actionParam], [action] FROM 
                (SELECT [matrixCode] as action_code, [actionParam] as actionParam, [action] FROM dm.x_matrix_lookup
                UNION
                SELECT [matrixCode] as action_code, [actionParam] as actionParam, [action] FROM dm.matrix_lookup) result
                ORDER BY action_code ASC'''
     result_df = mssql_get_df_by_query_without_param(query, base, connector)
     print(result_df)
     result_df['action_code'] = result_df['action_code'].astype(int)
     
     return result_df
  except Exception as error:
    logging('get_matrix_lookup', error, base)

def convert_df_to_action_codes(task_list, base):
  audit_log('convert_df_to_action_codes' , 'convert action code from process_output table into data frame.', base)
  try:
    task_list = [pd.read_json(row, orient='records') for row in task_list]
    merge_list = [np.array([task['_merge'].to_numpy()]).T for task in task_list]
    stack_list = np.array([])
    for merge in merge_list:
      stack_list = np.hstack([stack_list, merge]) if stack_list.size else merge
    action_codes = np.sum(stack_list, axis=1)
  except Exception as error:
    logging('convert_df_to_action_codes', error, base)
    print('convert_df_to_action_codes error: {0}'.format(error))
  return action_codes, task_list[0]

def convert_action_code_to_actions(action_code, import_type, base):
  audit_log('convert_action_code_to_actions' ,
           'convert action code from process_output table into data frame. include import type for identify matrix_lookup or x_matrix_lookup',
           base)
  print('Extract action code from db')
  connector = MSSqlConnector
  if import_type == "X":
    table = "[x_matrix_lookup]"
  else:
    table = "[matrix_lookup]"
  query = '''SELECT * FROM [dm].{0} where [dm].{1}.[matrixCode] = %s'''.format(table, table)
  param = (str(action_code))
  #print('query: {0}, param: {1}'.format(query, param))
  result_df = mssql_get_df_by_query(query, param, base, connector)
  
  try:
    return result_df['action'][0], result_df['actionParam'][0], result_df['searchCriteria'][0]
  except Exception as error:
    logging('convert_action_code_to_actions: %s' % action_code, error, base)
    print('convert_action_code_to_actions: {0}'.format(action_code, error))
    raise Exception('Action lookup for code %s failed' % action_code)


def get_raw_df_with_action_codes(result_df, code_list, base):
  audit_log('get_raw_df_with_action_codes' , 'assign code list into array.', base)
  result_df['action_code'] = code_list
  return result_df
