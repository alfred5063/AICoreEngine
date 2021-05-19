import pandas as pd
import numpy as np
from connector.connector import MSSqlConnector
from utils.audit_trail import audit_log
from utils.logging import logging
from loading.query.query_as_df import mssql_get_df_by_query
from loading.query.query_as_df import convert_query_as_df
from directory.get_filename_from_path import get_file_name
#from extraction.encryption.encryption import decrypt_json_string

def get_process_dm_validation(base):
  try:
    audit_log('get_process_dm_validation' , 'load result as dataframe. get process_dm_validation table result.', base)
    connector = MSSqlConnector
    query = '''SELECT [id]
                     ,[key]
                     ,[matrixCode]
                     ,[sessionId]
                     ,[marc_field_name]
                     ,[excel_field_name]
                     ,[taskid]
               FROM [dm].[process_dm_validation]
	             WHERE [sessionId] = %s and [taskid] = %s'''
    params = (base.guid, base.taskid)
    result_df = mssql_get_df_by_query(query, params, base, connector)
    
    return result_df
  except Exception as error:
    logging('get_process_dm_validation', error, base)
    raise error

def get_process_dm_validation_bykey(key, base):
  try:
    audit_log('get_process_dm_validation_bykey' , 'load result as dataframe. get process_dm_validation table result.', base)
    connector = MSSqlConnector
    query = '''SELECT [matrixCode]
	             FROM [dm].[process_dm_validation]
	             WHERE [sessionId] = %s and [taskid] = %s and [key] = %s'''
    params = (base.guid, base.taskid, key)
    result_df = mssql_get_df_by_query(query, params, base, connector)
    return result_df
  except Exception as error:
    logging('get_process_dm_validation_bykey', error, base)
    raise error


