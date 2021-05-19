import pandas as pd
import functools
import operator
from connector.connector import MySqlConnector
from connector.connector import MSSqlConnector
from utils.audit_trail import audit_log
from utils.logging import logging

def flatten(base, args):
  audit_log('flatten', 'flatten array.', base)
  try:
    flatten_list = []
    for arg in args:
      if isinstance(arg, list):
        flatten_arg = functools.reduce(operator.iconcat, [arg], [])
        flatten_list.extend(flatten_arg)
      else:
        flatten_list.append(arg)
  except Exception as error:
    logging('flatten', error, base)
  return flatten_list

def validate_type_list(arg, base):
  audit_log('validate_type_list', 'validate data type.', base)
  return False if isinstance(arg, list) and len(arg) == 0 else True

def prepare_statement(arg):
  result = ','.join(['%s'] * len(arg)) if isinstance(arg, list) else '%s'
  return result

def bind_parameters(query, params, base):
  audit_log('bind_parameters', 'binding parameters for comparison.', base)
  formatted_params = [prepare_statement(arg) for arg in params ]
  query = query.format(*formatted_params)
  return query

def mssql_get_df_by_query_without_base(query, params, connector):
 
  connector = connector()
  cursor = connector.cursor(as_dict=True)
  cursor.execute(query, params)
  data=cursor.fetchall()
  
  result=pd.DataFrame(data)
  
  cursor.close()
  connector.close()
  return result

def mssql_get_df_by_query(query, params, base, connector):
  audit_log('mssql_get_df_by_query', 'passing query to execute and return data frame result', base)
  connector = connector()
  cursor = connector.cursor(as_dict=True)
  cursor.execute(query, params)
  data=cursor.fetchall()
  result=pd.DataFrame(data)
  
  cursor.close()
  connector.close()
  return result

def execute_query_preprod(query, base, connector, params=None):
  audit_log('execute_query_preprod', 'passing query to execute and return data frame result', base)
  connector = connector()
  cursor = connector.cursor()
  if params == None: 
    cursor.execute(query)
  else:
    cursor.execute(query, params)
  data=cursor.fetchall()

  if len(data) == 0:
      result=pd.DataFrame()
  else:
    result=pd.DataFrame.from_records(data)
    result.columns = cursor.column_names

  cursor.close()
  connector.close()
  return result

def mssql_get_df_by_query_without_param(query, base, connector):
  audit_log('mssql_get_df_by_query', 'passing query to execute and return data frame result', base)
  connector = connector()
  cursor = connector.cursor(as_dict=True)
  cursor.execute(query)
  data=cursor.fetchall()
  result=pd.DataFrame(data)
  
  cursor.close()
  connector.close()
  return result


def query_as_df(query, params, base,connector=None):
  print("query_as_df..." )
  audit_log('query_as_df', 'get query and convert value to data frame.', base)
  try:
    valid = validate_type_list(params, base)
    query = bind_parameters(query, params, base)
    params = tuple(flatten(base,params))
    connector = connector() if connector else MySqlConnector()
    result = pd.read_sql(query, con=connector, params=params)
    
  except Exception as error:
    logging('query_as_df', error, base)
  return result

def sp_query_as_df(query, params, columns, base, connector=None):
  print("Get value from MARC table..." )
  audit_log('sp_query_as_df', 'get query and convert value to data frame.', base)
  try:
    #valid = validate_type_list(params)
    connector = connector() if connector else MySqlConnector()
    cursor = connector.cursor()
    #args = ("'17PJB0003271-02', '17PJB0003271-01'",)
    params_value = ""
    str_param = ""
    i = 1
    for param in params[0]:
      if len(params[0]) == i:
        str_param = str_param + "'"+ str(param) +"'"
      else:
        str_param = str_param + "'"+ str(param) +"',"

      i = i+1
    
    params_value = (str_param,)
    print('query: {1} & params_value: {0}'.format(tuple(params_value), query))
    stored_proc = cursor.callproc(query, tuple(params_value))
    print(stored_proc)
    # EXTRACT RESULTS FROM CURSOR
    for i in cursor.stored_results(): results = i.fetchall()

    # LOAD INTO A DATAFRAME
    result = pd.DataFrame(results, columns=columns)
    #print(result)
    #result = pd.read_sql(query, con=connector, params=params)
  except Exception as error:
    print("Error {0}".format(error))
    logging('sp_query_as_df', error, base)
  return result

def convert_query_as_df(query, base, params=None):
  print("convert_query_as_df..." )
  audit_log('convert_query_as_df', 'get query and convert value to data frame.', base)
  try:
    connector = MSSqlConnector()
    
    result = pd.read_sql(query, con=connector)
  except Exception as error:
    logging('convert_query_as_df', error, base)
  return result
