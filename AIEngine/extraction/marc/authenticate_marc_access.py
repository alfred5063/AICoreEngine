from loading.query.query_as_df import mssql_get_df_by_query
import pandas as pd
import numpy as np
from connector.connector import MSSqlConnector
from utils.audit_trail import audit_log
from utils.logging import logging
from extraction.crypto.crypto import decrypt_pass

def get_marc_access(base):
  #session = Session(task_id, sessionid)
  try:
    #audit_log('get_marc_access' , 'extract key from rpa to decryp process for marc to authenticate.')
    connector = MSSqlConnector
    query = '''SELECT guid, public_key, salt, task_id
	             FROM dbo.encryption
               WHERE guid = %s and task_id = %s;'''
    params = (base.guid, base.taskid)
    result_df = mssql_get_df_by_query(query, params, base, connector)
    
    #result_df.iloc[0,2]
    salt = bytes(str(result_df.iloc[0,2]), 'ascii')
    public_key = bytes(str(result_df.iloc[0,1]), 'ascii')

    password = decrypt_pass(base.password, salt, public_key)
    
    return password
  except Exception as error:
    logging('get_marc_access', error, base)
    raise error
    

