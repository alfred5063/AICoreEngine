import json
import datetime
from connector.connector import MSSqlConnector
from utils.audit_trail import audit_log
from utils.logging import logging
from loading.excel.createexcel import create_excel
from output.get_action_codes_from_db import get_action_codes_from_db
from directory.get_filename_from_path import get_file_name
from output.get_action_codes_from_db import get_matrix_lookup
import pandas as pd
#from extraction.encryption.encryption import encrypt_json_string

def save_process_output(base, result_json, properties):
  try:
    audit_log('save_process_output' , 'Save json result into DB', base)
    head, filename = get_file_name(properties.source)
    connector = MSSqlConnector()
    cursor = connector.cursor()
    currenttime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    qry = '''INSERT INTO dm.process_outputs
             (task_id, guid, task_info, filename, timestamp)
            VALUES
            (%s, %s, %s, %s, %s)
            '''
    #encrypted_string = encrypt_json_string(result_json)
    #param_values = (task_id, guid, json.dumps(encrypted_string), filename, currenttime)
    param_values = (base.taskid, base.guid, json.dumps(result_json), filename, currenttime)
    cursor.execute(qry, param_values)
    connector.commit()
    cursor.close()
    connector.close()

    if destination:
      result_df = get_action_codes_from_db(base, properties) # If you want to try to combining the output
      
      #add additional column for user easy to know what is the action type
      matrix_df = get_matrix_lookup()
     
      result_df = pd.merge(result_df, matrix_df, left_on='action_code', right_on='action_code')
      create_excel(properties, result_df)
  except Exception as error:
    logging('save_process_output', error, base)
    raise error




