import mysql.connector
from mysql.connector import Error
from connector.dbconfig import read_db_config
import os
from utils.audit_trail import audit_log
from utils.logging import logging
from pathlib import Path

current_path = os.path.dirname(os.path.abspath(__file__))
dti_path = str(Path(current_path).parent)

def Connector(base, config=None):
  """ Connect to MySQL database """
  #os.chdir("")
  audit_log('Connector', 'Connect MySQL', base)
  try:
    conn = None
    if not config:
      #config = read_db_config(r'C:\Asia-Assistance\RPA3.0\SourceCode\RPA.DTI\config.ini', 'mysql')
      config = read_db_config(dti_path+r'\config.ini', 'mysql')
    print('Connecting to MARC database...')
    conn = mysql.connector.connect(host=config['host'],database=config['database'],
                                       user=config['user'], password=config['password'])

    if conn.is_connected():
      print('connection established.')
      return conn
    else:
        print('connection failed.')
  except Error as error:
    logging('Connector', error, base)
    print(error)

#Example
#config = read_db_config('D:\OneDrive\OneDrive - AXA\AsiaAssistance\SourceCode\RPA.DTI\config.ini', 'mysql')
#conn = Connector(config)
#print(conn)
