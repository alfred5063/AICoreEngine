import psycopg2
from connector.dbconfig import read_db_config
import datetime
from connector.connector import MSSqlConnector

def insert_process_dm_validation(key, matrixCode, marc_field_name, excel_field_name, base):
  connector = MSSqlConnector()
  cursor = connector.cursor()
  
  qry = '''INSERT INTO [dm].[process_dm_validation]
           ([key]
           ,[matrixCode]
           ,[sessionId]
           ,[marc_field_name]
           ,[excel_field_name]
           ,[taskid])
	VALUES (%s,%s,%s,%s,%s,%s);
          '''
  taskid = base.taskid
  sessionid = base.guid
  
  param_values = (str(key), int(matrixCode), str(sessionid), str(marc_field_name), str(excel_field_name), int(taskid))
  cursor.execute(qry, param_values)
  connector.commit()
  cursor.close()
  connector.close()

