import psycopg2
from connector.dbconfig import read_db_config
import datetime
from connector.connector import MSSqlConnector
import pandas as pd

class cs_audit_log:
  def audit_log(name, reason, session):
    currenttime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return (session.taskid, str(name), str(reason), currenttime, str(session.guid), str(session.email), str(session.filename))

def audit_logs_insert(audit_logs):
  try:
    connector = MSSqlConnector()
    cursor = connector.cursor()
    
    qry = '''INSERT INTO dbo.jobs
             (taskid
             ,name
             ,reason
             ,"createdOn"
             ,sessionid,"createdBy", "filename")
            VALUES
            (%s,%s,%s,%s,%s,%s,%s)
            '''
    cursor.executemany(qry, audit_logs)
    connector.commit()
    cursor.close()
    connector.close()
  except Exception as error:
    print('SQL Error: {0}'.format(error))

def audit_log(name, reason, session, status = True):
  if status == True:
    connector = MSSqlConnector()
    cursor = connector.cursor()
    currenttime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    qry = '''INSERT INTO dbo.jobs
             (taskid
             ,name
             ,reason
             ,"createdOn"
             ,sessionid,"createdBy", "filename")
            VALUES
            (%s,%s,%s,%s,%s,%s,%s)
            '''
    taskid = session.taskid
    sessionid = session.guid
    email = session.email
    filename = session.filename

    param_values = (taskid, str(name), str(reason), currenttime, str(sessionid), str(email), str(filename))
    cursor.execute(qry, param_values)
    connector.commit()
    cursor.close()
    connector.close()

def update_audit(name, reason, session):
  taskid = session.taskid
  sessionid = session.guid
  email = session.email
  filename = session.filename
  currenttime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
  connector = MSSqlConnector()
  cursor = connector.cursor()
  qry = '''UPDATE dbo.jobs
	         SET reason=%s, filename=%s, "createdBy"=%s, "createdOn"=%s
	         WHERE sessionid=%s and name=%s and taskid=%s;'''
  #print(qry)
  param_values = (str(reason),str(filename), str(email), currenttime, str(sessionid), str(name), taskid)
  cursor.execute(qry, param_values)
  connector.commit()
  cursor.close()
  connector.close()

