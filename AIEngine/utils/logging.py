from connector.connector import MSSqlConnector
import datetime

def logging(step_name, error, session, status = True):
  if status == True:
    taskid = session.taskid
    sessionid = session.guid
    email = session.email
    filename = session.filename
    connector = MSSqlConnector()
    cursor = connector.cursor()
    currenttime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    qry = '''INSERT INTO dbo.error
             (taskid
             ,name
             ,error
             ,"createdOn"
             ,sessionid, filename, "createdBy")
            VALUES
            (%s,%s,%s,%s,%s,%s,%s)
            '''
    param_values = (taskid, str(step_name), str(error), currenttime, str(sessionid), str(filename), str(email))
    cursor.execute(qry, param_values)
    connector.commit()
    cursor.close()
    connector.close()

def logging_record(step_name, error, session):
  currenttime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
  return (session.taskid, str(step_name), str(error), currenttime, str(session.guid), str(session.filename), str(session.email))

def logging_insert(loggings):
  if status == True:
    connector = MSSqlConnector()
    cursor = connector.cursor()
    qry = '''INSERT INTO dbo.error
             (taskid
             ,name
             ,error
             ,"createdOn"
             ,sessionid, filename, "createdBy")
            VALUES
            (%s,%s,%s,%s,%s,%s,%s)
            '''
    cursor.executemany(qry, param_values)
    connector.commit()
    cursor.close()
    connector.close()


