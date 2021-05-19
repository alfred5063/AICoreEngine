import datetime
from connector.connector import MSSqlConnector
from regex import sub, findall

def analytic_dm_record(row, dm_base):
  currenttime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
  return (str(row['no']),str(row['import type']),str(row['address 2']),str(row['address 3']),str(row['address 4'])
                    ,str(row['gender']),datetime.datetime.strptime(row['dob'], '%d/%m/%Y'),str(row['marital status']),str(row['policy type']),str(row['policy owner'])
                    ,str(row['action_code']),row['action'],str(row['fields_update']),str(row['search_criteria']),str(row['error'])
                    ,str(row['record_status']),str(dm_base.guid),currenttime, dm_base.email)

def analytic_dm(records):
  connector = MSSqlConnector()
  cursor = connector.cursor()
 
  qry = '''INSERT INTO [dm].[validation_membership]
             ([no]
             ,[import_type]
             ,[city]
             ,[state]
             ,[postcode]
             ,[gender]
             ,[dob]
             ,[marital_status]
             ,[policy_type]
             ,[policy_owner]
             ,[action_code]
             ,[action]
             ,[fields_update]
             ,[search_criteria]
             ,[error]
             ,[record_status]
             ,[session_id]
             ,[createdOn]
             ,[createdBy])
       VALUES
             (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            '''
  cursor.executemany(qry, records)
  connector.commit()

  cursor.close()
  connector.close()

def analytic_dm_by_row(row, dm_base):
  connector = MSSqlConnector()
  cursor = connector.cursor()
  currenttime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
  sessionid = dm_base.guid
 
  qry = '''INSERT INTO [dm].[validation_membership]
             ([no]
             ,[import_type]
             ,[city]
             ,[state]
             ,[postcode]
             ,[gender]
             ,[dob]
             ,[marital_status]
             ,[policy_type]
             ,[policy_owner]
             ,[action_code]
             ,[action]
             ,[fields_update]
             ,[search_criteria]
             ,[error]
             ,[record_status]
             ,[status]
             ,[session_id]
             ,[createdOn]
             ,[createdBy])
       VALUES
             (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            '''
  param_values = (str(row['no']),str(row['import type']),str(row['address 2']),str(row['address 3']),str(row['address 4'])
                    ,str(row['gender']),datetime.datetime.strptime(row['dob'], '%d/%m/%Y'),str(row['marital status']),str(row['policy type']),str(row['policy owner'])
                    ,str(row['action_code']),row['action'],str(row['fields_update']),str(row['search_criteria']),str(row['error'])
                    ,str(row['record_status']),str(row['status']),str(sessionid),currenttime, dm_base.email)

  cursor.execute(qry, param_values)
  connector.commit()

  cursor.close()
  connector.close()
