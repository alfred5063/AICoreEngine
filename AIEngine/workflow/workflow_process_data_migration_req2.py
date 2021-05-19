#!/usr/bin/python
# Updated as of 29th April 2020
# Workflow - REQ File Migration Incremental Load Script
# Version 1

# Declare Python libraries needed for this script
import json
import pandas as pd
import xlrd
import numpy as np
import os
import pymssql
from datetime import date
from connector.connector import *
from pandas.io.json import json_normalize
from utils.Session import session
from utils.guid import get_guid
from utils.audit_trail import audit_log
from utils.notification import send
from utils.logging import logging
from data_migration.Ops_Disbursement_Master import *
from configparser import ConfigParser
from connector.dbconfig import read_db_config
from pathlib import Path
from connector.connector import MSSqlConnector
from loading.query.query_as_df import *

###Please use this folder
#migration_folder = r'\\dtisvr2\CBA_UAT\Result\REQDCA_Paths\1. DISBURSEMENT CLAIMS INVOICE LOG FILE\2019\REQ'

###testing folder
#migration_folder = r'C:\Users\NURFAZREEN.RAMLY\Desktop\dm_req\req2021'
#year = 2021
#email = 'nurfazreen.ramly@asia-assistance.com'
#myjson = {'parent_guid': '5a48c5a9-5d9c-4ac9-9a0c-bf6a7d5f2ee0', 'guid': '8fba8440-6b31-472a-b0ca-3ebedc95fad5',
#          'authentication': {'username': '', 'password': 'Gcdnr5B0SVtfNvCnrYJoGg=='},
#          'email': 'nurfazreen.ramly@asia-assistance.com',
#          'stepname': 'process_data_migration_req',
#          'step_parameters': [{'id': 'tr_prop_1_step_1', 'key': 'req_path', 'value': '\\\\dtisvr2\\CBA_UAT\\Data Migration\\REQ\\2021', 'description': 'dir'}],
#          'taskId': '2465', 'execution_parameters_att_0': '', 'execution_parameters_txt': {'docType': 'Data Migration REQ', 'parameters': [{'year': '2021'}]},
#          'preview': 'False'}


#guid = str(myjson['guid'])
#stepname = str(myjson['stepname'])
#password = str(myjson['authentication']['password'])
#username = str(myjson['authentication']['username'])
#email = str(myjson['email'])
#taskid = str(myjson['taskId'])
#preview = str(myjson['preview'])
#year = int(myjson['execution_parameters_txt']['parameters'][0]['year'])

#for item in myjson['step_parameters']:
#  if item['key']=='req_path':
#    migration_folder = str(item['value'])


def process_data_migration_REQ_incremental_load(migration_folder, year, email):

  req_base = session.base(taskid, guid, email, stepname, username = username, password = password)
  #read database [asia_assistance_rpa].[cba].[ops_disbursement_req_master]

  #current_path = os.path.dirname(os.path.abspath(__file__))
  #dti_path = str(Path(current_path).parent)
  #config = read_db_config(dti_path+r'\config.ini', 'mssql')
  #config = read_db_config(r'C:\Users\NURFAZREEN.RAMLY\Source\Repos\rpa_engine\RPA.DTI\config.ini', 'mssql')
  #serv = config['server']
  #db = config['database']
  #userdb = config['user']
  #pwd = config['password']
  #conn = pymssql.connect(server=serv, user=userdb, password=pwd,
  #                        database=db)
  
  path = get_filenames(migration_folder)
  for filepath in path:
      with xlrd.open_workbook(filepath) as wb:
          print([filepath])
          sheet = xlrd.open_workbook(filepath).sheet_names()
          for filesheet in sheet:
              worksheet = wb.sheet_by_name(filesheet)
              print(filesheet)
              row_data = pd.read_excel(filepath, sheet_name = filesheet, header = 3)
              print(row_data)
              if len(row_data) != 0: 
                row_data = row_data[pd.notnull(row_data['Date'])] # Only take records where Date is the latest date
                print("checkpoint1")
                row_data = row_data.rename(columns =
                                           {"Date" : "Date",
                                            "Disbursement claims no" : "Disbursement claims no",
                                            "Claims listing no" : "Claims listing no",
                                            "SAGE ID" : "SAGE ID",
                                            "File" : "File",
                                            "Bill to" : "Bill to",
                                            "Client" : "Client",
                                            "Patient Name" : "Patient Name",
                                            "Status (VIP / Non-Vip Tan chong claims) " : "Status (VIP / Non-Vip Tan chong claims) ",
                                            "Policy no " : "Policy no ",
                                            "Admission date" : "Admission date",
                                            "Discharge date" : "Discharge date",
                                            "Hospital" : "Hospital",
                                            "Bill no" : "Bill no",
                                            "Amounts" : "Amounts",
                                            "OB Received date" : "OB Received date",
                                            "OB Registered date" : "OB Registered date",
                                            "Reasons" : "Reasons",
                                            "Cashless / Post / Fruit Basket / Reimbursement " : "Cashless / Post / Fruit Basket / Reimbursement ",
                                            "Initial" : "Initial",
                                            "Bank in date" : "Bank in date",
                                            "Details" : "Details",
                                            "Amounts Received" : "Amounts Received",
                                            "AP Team Remarks" : "AP Team Remarks",
                                            "Date payment issued" : "Date payment issued",
                                            "Chq No. / Giro" : "Chq No. / Giro",
                                            "Amount (RM)" : "Amount (RM)",
                                            "BIMY 01" : "BIMY 01",
                                            "DBMY37 (Etiqa)" : "DBMY37 (Etiqa)",
                                            "HLBB 2 (Axa:Chq)" : "HLBB 2 (Axa:Chq)",
                                            "DBMY14 (Axa:Giro)" : "DBMY14 (Axa:Giro)",
                                            "DBMY15 (RHB:Giro)" : "DBMY15 (RHB:Giro)",
                                            "DBMY16 (Mpower/SS2)" : "DBMY16 (Mpower/SS2)",
                                            "DBMY42 (STMB)" : "DBMY42 (STMB)",
                                            "DBMY39 (Finance) Borrowing Bank Account" : "DBMY39 (Finance) Borrowing Bank Account",
                                            "DBMY38 (Operation)" : "DBMY38 (Operation)",
                                            "Repayment Date" : "Repayment Date"})

              
                connector = MSSqlConnector
                query = "SELECT * FROM [asia_assistance_rpa_uat].[cba].[ops_disbursement_req_master] WHERE [year] = %s"
                params = year
                print("check2")
                df1 = mssql_get_df_by_query_without_base(query, params, connector)
                if len(df1) != 0:
                  #df1 = pd.read_sql(query, conn)
                  dict_obj = json.loads(df1.loc[0,'content'])
                  df2= pd.read_json(json.dumps(dict_obj.get('Data')))
                  ##Incremental_load
                  counter=len(row_data)-len(df2)
                  newRow = len(df2)
                  df3 = df2
                  i = 0
                  for i in range(counter):
                    df3 = df3.append(row_data.iloc[newRow+i],ignore_index=True)
                  json_df = df3.to_json(orient='records', lines=False, date_format='iso')
                  new_df = json.dumps({"Data": json.loads(json_df)})
                  print(new_df)
                  insert_mssql(new_df, cmurl, year21, email)
                  cursor = conn.cursor()
                  sql = "UPDATE [asia_assistance_rpa].[cba].[ops_disbursement_req_master] SET [content] = %s WHERE [year]= %s"
                  val = (new_df,year)
                  cursor.execute(sql, val)
    
                  conn.commit()
                  conn.close()
                  print('Data migration incremental load for REQ into database - done')
                  send(req_base, email, "REQ Data Migration - RPA Task Execution Completed.", "Hi, <br/><br/><br/>You task has been completed. <br/>Reference Number: " + str(req_base.guid) + "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
                else:
                  query = "SELECT * FROM [asia_assistance_rpa_uat].[cba].[ops_disbursement_req_master] WHERE [year] = %s"
                  params = year - 1
                  df1_previous = mssql_get_df_by_query_without_base(query, params, connector)
                  dict_obj = json.loads(df1_previous.loc[0,'content'])
                  df2= pd.read_json(json.dumps(dict_obj.get('Data')))
                  df3 = row_data
                  i = 0
                  for i in range(len(df2)):
                    df3 = df3.append(df2.iloc[i],ignore_index=True)
                  json_df = df3.to_json(orient='records', lines=False, date_format='iso')
                  new_df = json.dumps({"Data": json.loads(json_df)})
                  print(new_df)
                  insert_mssql(new_df, filepath, year, email)
                  print('Data migration incremental load for REQ into database - done')
                  send(req_base, email, "REQ Data Migration - RPA Task Execution Completed.", "Hi, <br/><br/><br/>You task has been completed. <br/>Reference Number: " + str(req_base.guid) + "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
              else:
                pass
  return 'completed'

def read_db_config(filename, section):
    # create parser and read ini configuration file
    parser = ConfigParser()
    parser.read(filename)
    #print('parser')
    #print(parser)

    # get section, default to mysql
    db = {}
    if parser.has_section(section):
        #print('FOUND SECTION')
        items = parser.items(section)
        for item in items:
            db[item[0]] = item[1]
    else:
        raise Exception('{0} not found in the {1} file'.format(section, filename))
    return db

def get_filenames(root):
    filename_list = []
    for path, subdirs, files in os.walk(root):
        for filename in filter(excel_file_filter, files):
            filename_list.append(os.path.join(path, filename))
    return filename_list

#function to check file extension. Only support xls. and xlsx.and return file path
def excel_file_filter(filename, extensions=['.xls', '.xlsx']):
    return any(filename.endswith(e) for e in extensions)


