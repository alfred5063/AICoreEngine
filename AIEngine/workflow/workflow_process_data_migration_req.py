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
from data_migration.Ops_Disbursement_Master import *
from configparser import ConfigParser
from connector.dbconfig import read_db_config
from pathlib import Path

###Please use this folder
migration_folder = r'\\dtisvr2\CBA_UAT\Result\REQDCA_Paths\1. DISBURSEMENT CLAIMS INVOICE LOG FILE\2019\REQ'

###testing folder
#migration_folder = r'C:\Users\Asus\Desktop\migration folder'

year = '2020'
email = 'alfred.simbun@asia-assistance.com'

def process_data_migration_REQ_incremental_load(migration_folder, year, email):
  #read database [asia_assistance_rpa].[cba].[ops_disbursement_req_master]

  #current_path = os.path.dirname(os.path.abspath(__file__))
  #dti_path = str(Path(current_path).parent)
  #config = read_db_config(dti_path+r'\config.ini', 'mssql')
  config = read_db_config(r'C:\Users\Asus\Desktop\rpa_engine\RPA.DTI\config.ini', 'mssql')
  serv = config['server']
  db = config['database']
  userdb = config['user']
  pwd = config['password']
  conn = pymssql.connect(server=serv, user=userdb, password=pwd,
                          database=db)

  query = "SELECT * FROM [asia_assistance_rpa].[cba].[ops_disbursement_req_master] WHERE [year] = %s"%year
  df1 = pd.read_sql(query, conn)

  dict_obj = json.loads(df1.loc[0,'content'])
  df2= pd.read_json(json.dumps(dict_obj.get('Data')))
    
  path = get_filenames(migration_folder)
  for filepath in path:
      with xlrd.open_workbook(filepath) as wb:
          print([filepath])
          sheet = wb.sheet_names()
          for filesheet in sheet:
              worksheet = wb.sheet_by_name(filesheet)
              print(filesheet)
              row_data = pd.read_excel(filepath, sheet_name = filesheet, header = 3)
              row_data = row_data[pd.notnull(row_data['Date'])] # Only take records where Date is the latest date
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

              ##Incremental_load
              counter=len(row_data)-len(df2)
              newRow = len(df2)
              for i in range(counter):
                df2 = df2.append(row_data.iloc[newRow+i],ignore_index=True)

              json_df = df2.to_json(orient='records', lines=False, date_format='iso')
              new_df = json.dumps({"Data": json.loads(json_df)})
              print(new_df)

              cursor = conn.cursor()
              sql = "UPDATE [asia_assistance_rpa].[cba].[ops_disbursement_req_master] SET [content] = %s WHERE [year]= %s"
              val = (new_df,year)
              cursor.execute(sql, val)
    
              conn.commit()
              conn.close()
              print('Data migration incremental load for REQ into database - done')

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
