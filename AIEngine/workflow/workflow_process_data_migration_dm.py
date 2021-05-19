#!/usr/bin/python
# Updated as of 27st April 2020
# Workflow - DCM File Migration Incremental Load Script
# Version 1

# Declare Python libraries needed for this script
import json
import pandas as pd
import xlrd
import numpy as np
import os
import pymssql
from datetime import datetime as dt
from data_migration.Disbursement_Master import *
from connector.connector import *
from configparser import ConfigParser
from connector.dbconfig import read_db_config
from pathlib import Path
from loading.query.query_as_df import *

###Please use this folder
migration_folder = r'\\dtisvr2\CBA_UAT\Result\REQDCA_Paths\1. DISBURSEMENT CLAIMS INVOICE LOG FILE'
year = '2021'
email = 'nurfazreen.ramly@asia-assistance.com'

def process_data_migration_DM_incremental_load(migration_folder, year, email):

  try:
    #read database [cba].[disbursement_master]

    #current_path = os.path.dirname(os.path.abspath(__file__))
    #dti_path = str(Path(current_path).parent)
    #config = read_db_config(dti_path+r'\config.ini', 'mssql')
    #config = read_db_config(r'C:\Users\CHOONLAM.CHIN\Source\Repos\rpa_engine\RPA.DTI\config.ini', 'mssql')
    #serv = config['server']
    #db = config['database']
    #userdb = config['user']
    #pwd = config['password']
    #conn = pymssql.connect(server=serv, user=userdb, password=pwd,
    #                       database=db)
    connector = MSSqlConnector
    query = "SELECT * FROM [asia_assistance_rpa_uat].[cba].[disbursement_master] WHERE [year] = %s"
    params = year
    df1 = mssql_get_df_by_query_without_base(query, params, connector)

    pd.set_option('display.max_columns', None)


    if df1.empty != True:

      dict_obj = json.loads(df1.loc[0,'content'])

      df2= pd.read_json(json.dumps(dict_obj.get('Data')))
      df2 = df2.rename(columns = {'DCA': 'Types of Invoice'})
      df2 = df2[(df2['Types of Invoice'].astype(str).str.contains('Disbursement claims')) | (df2['Types of Invoice'].astype(str).str.contains('DCA'))]
    
      #df2['Date'] = df2['Date'].dt.tz_localize(None)
      #df2.to_excel(r'C:\Users\Asus\Desktop\migration folder\new.xlsx')

      #try to open the workbook, if encrypted then it wil raise error
      file_name = get_filenames(migration_folder)
      path = []
      current_year = dt.now().year
      real = migration_folder + '\\' + str(current_year) + '\\' + 'Disbursement Claims Master ' + str(current_year)
      a = 0
      for a in range(len(file_name)):
        if real in file_name[a]:
          path.append(file_name[a])

      for filepath in path:
          with xlrd.open_workbook(filepath) as wb:
              print([filepath])
              sheet = wb.sheet_names()[0]
              worksheet = wb.sheet_by_name(sheet)     
              row_data = pd.read_excel(filepath, sheet_name = sheet, header = 2)
              row_data = row_data[pd.notnull(row_data['Date'])] # Only take records where Date is the latest date         
              row_data = row_data.rename(columns =
                                          {"Date":"Date",
                                          "DCA":"Types of Invoice",
                                          "Dis. claims no":"Disbursement Claims No",
                                          "Claims listing no":"Claims Listing No",
                                          "No. of cases":"No of Cases",
                                          "File no":"File No",
                                          "Customer Name_Master":"Customer Name Master",
                                          "Hospital":"Hospital",
                                          "Bill no.":"Bill No",
                                          "Patient":"Patient",
                                          "Amounts":"Bill Amount",
                                          "Reasons":"Reasons",
                                          "Action":"Action",
                                          "Initial":"Initial",
                                          "Bank in date":"Bank in Date",
                                          "Cheque no / TT":"Cheque No or TT",
                                          "Amount Paid":"Cheque Amount Paid",
                                          "Remarks":"Remarks",
                                          "Reason for adjustment":"Reason for Adjustment",
                                          "Next action taken":"Next Action Taken",
                                          "New invoice no":"New invoice no",
                                          "Amount":"Invoice Amount ",
                                          "Issue Date":"Issue Date",
                                          "Chq.No.":"Cheque No",
                                          "Amount (RM)":"Settlement Amount",
                                          "OCBC 4":"OCBC 4",
                                          "DB 04":"DB 04",
                                          "DB 10":"DB 10",
                                          "DB 20":"DB 20",
                                          "DB 21":"DB 21",
                                          "DB 39":"DB 39",
                                          "BIMB 1":"BIMB 1",
                                          "DB 03":"DB 03",
                                          "DB 27":"DB 27",
                                          "HSBC 2":"HSBC 2",
                                          "AP Team Remarks":"AP Team Remarks",
                                          "Column1":""
                                          })

              row_data['No of Cases'] = pd.to_numeric(row_data['No of Cases'], downcast='signed').fillna(0)
              row_data['No of Cases'] = row_data['No of Cases'].astype(int)
              df2['No of Cases'] = pd.to_numeric(df2['No of Cases'], downcast='signed').fillna(0)
              df2['No of Cases'] = df2['No of Cases'].astype(int)

              ##Incremental_load
              counter=len(row_data)-len(df2)
              newRow = len(df2)
              for i in range(counter):
                df2 = df2.append(row_data.iloc[newRow+i],ignore_index=True)

              json_df = df2.to_json(orient='records', lines=False, date_format='iso')
              new_df = json.dumps({"Data": json.loads(json_df)})
              print(new_df)
              conn = MSSqlConnector()
              cursor = conn.cursor()

              sql = "UPDATE [cba].[disbursement_master] SET [content] = %s WHERE [year]= %s"
              val = (new_df,year)
              cursor.execute(sql, val)
    
              conn.commit()
              conn.close()
              print('Data migration incremental load for DM into database - done')

    else:
      connector = MSSqlConnector
      query = "SELECT * FROM [asia_assistance_rpa_uat].[cba].[disbursement_master] WHERE [year] = %s"
      params = int(year) - 1
      df1_previous = mssql_get_df_by_query_without_base(query, str(params), connector)

      dict_obj = json.loads(df1_previous.loc[0,'content'])
      df2= pd.read_json(json.dumps(dict_obj.get('Data')))
      df2 = df2.rename(columns = {'DCA': 'Types of Invoice'})
      df2 = df2[(df2['Types of Invoice'].astype(str).str.contains('Disbursement claims')) | (df2['Types of Invoice'].astype(str).str.contains('DCA'))]


      file_name = get_filenames(migration_folder)
      path = []
      current_year = dt.now().year
      real = migration_folder + '\\' + str(current_year) + '\\' + 'Disbursement Claims Master ' + str(current_year)
      a = 0
      for a in range(len(file_name)):
        if real in file_name[a]:
          path.append(file_name[a])

      for filepath in path:
          with xlrd.open_workbook(filepath) as wb:
              print([filepath])
              sheet = wb.sheet_names()[0]
              worksheet = wb.sheet_by_name(sheet)     
              row_data = pd.read_excel(filepath, sheet_name = sheet, header = 2)
              row_data = row_data[pd.notnull(row_data['Date'])] # Only take records where Date is the latest date         
              row_data = row_data.rename(columns =
                                          {"Date":"Date",
                                          "DCA":"Types of Invoice",
                                          "Dis. claims no":"Disbursement Claims No",
                                          "Claims listing no":"Claims Listing No",
                                          "No. of cases":"No of Cases",
                                          "File no":"File No",
                                          "Customer Name_Master":"Customer Name Master",
                                          "Hospital":"Hospital",
                                          "Bill no.":"Bill No",
                                          "Patient":"Patient",
                                          "Amounts":"Bill Amount",
                                          "Reasons":"Reasons",
                                          "Action":"Action",
                                          "Initial":"Initial",
                                          "Bank in date":"Bank in Date",
                                          "Cheque no / TT":"Cheque No or TT",
                                          "Amount Paid":"Cheque Amount Paid",
                                          "Remarks":"Remarks",
                                          "Reason for adjustment":"Reason for Adjustment",
                                          "Next action taken":"Next Action Taken",
                                          "New invoice no":"New invoice no",
                                          "Amount":"Invoice Amount ",
                                          "Issue Date":"Issue Date",
                                          "Chq.No.":"Cheque No",
                                          "Amount (RM)":"Settlement Amount",
                                          "OCBC 4":"OCBC 4",
                                          "DB 04":"DB 04",
                                          "DB 10":"DB 10",
                                          "DB 20":"DB 20",
                                          "DB 21":"DB 21",
                                          "DB 39":"DB 39",
                                          "BIMB 1":"BIMB 1",
                                          "DB 03":"DB 03",
                                          "DB 27":"DB 27",
                                          "HSBC 2":"HSBC 2",
                                          "AP Team Remarks":"AP Team Remarks",
                                          "Column1":""
                                          })

              row_data['No of Cases'] = pd.to_numeric(row_data['No of Cases'], downcast='signed').fillna(0)
              row_data['No of Cases'] = row_data['No of Cases'].astype(int)
              df2['No of Cases'] = pd.to_numeric(df2['No of Cases'], downcast='signed').fillna(0)
              df2['No of Cases'] = df2['No of Cases'].astype(int)

              ##Incremental_load
              counter=len(row_data)-len(df2)
              newRow = len(df2)
              for i in range(counter):
                df2 = df2.append(row_data.iloc[newRow+i],ignore_index=True)

              json_df = df2.to_json(orient='records', lines=False, date_format='iso')
              new_df = json.dumps({"Data": json.loads(json_df)})
              print(new_df)
              #conn = MSSqlConnector()
              #cursor = conn.cursor()
              insert_mssql(new_df, filepath, year, email)
              #sql = "UPDATE [cba].[disbursement_master] SET [content] = %s WHERE [year]= %s"
              #val = (new_df,year)
              #cursor.execute(sql, val)
    
              #conn.commit()
              #conn.close()
              print('Data migration incremental load for DM into database - done')

  except Exception as error:
    print("Exception error occur: " + error)
    
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
