import json
import pandas as pd
import xlrd
import numpy as np
import os
import pymssql
from datetime import date
from utils.guid import get_guid
from utils.audit_trail import audit_log
from utils.logging import logging
from utils.notification import send
import requests
from utils.notification import send
from data_migration.Client_Master import *

def process_data_migration_CM(migration_folder,clientId, year, email):
  path = get_filenames(migration_folder)
  for filepath in path:
    #try to open the workbook, if encrypted then it wil raise error
    try:
      with xlrd.open_workbook(filepath) as wb:
          print([filepath])
          sheet = wb.sheet_names()
          #start check throughout file tabs
          for filesheet in sheet:
              worksheet = wb.sheet_by_name(filesheet)
              print(filesheet)
              countkwd = 0
              countkwd2 = 0
              keyword = 'sage'
              keyword2 = 'adm'
              #start count how many rows to skip before keyword 1: sage
              try:
                  skip = skip_row(worksheet, countkwd, keyword)
              except:
                  #if sage couldn't found, start count how many rows to skip before keyword 2: adm 
                  try:
                      skip = skip_row2(worksheet, countkwd2, keyword2)
                  except:
                      continue
              row_data = pd.read_excel(filepath, sheet_name=filesheet, skiprows=skip, header=[0, 1])
              try:
                  rename_unnamed(row_data) 
                  test = row_data.loc[:, ~row_data.columns.duplicated()] #remove column's duplicated
                  test.columns = test.columns.map(' '.join) #
              except Exception as error:
                  print ("Exception error occur:" + error +", please check your excel structure")
                  continue
              df = test.to_json(orient='records', lines=False, date_format='iso')
              new_df = json.dumps({"Data":json.loads(df)})
              print(new_df)
              insert_mssql(clientId, new_df, filepath, filesheet, year, email)
    except Exception as error:
      print("Exception error occur: " + error)
      continue
