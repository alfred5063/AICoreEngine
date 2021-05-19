import json
import pandas as pd
import xlrd
import numpy as np
import os
import pymssql
from datetime import date

#function to check file extension. Only support xls. and xlsx.and return file path
def excel_file_filter(filename, extensions=['.xls', '.xlsx']):
    return any(filename.endswith(e) for e in extensions)


#function to get file name 
def get_filenames(root):
  filename_list = []
  for path, subdirs, files in os.walk(root):
      for filename in filter(excel_file_filter, files):
          filename_list.append(os.path.join(path, filename))
  return filename_list
    

#function to rename unnamed:_1_ index for multiindex header 
def rename_unnamed(df1):
    for i, columns_old in enumerate(df1.columns.levels):
        columns_new = np.where(columns_old.str.contains('Unnamed'), '', columns_old)
        df1.rename(columns=dict(zip(columns_old, columns_new)), level=i, inplace=True)
    return df1

#function to skip row before keyword 1 : sage
def skip_row(ws, n1, kwd):
    while kwd not in ws.cell_value(n1, 0).lower():
        n1 += 1
    return n1

#function to skip row before keyword 2 : adm 
def skip_row2(ws, n2, kwd2):
    while kwd2 not in ws.cell_value(n2, 0).lower():
        n2 += 1
    return n2

#function to get today's date
def GETDATE():
    today = str(date.today())
    return today

#function to insert data into mssql by query
def insert_mssql(clientid, datajson, clienturl, invtype, year, cmemail):
    conn = pymssql.connect(server='10.147.78.70\dti', user='dti.rpa', password='A@nn@rpa2019',
                           database='asia_assistance_rpa')

    cursor = conn.cursor()
    cursor.execute(''' INSERT INTO [cba].[client_master]
                                   ([year]
                                   ,[client_id]
                                   ,[invoice_type]
                                   ,[content]
                                   ,[createdby]
                                   ,[modifiedby]
                                   ,[createdon]
                                   ,[modifiedon]
                                   ,[url])
                             VALUES
                                   (%s
                                   ,%s
                                   ,%s
                                   ,%s
                                   ,%s
                                   ,%s
                                   ,%s
                                   ,%s
                                   ,%s)
                    ''', (year, clientid, invtype, datajson, cmemail, cmemail, GETDATE(), GETDATE(), clienturl))
    conn.commit()
    conn.close()

