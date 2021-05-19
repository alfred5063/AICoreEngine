#!/usr/bin/python
# FINAL SCRIPT updated as of 3rd Dec 2020
# Workflow - STP-DM

# Declare Python libraries needed for this script
from win32com.client import Dispatch
from utils.Session import session
import pandas as pd
from transformation.dm_file_download import *
from directory.get_filename_from_path import get_file_name
from utils.audit_trail import audit_log
from utils.logging import logging

#source =  r'\\dtisvr2\Shared\stp.dm'
#DM_base = session.base('111', '112','iyliaasyqin.abdmajid@asia-assistance','test download', filename=source)
#password= 'Pib@2019'
#client= 'PACIFIC'
#filename = r'\\dtisvr2\Shared\stp.dm\iyliaasyqin.abdmajid@asia-assistance\New\PIB_TPA_G_28082020_password.xls'
#allpass = ['Pib@2020', 'Pib@2019']
#zip_status = False #return from DM_File_download

#decrypt password excel/zip
def decrypt_password(DM_base, filename, allpass, zip_status):
  try:
      if zip_status == False:
        for password in allpass:
          instance = Dispatch('Excel.Application')
          try:
            xlwb = instance.Workbooks.Open(filename,False, True, None,password)
            instance.Visible = False
            xl_sh = xlwb.Sheets(1)
            print(xl_sh)
            last_row,last_col = get_last_row(DM_base, filename,xl_sh)           
            content = xl_sh.Range(xl_sh.Cells(1, 1), xl_sh.Cells(last_row, last_col)).Value
            df = pd.DataFrame(list(content[1:]), columns=content[0])
            xlwb.Close(False)
            return df
          except Exception as error:
            print(error)
      elif zip_status == True:
        for password in allpass:
          try:
            unzip_file(filename, head, password)
            audit_log("STP-DM - Unzipped file LONPAC." % filename , "Completed...", DM_base)
          except Exception as error:
            logging("STP-DM - Unzipped file LONPAC.", error, DM_base)
  except Exception as error:
    logging("STP-DM - Error in unzipped file LONPAC.", error, DM_base)



def get_last_row(DM_base, filename, xl_sh):

  # Get last_row
  row_num = 0
  cell_val = ''
  while cell_val != None:
      row_num += 1
      cell_val = xl_sh.Cells(row_num, 1).Value
      # print(row_num, '|', cell_val, type(cell_val))
  last_row = row_num - 1
  # print(last_row)

  # Get last_column
  col_num = 0
  cell_val = ''
  while cell_val != None:
      col_num += 1
      cell_val = xl_sh.Cells(1, col_num).Value
      # print(col_num, '|', cell_val, type(cell_val))
  last_col = col_num - 1
  # print(last_col)
  return last_row, last_col


