
import json
import pandas as pd
import xlrd
import numpy as np
import os
import pymssql
from datetime import date
import requests
from pathlib import Path
from os.path import join, dirname, abspath, isfile
from collections import Counter
from xlrd.sheet import ctype_text   

fname_address = r"C:\Users\CHUNKIT.LEE\Desktop\excel\All Client Address - Cleaned.xls"

def get_excel_sheet_object(fname_address):

    # Declare a new dataframe
    insurer_detailsdf = pd.DataFrame()

    #Check file exist or not
    if isfile(fname_address) == True:
        print("file exists")
    else:
        print('File doesnt exist: ', fname_address)

    # Open the workbook 
    wb = xlrd.open_workbook(fname_address)
    profile_data_list = []

    sheet_names = wb.sheet_names()
    print (20 * '-' + 'Retrieved worksheet: %s' %sheet_names)

    '''worksheet = wb.sheet_by_name('ALG')
    num_rows = worksheet.nrows - 1
    curr_row = -1
    while curr_row < num_rows:
            curr_row += 1
            row = worksheet.row(curr_row)
            print(row)
            '''


    for s in wb.sheets():
      try:
        for row in range(1,10):
          if row is not None:
            values = []
            for column in range(s.ncols):
              values.append(str(s.cell(row, column).value))
            profile_data_list.append(values)
          else:
            pass
      except Exception:
        continue
   
    df = pd.DataFrame(profile_data_list,columns=["id","insurer","atten_1", "atten_2", "atten_3", "comp_department","comp_name","comp_address","billing_address"])

    i = 0
    x = 0
    y = 9
    for i in range(len(df)):
        row_df = df[x:y]
      
        # Remove all empty rows based on column insurer
        row_df = row_df[row_df.insurer != ''].reset_index().drop(columns = ['id'])
        insurer_detailsdf = pd.concat([insurer_detailsdf,row_df], ignore_index = True, sort = False)
        #insurer_detailsdf = insurer_detailsdf.append(row_df, ignore_index = True, sort = False)
        #print(insurer_detailsdf['insurer'])  

        # Clear the none in other rows. Save as insurer_detailsdf
      
        #insurer_detailsdf = insurer_detailsdf.dropna()
        #insurer_detailsdf = insurer_detailsdf.dropna( axis=0,  how="all").drop(insurer_detailsdf.index[0])
        #insurer_detailsdf.to_csv(r'C:\Users\CHUNKIT.LEE\Desktop\results\insurer_detailsdf.csv')
      
        x = x + 9
        y = y + 9

    insurer_detailsdf.to_csv(r'C:\Users\CHUNKIT.LEE\Desktop\results\insurer_detailsdf.csv')

