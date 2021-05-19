#!/usr/bin/python
# FINAL SCRIPT updated as of 9th Dec 2020
# Workflow - REQ/DCA

# Declare Python libraries needed for this script
from automagica import *
import openpyxl
import datetime
import pandas as pd
from openpyxl import Workbook
import string

# Convert column name to column number
def column_to_number(c):
  number=-25
  for l in c:
    if not l in string.ascii_letters:
      return False
    number+=ord(l.upper())-64+25
  return number

# Search the cell position in excel file
def search_excel(keyword, sheet, workbook):
  ws = workbook.worksheets[sheet]
  for row in ws:
    for cell in row:
      if cell.value == keyword:
        cell_no = [cell.column,cell.row]
        return cell_no
        break

# Search the latest row to append data in excel file from dataframe
def search_last_row(cell_no, CMfile, sheet, workbook, keyword):
  df = pd.read_excel(CMfile, header = cell_no[1]-1, sheet_name=sheet)
  main_df = df.loc[:,keyword]
  main_df = main_df.dropna()

  if(main_df.first_valid_index()==None):
    return cell_no[1]+2
  else:
    try:
      return search_excel(main_df[main_df.first_valid_index()], sheet, workbook)[1] + main_df.last_valid_index() - main_df.first_valid_index()
    except Exception:
      return search_excel(str(int(main_df[main_df.first_valid_index()])), sheet, workbook)[1] + main_df.last_valid_index() - main_df.first_valid_index()

# Compare the dataframe of Cashless and DCA, then copy information available in matched header from Cashless dataframe into DCA dataframe
# Return the final dataframe
# Script 1
def compare_df(df1, df2, index):
  col1 = 0
  for col1 in range(len(df1.columns)):
    col2 = 15
    for col2 in range(len(df2.columns)):    
      if df1.columns[col1] == df2.columns[col2]:
        df2.iloc[index, col2] = df1.loc[index, col1]
      elif df2.columns[col2] == 'Bill received':
        df2.loc[index, 'Bill received'] = df1.loc[index, df1.filter(regex = 'OB Registered date|OB Registered Date|Bill received').columns.item()]
      else:
        df2.loc[index, df2.filter(regex = 'Bill No|Plan Type|Bill No.').columns.item()] = df1.loc[index,df1.filter(regex = 'Bill No|Plan Type|Bill No.').columns.item()]
  return df2

# Script 2
def copy_df(df1, df2, index):

  df2.loc[index, 'Adm. No'] = int(df1['Adm. No'][index])
  df2.loc[index, 'File No.'] = int(df1['Adm. No'][index])
  df2.loc[index, 'Disbursement Listing'] = str(df1['Disbursement Listing'][index])
  df2.loc[index, 'Disbursement Claims'] = str(df1['Disbursement Claims'][index])
  df2.loc[index, '''Member's Name'''] = str(df1['''Member's Name'''][index])
  df2.loc[index, 'Policy No'] = str(df1['Policy No'][index])
  df2.loc[index, 'Bill No'] = str(df1['Bill No'][index])
  df2.loc[index, 'Adm. Date'] = df1['Adm. Date'][index]
  df2.loc[index, 'Dis. Date'] = df1['Dis. Date'][index]
  df2.loc[index, 'Hospital'] = str(df1['Hospital'][index])
  df2.loc[index, 'Diagnosis'] = str(df1['Diagnosis'][index])
  df2.loc[index, 'Insurance'] = float(df1['Insurance'][index])
  df2.loc[index, 'Patient'] = float(df1['Patient'][index])
  df2.loc[index, 'Total'] = ''
  df2.loc[index, 'Submission Date'] = ''
  df2.loc[index, 'Bill received'] = df1['OB Received Date'][index].strftime("%d/%m/%Y")
  df2.loc[index, 'CSU Received Date'] = df1['CSU Received Date'][index].strftime("%d/%m/%Y")
  df2.loc[index, 'Remarks'] = ''

  return df2

# Update the information in dataframe into Excel sheet
def update_excel(sheet, cellno, data, path, index):
  columns = list(data)
  ExcelOpenWorkbook(path)
  Wait(4)
  PressKey("f5")
  Wait(2)
  Type(sheet + '!' + cellno, interval_seconds = 0.03)
  PressKey("enter")
  for i in columns:
    if(isinstance(data[i][index], datetime.datetime)):
      Type(data[i][index].strftime("%d/%m/%Y"), interval_seconds = 0.03)
    elif(pd.isna(data[i][index])):
      Type('', interval_seconds = 0.02)
    else:
      Type(str(data[i][index]), interval_seconds = 0.03)
    PressKey("tab")
  PressHotkey("ctrl", "s")
  Wait(3)
  PressHotkey("alt", "f4")
  Wait(3)

def find_sheet(sheet, start_cell):
  Failsafe(False)
  Wait(4)
  PressKey("f5")
  Wait(2)
  Type(sheet + '!' + start_cell, interval_seconds = 0.01)
  PressKey("enter")

def update_entry(sheet, dataframe, start_cell, i, start_column):
  #dataframe['No']
  columns = list(dataframe)
  val = ''
  for j in columns:
    try:
      if isinstance(dataframe[j][i], datetime.datetime) & pd.notnull(dataframe[j][i]):
        val = dataframe[j][i].strftime("%d/%m/%Y")
      elif(pd.isna(dataframe[j][i])):
        val = ''
      else:
        if j == 'No':
          val = int(dataframe[j][i])
        else:
          val = str(dataframe[j][i])
      sheet.cell(row=start_cell, column=start_column).value = val
      start_column = start_column + 1
    except Exception as error:
      start_column = start_column + 1
      continue

def save_and_close(workbook, excel_file):
  workbook.save(excel_file)
  workbook.close()
