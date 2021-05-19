#!/usr/bin/python
# FINAL SCRIPT updated as of 8th April 2020
# Workflow - CBA/MEDICLINIC
# Version 1

# Declare Python libraries needed for this script
import pandas as pd
import numpy as np
from xlrd import open_workbook
import xlrd
import os
import math

def populate_data_medi(disbursementClaim, bordlisting, disbursementMaster, destination, medi_base):

  try:
    DCM_df, BordListing_df, DC_df = medi_mapping(disbursementMaster, bordlisting, disbursementClaim)
    DCM_df.to_excel(destination + 'disbursement master(new).xlsx', index = False)
    BordListing_df.to_excel(destination + "Bord Listing(new).xlsx", index = False, header = False)
    DC_df.to_excel(destination + "Disbursement Claim(new).xlsx", index = False, header = False)
  except Exception as error:
      print('ERROR')

def medi_mapping(disbursementMaster, bordereauxListing, disbursementClaim):
  
  wb = xlrd.open_workbook(bordereauxListing)
  bordlist_df = pd.read_excel(wb)

  dcm_workbook = open_workbook(disbursementMaster)
  dcm_sheet = dcm_workbook.sheet_by_index(0)
  df = pd.read_excel(dcm_workbook)
  
  if df.columns[0] != 'Date':
      dcm_df = pd.read_excel(disbursementMaster, sheet_by_index = 0, skiprows = 2, usecols = list(range(dcm_sheet.ncols + 1)))
  else:
      dcm_df = pd.read_excel(disbursementMaster, sheet_by_index = 0, skiprows = 0, usecols = list(range(dcm_sheet.ncols + 1)))

  newRunnningNo = dcm_df.iloc[get_DCM_fill_index(disbursementMaster), 1]
  newRowIndex = get_DCM_fill_index(disbursementMaster)

  # Perform update data into new row in Disbursement Claim Master file
  print("- Perform update data into new row in Disbursement Claim Master file.")
 
  data2 = pd.read_excel(bordereauxListing, skiprows = 11)
  totalCases = get_number_cases_price(bordereauxListing)
  price = data2.loc[get_number_cases_price(bordereauxListing), 'Aetna Amount']
  initial = data2.loc[get_Initial_index(bordereauxListing), 'Aetna Amount']
  ##Mapping
  bordlist_df.iloc[5,2] = newRunnningNo
  dcm_df.loc[newRowIndex, 'Date'] = bordlist_df.iloc[3, 2]
  dcm_df.loc[newRowIndex, 'Bord No'] = bordlist_df.iloc[2, 2]
  dcm_df.loc[newRowIndex, 'Corporate'] = bordlist_df.iloc[0, 2]
  dcm_df.loc[newRowIndex, 'Amount (RM) \n(Kindly put (RM0.00) for CN)'] = price
  dcm_df.loc[newRowIndex, 'Initial'] = initial
  dcm_df.loc[newRowIndex, 'Total no of cases bord'] = totalCases

  ###Populate Disbursement Claim
  wb = xlrd.open_workbook(disbursementClaim)
  dc_df = pd.read_excel(wb)

  dc_df.iloc[18, 3] = newRunnningNo
  dc_df.iloc[18, 8] = dcm_df.loc[newRowIndex, 'Date']
  dc_df.iloc[28, 3] = dcm_df.loc[newRowIndex, 'Bord No']
  dc_df.iloc[28, 0] = dcm_df.loc[newRowIndex, 'Corporate']
  dc_df.iloc[28, 7] = price
  dc_df.iloc[54, 6] = initial


  return dcm_df, bordlist_df, dc_df
 
def get_DCM_fill_index(disbursementMaster):
  dcm_workbook = open_workbook(disbursementMaster)
  dcm_sheet = dcm_workbook.sheet_by_index(0)
  df = pd.read_excel(dcm_workbook)
  if df.columns[0] != 'Date':
      data=pd.read_excel(disbursementMaster, skiprows = 2 , na_values = "Missing")
  else:
      data=pd.read_excel(disbursementMaster, skiprows = 0 , na_values = "Missing")

  Bord_No_list = pd.DataFrame(data, columns=['Bord No']).values.tolist()
  counter=len(Bord_No_list)-1
  try:
    while True:
      math.isnan(Bord_No_list[counter][0])
      counter-=1
  except:
    a=None
  fill_index=counter+1
  return fill_index

def get_number_cases_price(bordereauxListing):
  data=pd.read_excel(bordereauxListing,skiprows = 11 , na_values = "Missing")
  Diagnosis_Description_list = pd.DataFrame(data, columns=['Diagnosis Description [Code]']).values.tolist()
  counter=len(Diagnosis_Description_list)-1
  try:
    while True:
      math.isnan(Diagnosis_Description_list[counter][0])
      counter-=1
  except:
    a=None
  fill_index=counter+1
  return fill_index

def get_Initial_index(bordereauxListing):
  data=pd.read_excel(bordereauxListing,skiprows = 11 , na_values = "Missing")
  Anetna_Amount_Column_list = pd.DataFrame(data, columns=['Aetna Amount']).values.tolist()
  counter=len(Anetna_Amount_Column_list)-1
  try:
    while True:
      math.isnan(Anetna_Amount_Column_list[counter][0])
      counter-=1
  except:
    a=None
  fill_index=counter-1
  return fill_index

