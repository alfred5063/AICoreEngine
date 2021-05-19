#!/usr/bin/python
# FINAL SCRIPT updated as of 22th May 2020
# Workflow - Finance SOA
# Version 1

# Declare Python libraries needed for this script
import re
import shutil
import sys
import os
import glob
import pandas as pd
from datetime import datetime
from utils.audit_trail import audit_log
from utils.logging import logging

def soa_mapping(soa_df, soarpa_base):

  # Check if required Bill Date columns exist
  # To filter the selected and similar columns
  check_cols = [col for col in soa_df.columns if 'Date' in col]
  filter1_cols = [col for col in check_cols if 'Bill' in col]
  filter2_cols = [col for col in check_cols if 'Invoice' in col]
  filter3_cols = [col for col in check_cols if 'Document' in col]
  filter4_cols = [col for col in check_cols if 'Doc' in col]

  # Create a list to store the data and compare with system excel columns
  filter_list = list()
  filter_value = ""

  try:
    if filter1_cols != [] :
      filter_value = filter1_cols
      filter_list.append(filter_value)
    elif filter2_cols != [] :
      filter_value = filter2_cols
      filter_list.append(filter_value)
    elif filter3_cols != [] :
      filter_value = filter3_cols
      filter_list.append(filter_value)
    elif filter4_cols != [] :
      filter_value = filter4_cols
      filter_list.append(filter_value)
    else:
      pass
  except Exception as error:
    logging("Similar Bill Date Column not found", error, soarpa_base)
    print(error)

  # To remove the special characters
  delete_char = "[\'\']"

  table = str(filter_list).maketrans('','',delete_char)

  bill_date = str(filter_list).translate(table)

  # To rename the columns according to the mapping excel file
  try:
    if pd.Series(bill_date).isin(soa_df.columns).all():
      soa_df = soa_df.rename(columns = {bill_date: "Bill Date"})
    else:
      pass
  except Exception as error:
    logging("Bill Date Column does not exist", error, soarpa_base)

  # Check if required Invoice No columns exist
  # To filter the selected and similar columns
  check1_cols = [col for col in soa_df.columns if 'Invoice' in col]
  check2_cols = [col for col in soa_df.columns if 'Bill' in col]
  check3_cols = [col for col in soa_df.columns if 'No' in col]
  filter5_cols = [col for col in check1_cols if 'ID' in col]
  filter6_cols = [col for col in check1_cols if 'No' in col]
  filter7_cols = [col for col in check2_cols if 'No' in col]
  filter8_cols = [col for col in check2_cols if 'Reference' in col]
  filter9_cols = [col for col in check3_cols if 'Document' in col]
  filter10_cols = [col for col in check3_cols if 'Ref' in col]
  filter11_cols = [col for col in check3_cols if 'A/C' in col]
  filter12_cols = [col for col in check3_cols if 'Doc' in col]
  filter13_cols = [col for col in check3_cols if 'Visit' in col]

  # Create a list to store the data and compare with system excel columns
  filter_list = list()
  filter_value = ""

  try:
    if filter5_cols != [] :
      filter_value = filter5_cols
      filter_list.append(filter_value)
    elif filter6_cols != [] :
      filter_value = filter6_cols
      filter_list.append(filter_value)
    elif filter7_cols != [] :
      if len(filter7_cols) >= 2:
        filter_value = filter7_cols[0]
      else:
        filter_value = filter7_cols
      filter_list.append(filter_value)
    elif filter8_cols != [] :
      filter_value = filter8_cols
      filter_list.append(filter_value)
    elif filter9_cols != [] :
      filter_value = filter9_cols
      filter_list.append(filter_value)
    elif filter10_cols != [] :
      filter_value = filter10_cols
      filter_list.append(filter_value)
    elif filter11_cols != [] :
      filter_value = filter11_cols
      filter_list.append(filter_value)
    elif filter12_cols != [] :
      filter_value = filter12_cols
      filter_list.append(filter_value)
    elif filter13_cols != [] :
      filter_value = filter13_cols
      filter_list.append(filter_value)
    else :
      pass
  except Exception as error:
    logging("Similar Invoice No Column not found", error, soarpa_base)
    print(error)

  # To remove the special characters
  delete_char = "[\'\']"

  table = str(filter_list).maketrans('','',delete_char)

  invoice_no = str(filter_list).translate(table)

  # To rename the columns according to the mapping excel file
  try:
    if pd.Series(invoice_no).isin(soa_df.columns).all() == True:
      soa_df = soa_df.rename(columns = {invoice_no: "Invoice No"})
    else:
      pass
  except Exception as error:
    logging("Invoice No Column does not exist", error, soarpa_base)


  # To check whether the columns name is exist in the soa excel columns or not
  # To filter the selected and similar columns
  check4_cols = [col for col in soa_df.columns if 'Name' in col]
  filter14_cols = [col for col in check4_cols if 'Patient' in col]
  filter15_cols = [col for col in soa_df.columns if 'Policyholder' in col]
  filter16_cols = [col for col in soa_df.columns if 'Particulars' in col]
  filter17_cols = [col for col in check4_cols if 'Member' in col]

  # Create a list to store the data and compare with system excel columns
  filter_list = list()
  filter_value = ""

  try:
    if filter14_cols != [] :
      filter_value = filter14_cols
      filter_list.append(filter_value)
    elif filter15_cols != [] :
      filter_value = filter15_cols
      filter_list.append(filter_value)
    elif filter16_cols != [] :
      filter_value = filter16_cols
      filter_list.append(filter_value)
    elif filter17_cols != [] :
      filter_value = filter17_cols
      filter_list.append(filter_value)
    else :
      pass
  except Exception as error:
    logging("Similar Patient Name Column not found", error, soarpa_base)

  # To remove the special characters
  delete_char = "[\'\']"

  table = str(filter_list).maketrans('','',delete_char)

  patient_name = str(filter_list).translate(table)

  # To rename the columns according to the mapping excel file
  try:
    if pd.Series(patient_name).isin(soa_df.columns).all() == True:
      soa_df = soa_df.rename(columns = {patient_name: "Patient Name"})
    else:
      pass
  except Exception as error:
    logging("Patient Name Column does not exist", error, soarpa_base)

  # Check if required GL No columns exist
  # To filter the selected and similar columns
  check5_cols = [col for col in soa_df.columns if 'GL' in col]
  check6_cols = [col for col in soa_df.columns if 'No' in col]
  filter27_cols = [col for col in soa_df.columns if 'GL No' in col]
  filter18_cols = [col for col in check5_cols if 'Ref' in col]
  filter19_cols = [col for col in check5_cols if 'Number' in col]
  filter20_cols = [col for col in check5_cols if 'Reference' in col]
  filter21_cols = [col for col in check5_cols if 'INS' in col]
  filter22_cols = [col for col in check6_cols if 'File' in col]
  filter23_cols = [col for col in check6_cols if 'Claim' in col]
  filter24_cols = [col for col in check6_cols if 'Serial' in col]
  filter25_cols = [col for col in check6_cols if 'Policy' in col]
  filter26_cols = [col for col in check6_cols if 'Episode' in col]


  # Create a list to store the data and compare with system excel columns
  filter_list = list()
  filter_value = ""

  try:
    if filter27_cols != [] :
      filter_value = filter27_cols
      filter_list.append(filter_value)
    elif filter18_cols != [] :
      filter_value = filter18_cols
      filter_list.append(filter_value)
    elif filter19_cols != [] :
      filter_value = filter19_cols
      filter_list.append(filter_value)
    elif filter20_cols != [] :
      filter_value = filter20_cols
      filter_list.append(filter_value)
    elif filter21_cols != [] :
      filter_value = filter21_cols
      filter_list.append(filter_value)
    elif filter22_cols != [] :
      filter_value = filter22_cols
      filter_list.append(filter_value)
    elif filter23_cols != [] :
      filter_value = filter23_cols
      filter_list.append(filter_value)
    elif filter24_cols != [] :
      filter_value = filter24_cols
      filter_list.append(filter_value)
    elif filter25_cols != [] :
      filter_value = filter25_cols
      filter_list.append(filter_value)
    elif filter26_cols != [] :
      filter_value = filter26_cols
      filter_list.append(filter_value)
    else:
      pass
  except Exception as error:
    logging("Similar GL No Column not found", error, soarpa_base)

  # To remove the special characters
  delete_char = "[\'\']"

  table = str(filter_list).maketrans('','',delete_char)

  gl_no = str(filter_list).translate(table)

  # To rename the columns according to the mapping excel file
  try:
    if pd.Series(gl_no).isin(soa_df.columns).all() == True:
      soa_df = soa_df.rename(columns = {gl_no: "GL No"})
    else:
      pass
  except Exception as error:
    logging("GL No Column does not exist", error, soarpa_base)

  # Check if required Invoice Amount columns exist
  # To filter the selected and similar columns
  check7_cols = [col for col in soa_df.columns if 'Amount' in col]
  check8_cols = [col for col in soa_df.columns if 'Amt' in col]
  filter28_cols = [col for col in soa_df.columns if 'Value' in col]
  filter29_cols = [col for col in soa_df.columns if 'Balance' in col]
  filter30_cols = [col for col in soa_df.columns if 'Debit' in col]
  filter31_cols = [col for col in soa_df.columns if 'Total' in col]
  filter32_cols = [col for col in soa_df.columns if 'Credit' in col]
  filter33_cols = [col for col in soa_df.columns if 'Size' in col]
  filter34_cols = [col for col in check7_cols if 'Outstanding' in col]
  filter35_cols = [col for col in check7_cols if 'O/S' in col]
  filter36_cols = [col for col in check7_cols if 'Invoice' in col]
  filter37_cols = [col for col in check8_cols if 'OS' in col]
  filter38_cols = [col for col in check8_cols if 'Net' in col]
  filter39_cols = [col for col in check8_cols if 'Unsettle' in col]
  filter40_cols = [col for col in check8_cols if 'Bill' in col]

  # Create a list to store the data and compare with system excel columns
  filter_list = list()
  filter_value = ""

  try:
    if filter28_cols != [] :
      filter_value = filter28_cols
      filter_list.append(filter_value)
    elif filter29_cols != [] :
      filter_value = filter29_cols
      filter_list.append(filter_value)
    elif filter30_cols != [] :
      filter_value = filter30_cols
      filter_list.append(filter_value)
    elif filter31_cols != [] :
      filter_value = filter31_cols
      filter_list.append(filter_value)
    elif filter32_cols != [] :
      filter_value = filter32_cols
      filter_list.append(filter_value)
    elif filter33_cols != [] :
      filter_value = filter33_cols
      filter_list.append(filter_value)
    elif filter34_cols != [] :
      filter_value = filter34_cols
      filter_list.append(filter_value)
    elif filter35_cols != [] :
      filter_value = filter35_cols
      filter_list.append(filter_value)
    elif filter36_cols != [] :
      filter_value = filter36_cols
      filter_list.append(filter_value)
    elif filter37_cols != [] :
      filter_value = filter37_cols
      filter_list.append(filter_value)
    elif filter38_cols != [] :
      filter_value = filter38_cols
      filter_list.append(filter_value)
    elif filter39_cols != [] :
      filter_value = filter39_cols
      filter_list.append(filter_value)
    elif filter40_cols != []:
      filter_value = filter40_cols
      filter_list.append(filter_value)
    else:
      pass
  except Exception as error:
    logging("Similar Invoice Amount Column not found", error, soarpa_base)

  # To remove the special characters
  delete_char = "[\'\']"

  table = str(filter_list).maketrans('','',delete_char)

  invoice_amount = str(filter_list).translate(table)

  # To rename the columns according to the mapping excel file
  try:
    if pd.Series(invoice_amount).isin(soa_df.columns).all() == True:
      soa_df = soa_df.rename(columns = {invoice_amount: "Invoice Amount"})
    else:
      pass
  except Exception as error:
    logging("GL No Column does not exist", error, soarpa_base)

  return soa_df.columns
