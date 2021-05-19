from transformation.medi_file_manipulation import file_manipulation,all_excel_move_to_archive,create_path,createfolder,move_file,get_all_File
from transformation.medi_medi import Medi_Mos
from transformation.medi_bdx_manipulation import bdx_automation
from transformation.medi_update_dcm import update_to_dcm
from transformation.medi_generate_dc import generate_dc
from pandas.io.json import json_normalize
from utils.Session import session
from utils.guid import get_guid
from utils.audit_trail import audit_log
from utils.logging import logging
from extraction.marc.authenticate_marc_access import get_marc_access
import pandas as pd
import traceback,os
from utils.notification import send
from datetime import datetime
import xlrd

uploadfile = r"C:\Users\CHUNKIT.LEE\Desktop\test"
disbursementMaster = r"C:\Users\CHUNKIT.LEE\Desktop\test\Disbursement Claims Running No 2020.xls"
#disbursementClaim
disbursementClaim = r"C:\Users\CHUNKIT.LEE\Desktop\test\MCLXXXXX.xls"
# Bordereaux Listing
bordereauxListing = r"C:\Users\CHUNKIT.LEE\Desktop\test\AETNA11324-2019-09 WEB.xls"

def medi_mapping(disbursementClaim,bordereauxListing):
  wb = xlrd.open_workbook(disbursementClaim)
  df = pd.read_excel(wb)
  
  df.to_excel(r"C:\Users\CHUNKIT.LEE\Desktop\test\testw.xlsx")
  
  getDC = df.iloc[28, 3]

  wb = xlrd.open_workbook(disbursementMaster)
  
  df1 = pd.read_excel(wb,skiprows=[0])
  new_header = df1.iloc[0] #grab the first row for the header
  df1 = df1[1:] #take the data less the header row
  df1.columns = new_header #set the header row as the df header

  mybordnum = 'TPAAY-0001-202001'
  df1.loc[df1['Bord No'] == '%s' % str(mybordnum)]

  

  for bordno in df1[df1['col4']]:
    if df1[df1['col4']=='TPAAY-0001-202001']:
        print(exist)
    else:
        print(not exist)

  df1.iloc['Bord No']
  


