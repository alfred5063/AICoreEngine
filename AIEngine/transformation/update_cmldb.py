#!/usr/bin/python
# FINAL SCRIPT updated as of 16th June 2020
# Workflow - Finance Update CML
# Version 1

# Declare Python libraries needed for this script
import openpyxl
import time
from datetime import datetime as dt
import pandas as pd
import json as json
from openpyxl import Workbook
from transformation.reqdca_excel_manipulation import *
from transformation.reqdca_locate_path import *
from loading.excel.checkExcelLoading import *
from utils.audit_trail import audit_log
from utils.notification import send
from utils.logging import logging
from connector.connector import MySqlConnector, MSSqlConnector
from connector.dbconfig import *
import pyodbc as db
from loading.query.query_as_df import *

def update_cmldb(mydataframe, sheet, client_id, cmlrpa_base):

  # Determine the current year and date
  current_year = dt.now().year
  current_date = time.strftime("%Y-%m-%d")

  # Database connection
  print("- Connecting to database")
  conn = MSSqlConnector()
  cur = conn.cursor()
  conn_mysql = MySqlConnector()
  cur_mysql = conn_mysql.cursor()
  final_df2 = pd.DataFrame()

  try:

    my_dict = []
    my_dict = {
      "SAGE DOC ID": "", "Adm. No": "", "File No.": "", "Disbursement Listing": "", "Disbursement Claims": "", "Member's Name": "",
      "Policy No": "", "Plan Type": "", "Adm. Date": "", "Dis. Date": "", "Hospital": "", "Diagnosis": "", "Insurance": "", "Patient": "",
      "Discrepancy": "", "Total": "", "Submission Date": "", "OB Received Date": "", "OB Registered Date": "", "CSU Received Date": "",
      "Primary GL No": "", "Remarks": "", "Payment Due Date": "", "Aging (days)": "", "SCAN (TAT)": "", "CSU (TAT)": "", "DB BANK CODE": "",
      "Under / (Over) Claimed fr Client (AA-J-AM+AP)": "", "Cheque DetailsBank-in Date": "", "Cheque DetailsChq.No.": "",
      "Cheque DetailsAmount (RM)": "", "Hospital Bills DetailsDate": "", "Hospital Bills DetailsBill No.": "", "Hospital Bills DetailsAmount (RM)": "",
      "Under / (Over) Reimbursed by Client(AB-Y-AN+AQ)": "", "UP/OP CASES": "", "AAN Settlement DetailsDate": "", "AAN Settlement DetailsChq.No.": "",
      "AAN Settlement DetailsAmount (RM)": "", "Under / (Over) Paid to Hosp (AB - AG)": "", "Designated  / Borrowing Bank AccountOCBC 16": "",
      "Designated  / Borrowing Bank AccountBIMB 1": "", "Designated  / Borrowing Bank AccountDBMY29": "",
      "Designated  / Borrowing Bank AccountDBMY29.1": "", "Designated  / Borrowing Bank AccountDBMY29.2": "",
      "Designated  / Borrowing Bank AccountDBMY29.3": "", "Designated  / Borrowing Bank AccountDBMY29.4": "", "Designated  / Borrowing Bank AccountBank1": "",
      "Designated  / Borrowing Bank AccountBank2": "", "Designated  / Borrowing Bank AccountBank3": "", "Designated  / Borrowing Bank AccountBank4": "",
      "Designated  / Borrowing Bank AccountBank5": "", "Designated  / Borrowing Bank AccountBank6": "", "Designated  / Borrowing Bank AccountBank7": "", "TPA Billing DetailsDate": "",
      "TPA Billing DetailsInv. No.": "", "TPA Billing DetailsAmount (RM)": "", "Reimbursement DetailsDate": "", "Reimbursement DetailsChq. No.": "",
      "Reimbursement DetailsAmount (RM)": "", "Adjustment DetailsDate": "", "Adjustment DetailsDoc. No.": "", "Adjustment DetailsAmount (RM)": "",
      "Hospital OR DetailsOR Date": "", "Hospital OR DetailsOR No.": "", "Hospital OR DetailsAmount (RM)": "", "Hospital OR DetailsDate of OR Rec'd": "",
      "Hospital OR DetailsSubm. Date": "", "Tax Invoice DetailsDate": "", "Tax Invoice DetailsInv No": "", "Tax Invoice DetailsCase Fee": "",
      "Tax Invoice DetailsGST Amount": "", "Tax Invoice DetailsTotal": ""
      }

    my_cm = { 'Data': []}

    my_dict["SAGE DOC ID"] = mydataframe.iloc[b]['Sage ID']
    my_dict["Adm. No"] = int(mydataframe.iloc[b]['Adm. No'])
    my_dict["File No."] = int(mydataframe.iloc[b]['File No.'])
    my_dict["Disbursement Listing"] = str(mydataframe.iloc[b]['Disbursement Listing'])
    my_dict["Disbursement Claims"] = mydataframe.iloc[b]['Disbursement Claims']
    my_dict['''Member's Name'''] = mydataframe.iloc[b]['''Member's Name''']
    my_dict["Policy No"] = mydataframe.iloc[b]['Policy No']
    my_dict["Plan Type"] = ''
    my_dict["Adm. Date"] = mydataframe.iloc[b]['Adm. Date']
    my_dict["Dis. Date"] = mydataframe.iloc[b]['Dis. Date']
    my_dict["Hospital"] = mydataframe.iloc[b]['Hospital']
    my_dict["Diagnosis"] = mydataframe.iloc[b]['Diagnosis']
    my_dict["Insurance"] = float(mydataframe.iloc[b]['Insurance'])
    my_dict["Patient"] = float(mydataframe.iloc[b]['Patient'])
    my_dict["Discrepancy"] = ''
    my_dict["Total"] = float(mydataframe.iloc[b]['Total'])
    my_dict["Submission Date"] = mydataframe.iloc[b]['Submission Date']
    my_dict["OB Received Date"] = mydataframe.iloc[b]['OB Received Date']
    my_dict["OB Registered Date"] = mydataframe.iloc[b]['OB Registered Date']
    my_dict["CSU Received Date"] = mydataframe.iloc[b]['CSU Received Date']
    my_dict["Primary GL No"] = ''
    my_dict["Remarks"] = mydataframe.iloc[b]['Remarks']

    # start
    my_dict["Payment Due Date"] = mydataframe.iloc[b]['Payment Due Date']
    my_dict["Aging (days)"] = mydataframe.iloc[b]['Aging (days)']
    my_dict["SCAN (TAT)"] = mydataframe.iloc[b]['SCAN (TAT)']
    my_dict["CSU (TAT)"] = mydataframe.iloc[b]['CSU (TAT)']

    my_dict["DB BANK CODE"] = ''

    my_dict["Under / (Over) Claimed fr Client (AA-J-AM+AP)"] = float(mydataframe.iloc[b]['Under / (Over) Claimed fr Client (AA-J-AM+AP)'])
    my_dict["Cheque DetailsBank-in Date"] = mydataframe.iloc[b]['Cheque Details - Bank-in Date']
    my_dict["Cheque DetailsChq.No."] = mydataframe.iloc[b]['Cheque Details - Chq.No.']
    my_dict["Cheque DetailsAmount (RM)"] = mydataframe.iloc[b]['Cheque Details - Amount(RM)']
    my_dict["Hospital Bills DetailsDate"] = mydataframe.iloc[b]['Hospital Bills Details - Date']
    my_dict["Hospital Bills DetailsBill No."] = mydataframe.iloc[b]['Hospital Bills Details - Bill No.']
    my_dict["Hospital Bills DetailsAmount (RM)"] = mydataframe.iloc[b]['Hospital Bills Details - Amount(RM)']
    my_dict["Under / (Over) Reimbursed by Client(AB-Y-AN+AQ)"] = mydataframe.iloc[b]['Under / (Over) Reimbursed by Client (AB-Y-AN+AQ)']
    my_dict["UP/OP CASES"] = mydataframe.iloc[b]['UP/OP CASES']
    my_dict["AAN Settlement DetailsDate"] = mydataframe.iloc[b]['AAN Settlement Details - Date']
    my_dict["AAN Settlement DetailsChq.No."] = mydataframe.iloc[b]['AAN Settlement Details - Chq.No.']
    my_dict["AAN Settlement DetailsAmount (RM)"] = mydataframe.iloc[b]['AAN Settlement Details - Amount(RM)']
    my_dict["Under / (Over) Paid to Hosp (AB - AG)"] = mydataframe.iloc[b]['Under / (Over) Paid to Hosp (AB - AG)']

    my_dict["Designated  / Borrowing Bank AccountOCBC 16"] = mydataframe.iloc[b]['DBMY21']
    my_dict["Designated  / Borrowing Bank AccountBIMB 1"] = mydataframe.iloc[b]['Designated / Borrowing Bank Account - Bank 2']
    my_dict["Designated  / Borrowing Bank AccountDBMY29"] = mydataframe.iloc[b]['Designated / Borrowing Bank Account - Bank 3']
    my_dict["Designated  / Borrowing Bank AccountDBMY29.1"] = mydataframe.iloc[b]['Designated / Borrowing Bank Account - Bank 4']
    my_dict["Designated  / Borrowing Bank AccountDBMY29.2"] = mydataframe.iloc[b]['Designated / Borrowing Bank Account - Bank 5']
    my_dict["Designated  / Borrowing Bank AccountDBMY29.3"] = mydataframe.iloc[b]['Designated / Borrowing Bank Account - Bank 6']
    my_dict["Designated  / Borrowing Bank AccountDBMY29.4"] = mydataframe.iloc[b]['Designated / Borrowing Bank Account - Bank 7']

    #my_dict["Designated  / Borrowing Bank AccountBank1"] = mydataframe.iloc[b]['DBMY21']
    #my_dict["Designated  / Borrowing Bank AccountBank2"] = mydataframe.iloc[b]['Designated / Borrowing Bank Account - Bank 2']
    #my_dict["Designated  / Borrowing Bank AccountBank3"] = mydataframe.iloc[b]['Designated / Borrowing Bank Account - Bank 3']
    #my_dict["Designated  / Borrowing Bank AccountBank4"] = mydataframe.iloc[b]['Designated / Borrowing Bank Account - Bank 4']
    #my_dict["Designated  / Borrowing Bank AccountBank5"] = mydataframe.iloc[b]['Designated / Borrowing Bank Account - Bank 5']
    #my_dict["Designated  / Borrowing Bank AccountBank6"] = mydataframe.iloc[b]['Designated / Borrowing Bank Account - Bank 6']
    #my_dict["Designated  / Borrowing Bank AccountBank7"] = mydataframe.iloc[b]['Designated / Borrowing Bank Account - Bank 7']

    my_dict["TPA Billing DetailsDate"] = mydataframe.iloc[b]['TPA Billing Details - Date']
    my_dict["TPA Billing DetailsInv. No."] = mydataframe.iloc[b]['TPA Billing Details - Inv. No.']
    my_dict["TPA Billing DetailsAmount (RM)"] = mydataframe.iloc[b]['TPA Billing Details - Amount(RM)']
    my_dict["Reimbursement DetailsDate"] = mydataframe.iloc[b]['Reimbursement Details - Date']
    my_dict["Reimbursement DetailsChq. No."] = mydataframe.iloc[b]['Reimbursement Details - Chq.No.']
    my_dict["Reimbursement DetailsAmount (RM)"] = mydataframe.iloc[b]['Reimbursement Details - Amount(RM)']
    my_dict["Adjustment DetailsDate"] = mydataframe.iloc[b]['Adjustment Details - Date']
    my_dict["Adjustment DetailsDoc. No."] = mydataframe.iloc[b]['Adjustment Details - Doc.No']
    my_dict["Adjustment DetailsAmount (RM)"] = mydataframe.iloc[b]['Adjustment Details - Amount(RM)']
    my_dict["Hospital OR DetailsOR Date"] = mydataframe.iloc[b]['Hospital OR Details - OR Date']
    my_dict["Hospital OR DetailsOR No."] = mydataframe.iloc[b]['Hospital OR Details - OR No.']
    my_dict["Hospital OR DetailsAmount (RM)"] = mydataframe.iloc[b]['Hospital OR Details - Amount(RM)']
    my_dict["Hospital OR DetailsDate of OR Rec'd"] = ''
    my_dict["Hospital OR DetailsSubm. Date"] = ''
    my_dict["Tax Invoice DetailsDate"] = mydataframe.iloc[b]['Tax Invoice Details - Date']
    my_dict["Tax Invoice DetailsInv No"] = mydataframe.iloc[b]['Tax Invoice Details - Inv No']
    my_dict["Tax Invoice DetailsCase Fee"] = mydataframe.iloc[b]['Tax Invoice Details - Case Fee']
    my_dict["Tax Invoice DetailsGST Amount"] = mydataframe.iloc[b]['Tax Invoice Details - GST Amount']
    my_dict["Tax Invoice DetailsTotal"] = mydataframe.iloc[b]['Tax Invoice Details - Total']

    param = (current_year, client_id, sheet)
    sql_query_insurer = "SELECT content FROM cba.client_master WHERE year = %s and client_id = %s and invoice_type = %s"
    cur.execute(sql_query_insurer, param)
    myquery = pd.DataFrame(cur.fetchall())

    if myquery.empty != True:

      # Return the SQL result as a JSON dictionary
      sql_query_insurer = "SELECT content FROM cba.client_master WHERE year = %s and client_id = %s and invoice_type = %s"
      cur.execute(sql_query_insurer, param)
      dictionary_db = json.loads(cur.fetchall()[0][0])

      c = 0
      for c in range(len(dictionary_db['Data'])):

        # 3 - filter based on file no. , disbursement listing, disbursement claims, policy no, bill no, adm. date
        if dictionary_db['Data'][c]['Adm. No'] == mydataframe.iloc[b]['Adm. No'] and dictionary_db['Data'][c]['Disbursement Listing'] == mydataframe.iloc[b]['Disbursement Claims']:

          # Replace the element based on the right record
          dictionary_db['Data'][c]["Payment Due Date"] = dictionary_db['Data'][c]["Payment Due Date"].replace(dictionary_db['Data'][c]["Payment Due Date"], mydataframe.iloc[b]['Payment Due Date'])
          dictionary_db['Data'][c]["Aging (days)"] = dictionary_db['Data'][c]["Aging (days)"].replace(dictionary_db['Data'][c]["Aging (days)"], mydataframe.iloc[b]['Aging (days)'])
          dictionary_db['Data'][c]["SCAN (TAT)"] = dictionary_db['Data'][c]["SCAN (TAT)"].replace(dictionary_db['Data'][c]["SCAN (TAT)"], mydataframe.iloc[b]['SCAN (TAT)'])
          dictionary_db['Data'][c]["CSU (TAT)"] = dictionary_db['Data'][c]["CSU (TAT)"].replace(dictionary_db['Data'][c]["CSU (TAT)"], mydataframe.iloc[b]['CSU (TAT)'])
          dictionary_db['Data'][c]["DB BANK CODE"] = dictionary_db['Data'][c]["DB BANK CODE"].replace(dictionary_db['Data'][c]["DB BANK CODE"], mydataframe.iloc[b]['DB BANK CODE'])
          dictionary_db['Data'][c]["Under / (Over) Claimed fr Client (AA-J-AM+AP)"] = dictionary_db['Data'][c]["Under / (Over) Claimed fr Client (AA-J-AM+AP)"].replace(dictionary_db['Data'][c]["Under / (Over) Claimed fr Client (AA-J-AM+AP)"], mydataframe.iloc[b]['Under / (Over) Claimed fr Client (AA-J-AM+AP)'])
          dictionary_db['Data'][c]["Cheque DetailsBank-in Date"] = dictionary_db['Data'][c]["Cheque DetailsBank-in Date"].replace(dictionary_db['Data'][c]["Cheque DetailsBank-in Date"], mydataframe.iloc[b]['Cheque Details - Bank-in Date'])
          dictionary_db['Data'][c]["Cheque DetailsChq.No."] = dictionary_db['Data'][c]["Cheque DetailsChq.No."].replace(dictionary_db['Data'][c]["Cheque DetailsChq.No."], mydataframe.iloc[b]['Cheque Details - Chq.No.'])
          dictionary_db['Data'][c]["Cheque DetailsAmount (RM)"] = dictionary_db['Data'][c]["Cheque DetailsAmount (RM)"].replace(dictionary_db['Data'][c]["Cheque DetailsAmount (RM)"], mydataframe.iloc[b]['Cheque Details - Amount(RM)'])
          dictionary_db['Data'][c]["Hospital Bills DetailsDate"] = dictionary_db['Data'][c]["Hospital Bills DetailsDate"].replace(dictionary_db['Data'][c]["Hospital Bills DetailsDate"], str(mydataframe.iloc[b]['Hospital Bills Details - Date']))
          dictionary_db['Data'][c]["Hospital Bills DetailsBill No."] = dictionary_db['Data'][c]["Hospital Bills DetailsBill No."].replace(str(dictionary_db['Data'][c]["Hospital Bills DetailsBill No."]), mydataframe.iloc[b]['Hospital Bills Details - Bill No.'])
          dictionary_db['Data'][c]["Hospital Bills DetailsAmount (RM)"] = str(dictionary_db['Data'][c]["Hospital Bills DetailsAmount (RM)"]).replace(str(dictionary_db['Data'][c]["Hospital Bills DetailsAmount (RM)"]), str(mydataframe.iloc[b]['Hospital Bills Details - Amount(RM)']))

          ##
          dictionary_db['Data'][c]["Under / (Over) Reimbursed by Client(AB-Y-AN+AQ)"] = dictionary_db['Data'][c]["Under / (Over) Reimbursed by Client(AB-Y-AN+AQ)"].replace(dictionary_db['Data'][c]["Under / (Over) Reimbursed by Client(AB-Y-AN+AQ)"], mydataframe.iloc[b]['Under / (Over) Reimbursed by Client (AB-Y-AN+AQ)'])
          dictionary_db['Data'][c]["UP/OP CASES"] = dictionary_db['Data'][c]["UP/OP CASES"].replace(dictionary_db['Data'][c]["UP/OP CASES"], mydataframe.iloc[b]['UP/OP CASES'])
          dictionary_db['Data'][c]["AAN Settlement DetailsDate"] = dictionary_db['Data'][c]["AAN Settlement DetailsDate"].replace(dictionary_db['Data'][c]["AAN Settlement DetailsDate"], mydataframe.iloc[b]['AAN Settlement Details - Date'])
          dictionary_db['Data'][c]["AAN Settlement DetailsChq.No."] = dictionary_db['Data'][c]["AAN Settlement DetailsChq.No."].replace(dictionary_db['Data'][c]["AAN Settlement DetailsChq.No."], mydataframe.iloc[b]['AAN Settlement Details - Chq.No.'])
          dictionary_db['Data'][c]["AAN Settlement DetailsAmount (RM)"] = dictionary_db['Data'][c]["AAN Settlement DetailsAmount (RM)"].replace(dictionary_db['Data'][c]["AAN Settlement DetailsAmount (RM)"], mydataframe.iloc[b]['AAN Settlement Details - Amount(RM)'])
          dictionary_db['Data'][c]["Under / (Over) Paid to Hosp (AB - AG)"] = dictionary_db['Data'][c]["Under / (Over) Paid to Hosp (AB - AG)"].replace(dictionary_db['Data'][c]["Under / (Over) Paid to Hosp (AB - AG)"], mydataframe.iloc[b]['Under / (Over) Paid to Hosp (AB - AG)'])
          dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountOCBC 16"] = dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountOCBC 16"].replace(dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountOCBC 16"], "")
          dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountBIMB 1"] = dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountBIMB 1"].replace(dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountBIMB 1"], "")
          dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountDBMY29"] = str(dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountDBMY29"]).replace(str(dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountDBMY29"]), "")
          dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountDBMY29.1"] = dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountDBMY29.1"].replace(dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountDBMY29.1"], "")
          dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountDBMY29.2"] = dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountDBMY29.2"].replace(dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountDBMY29.2"], "")
          dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountDBMY29.3"] = dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountDBMY29.3"].replace(dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountDBMY29.3"], "")
          dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountDBMY29.4"] = dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountDBMY29.4"].replace(dictionary_db['Data'][c]["Designated  / Borrowing Bank AccountDBMY29.4"], "")
          dictionary_db['Data'][c]["TPA Billing DetailsDate"] = dictionary_db['Data'][c]["TPA Billing DetailsDate"].replace(dictionary_db['Data'][c]["TPA Billing DetailsDate"], mydataframe.iloc[b]['TPA Billing Details - Date'])
          dictionary_db['Data'][c]["TPA Billing DetailsInv. No."] = dictionary_db['Data'][c]["TPA Billing DetailsInv. No."].replace(dictionary_db['Data'][c]["TPA Billing DetailsInv. No."], mydataframe.iloc[b]['TPA Billing Details - Inv. No.'])
          dictionary_db['Data'][c]["TPA Billing DetailsAmount (RM)"] = dictionary_db['Data'][c]["TPA Billing DetailsAmount (RM)"].replace(dictionary_db['Data'][c]["TPA Billing DetailsAmount (RM)"], mydataframe.iloc[b]['TPA Billing Details - Amount(RM)'])
          dictionary_db['Data'][c]["Reimbursement DetailsDate"] = dictionary_db['Data'][c]["Reimbursement DetailsDate"].replace(dictionary_db['Data'][c]["Reimbursement DetailsDate"], mydataframe.iloc[b]['Reimbursement Details - Date'])
          dictionary_db['Data'][c]["Reimbursement DetailsChq. No."] = dictionary_db['Data'][c]["Reimbursement DetailsChq. No."].replace(dictionary_db['Data'][c]["Reimbursement DetailsChq. No."], mydataframe.iloc[b]['Reimbursement Details - Chq.No.'])
          dictionary_db['Data'][c]["Reimbursement DetailsAmount (RM)"] = dictionary_db['Data'][c]["Reimbursement DetailsAmount (RM)"].replace(dictionary_db['Data'][c]["Reimbursement DetailsAmount (RM)"], mydataframe.iloc[b]['Reimbursement Details - Amount(RM)'])
          dictionary_db['Data'][c]["Adjustment DetailsDate"] = dictionary_db['Data'][c]["Adjustment DetailsDate"].replace(dictionary_db['Data'][c]["Adjustment DetailsDate"], mydataframe.iloc[b]['Adjustment Details - Date'])
          dictionary_db['Data'][c]["Adjustment DetailsDoc. No."] = dictionary_db['Data'][c]["Adjustment DetailsDoc. No."].replace(dictionary_db['Data'][c]["Adjustment DetailsDoc. No."], mydataframe.iloc[b]['Adjustment Details - Doc.No'])
          dictionary_db['Data'][c]["Adjustment DetailsAmount (RM)"] = dictionary_db['Data'][c]["Adjustment DetailsAmount (RM)"].replace(dictionary_db['Data'][c]["Adjustment DetailsAmount (RM)"], mydataframe.iloc[b]['Adjustment Details - Amount(RM)'])
          dictionary_db['Data'][c]["Hospital OR DetailsOR Date"] = dictionary_db['Data'][c]["Hospital OR DetailsOR Date"].replace(dictionary_db['Data'][c]["Hospital OR DetailsOR Date"], mydataframe.iloc[b]['Hospital OR Details - OR Date'])
          dictionary_db['Data'][c]["Hospital OR DetailsOR No."] = dictionary_db['Data'][c]["Hospital OR DetailsOR No."].replace(dictionary_db['Data'][c]["Hospital OR DetailsOR No."], mydataframe.iloc[b]['Hospital OR Details - OR No.'])
          dictionary_db['Data'][c]["Hospital OR DetailsAmount (RM)"] = dictionary_db['Data'][c]["Hospital OR DetailsAmount (RM)"].replace(dictionary_db['Data'][c]["Hospital OR DetailsAmount (RM)"], mydataframe.iloc[b]['Hospital OR Details - Amount(RM)'])
          dictionary_db['Data'][c]["Hospital OR DetailsDate of OR Rec'd"] = dictionary_db['Data'][c]["Hospital OR DetailsDate of OR Rec'd"].replace(dictionary_db['Data'][c]["Hospital OR DetailsDate of OR Rec'd"], "")
          dictionary_db['Data'][c]["Hospital OR DetailsSubm. Date"] = dictionary_db['Data'][c]["Hospital OR DetailsSubm. Date"].replace(dictionary_db['Data'][c]["Hospital OR DetailsSubm. Date"], "")
          dictionary_db['Data'][c]["Tax Invoice DetailsDate"] = dictionary_db['Data'][c]["Tax Invoice DetailsDate"].replace(dictionary_db['Data'][c]["Tax Invoice DetailsDate"], mydataframe.iloc[b]['Tax Invoice Details - Date'])
          dictionary_db['Data'][c]["Tax Invoice DetailsInv No"] = dictionary_db['Data'][c]["Tax Invoice DetailsInv No"].replace(dictionary_db['Data'][c]["Tax Invoice DetailsInv No"], mydataframe.iloc[b]['Tax Invoice Details - Inv No'])
          dictionary_db['Data'][c]["Tax Invoice DetailsCase Fee"] = dictionary_db['Data'][c]["Tax Invoice DetailsCase Fee"].replace(dictionary_db['Data'][c]["Tax Invoice DetailsCase Fee"], mydataframe.iloc[b]['Tax Invoice Details - Case Fee'])
          dictionary_db['Data'][c]["Tax Invoice DetailsGST Amount"] = dictionary_db['Data'][c]["Tax Invoice DetailsGST Amount"].replace(dictionary_db['Data'][c]["Tax Invoice DetailsGST Amount"], mydataframe.iloc[b]['Tax Invoice Details - GST Amount'])
          dictionary_db['Data'][c]["Tax Invoice DetailsTotal"] = dictionary_db['Data'][c]["Tax Invoice DetailsTotal"].replace(dictionary_db['Data'][c]["Tax Invoice DetailsTotal"], mydataframe.iloc[b]['Tax Invoice Details - Total'])

          # Dump the edited JSON string back to the database
          result = json.dumps(dictionary_db, default = str)
          cur.execute("UPDATE cba.client_master SET content = %s WHERE year = %s and client_id = %s and invoice_type = %s",(result, current_year, client_id, sheet))
          conn.commit()
        else:
          # Append as a new record
          dictionary_db["Data"].append(my_dict.copy())
          result = json.dumps(dictionary_db, default = str)
          cur.execute("UPDATE cba.client_master SET content = %s WHERE year = %s and client_id = %s and invoice_type = %s",(result, current_year, client_id, sheet))
          conn.commit()
      else:
        # Append as a new record
        result = json.dumps(my_dict.copy(), default = str)
        cur.execute('''INSERT INTO cba.client_master (year, client_id, invoice_type, content, createdby, modifiedby, createdon, modifiedon, url)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)''', (current_year, client_id, sheet, result, cmlrpa_base.email, cmlrpa_base.email, current_date, current_date, filelocation))
        conn.commit()
  except Exception as error:
    logging("REQDCA - Unable to update CM table in database.", error, reqdca_base)

