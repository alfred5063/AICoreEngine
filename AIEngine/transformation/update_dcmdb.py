#!/usr/bin/python
# FINAL SCRIPT updated as of 4th June 2020
# Workflow - Update DCM
# Version 1

# Declare Python libraries needed for this script
import time
from datetime import datetime as dt
import pandas as pd
import json as json
from connector.connector import MSSqlConnector
from connector.dbconfig import *
import pyodbc as db

# Read DCM
mydataframe = pd.read_excel(r'\\dtisvr2\CBA_UAT\Testing\DCM Update\DCM 2020.xlsx')

def update_dcmdb(mydataframe):

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

  # Get the current DCM records
  print("Querying latest JSON string from Disbursement Claim Master file.")
  param = (current_year)
  sql_query_insurer = "SELECT content FROM cba.disbursement_master WHERE year = %s"
  cur.execute(sql_query_insurer, param)
  myquery = json.loads(cur.fetchall()[0][0])
  mylist_a = []
  mylist_b = []

  try:

    # Update the Values in existing records
    print("Updating the existing records in Disbursement Claim Master file.")
    a = 0
    for a in range(len(mydataframe['DisbursementClaimsNo'])):
      b = 0
      for b in range(len(myquery['Data'])):
        if str(myquery['Data'][b]['Disbursement Claims No']) == str(mydataframe.iloc[a]['DisbursementClaimsNo']):

          try:
            myquery['Data'][b]["Date"] = str(myquery['Data'][b]["Date"]).replace(str(myquery['Data'][b]["Date"]), str(mydataframe.iloc[a]['MyDate']))
            myquery['Data'][b]["Types of Invoice"] = str(myquery['Data'][b]["Types of Invoice"]).replace(str(myquery['Data'][b]["Types of Invoice"]), str(mydataframe.iloc[a]['TypesofInvoice']))
            myquery['Data'][b]["Disbursement Claims No"] = str(myquery['Data'][b]["Disbursement Claims No"]).replace(str(myquery['Data'][b]["Disbursement Claims No"]), str(mydataframe.iloc[a]['DisbursementClaimsNo']))
            myquery['Data'][b]["Claims Listing No"] = str(myquery['Data'][b]["Claims Listing No"]).replace(str(myquery['Data'][b]["Claims Listing No"]), str(mydataframe.iloc[a]['ClaimsListingNo']))
            myquery['Data'][b]["No of Cases"] = str(myquery['Data'][b]["No of Cases"]).replace(str(myquery['Data'][b]["No of Cases"]), str(mydataframe.iloc[a]['NoofCases']))
            myquery['Data'][b]["File No"] = str(myquery['Data'][b]["File No"]).replace(str(myquery['Data'][b]["File No"]), str(mydataframe.iloc[a]['FileNo']))
            myquery['Data'][b]["Customer Name Master"] = str(myquery['Data'][b]["Customer Name Master"]).replace(str(myquery['Data'][b]["Customer Name Master"]), str(mydataframe.iloc[a]['CustomerNameMaster']))
            myquery['Data'][b]["Hospital"] = str(myquery['Data'][b]["Hospital"]).replace(str(myquery['Data'][b]["Hospital"]), str(mydataframe.iloc[a]['Hospital']))
            myquery['Data'][b]["Bill No"] = str(myquery['Data'][b]["Bill No"]).replace(str(myquery['Data'][b]["Bill No"]), str(mydataframe.iloc[a]['BillNo']))
            myquery['Data'][b]["Patient"] = str(myquery['Data'][b]["Patient"]).replace(str(myquery['Data'][b]["Patient"]), str(mydataframe.iloc[a]['Patient']))
            myquery['Data'][b]["Bill Amount"] = str(myquery['Data'][b]["Bill Amount"]).replace(str(myquery['Data'][b]["Bill Amount"]), str(mydataframe.iloc[a]['BillAmount']))
            myquery['Data'][b]["Reasons"] = str(myquery['Data'][b]["Reasons"]).replace(str(myquery['Data'][b]["Reasons"]), str(mydataframe.iloc[a]['Reasons']))
            myquery['Data'][b]["Action"] = str(myquery['Data'][b]["Action"]).replace(str(myquery['Data'][b]["Action"]), str(mydataframe.iloc[a]['MyAction']))
            myquery['Data'][b]["Initial"] = str(myquery['Data'][b]["Initial"]).replace(str(myquery['Data'][b]["Initial"]), str(mydataframe.iloc[a]['Initial']))
            myquery['Data'][b]["Bank in Date"] = str(myquery['Data'][b]["Bank in Date"]).replace(str(myquery['Data'][b]["Bank in Date"]), str(mydataframe.iloc[a]['BankinDate']))
            myquery['Data'][b]["Cheque No or TT"] = str(myquery['Data'][b]["Cheque No or TT"]).replace(str(myquery['Data'][b]["Cheque No or TT"]), str(mydataframe.iloc[a]['ChequeNoorTT']))
            myquery['Data'][b]["Remarks"] = str(myquery['Data'][b]["Remarks"]).replace(str(myquery['Data'][b]["Remarks"]), str(mydataframe.iloc[a]['Remarks']))
            myquery['Data'][b]["Reason for Adjustment"] = str(myquery['Data'][b]["Reason for Adjustment"]).replace(str(myquery['Data'][b]["Reason for Adjustment"]), str(mydataframe.iloc[a]['ReasonforAdjustment']))
            myquery['Data'][b]["Next Action Taken"] = str(myquery['Data'][b]["Next Action Taken"]).replace(str(myquery['Data'][b]["Next Action Taken"]), str(mydataframe.iloc[a]['NextActionTaken']))
            myquery['Data'][b]["New invoice no"] = str(myquery['Data'][b]["New invoice no"]).replace(str(myquery['Data'][b]["New invoice no"]), str(mydataframe.iloc[a]['Newinvoiceno']))
            myquery['Data'][b]["Invoice Amount "] = str(myquery['Data'][b]["Invoice Amount "]).replace(str(myquery['Data'][b]["Invoice Amount "]), str(mydataframe.iloc[a]['InvoiceAmount']))
            myquery['Data'][b]["Issue Date"] = str(myquery['Data'][b]["Issue Date"]).replace(str(myquery['Data'][b]["Issue Date"]), str(mydataframe.iloc[a]['IssueDate']))
            myquery['Data'][b]["Cheque No"] = str(myquery['Data'][b]["Cheque No"]).replace(str(myquery['Data'][b]["Cheque No"]), str(mydataframe.iloc[a]['ChequeNo']))
            myquery['Data'][b]["Settlement Amount"] = str(myquery['Data'][b]["Settlement Amount"]).replace(str(myquery['Data'][b]["Settlement Amount"]), str(mydataframe.iloc[a]['SettlementAmount']))
            myquery['Data'][b]["OCBC 4"] = str(myquery['Data'][b]["OCBC 4"]).replace(str(myquery['Data'][b]["OCBC 4"]), str(mydataframe.iloc[a]['OCBC4']))
            myquery['Data'][b]["DB 04"] = str(myquery['Data'][b]["DB 04"]).replace(str(myquery['Data'][b]["DB 04"]), str(mydataframe.iloc[a]['DB04']))
            myquery['Data'][b]["DB 10"] = str(myquery['Data'][b]["DB 10"]).replace(str(myquery['Data'][b]["DB 10"]), str(mydataframe.iloc[a]['DB10']))
            myquery['Data'][b]["DB 20"] = str(myquery['Data'][b]["DB 20"]).replace(str(myquery['Data'][b]["DB 20"]), str(mydataframe.iloc[a]['DB20']))
            myquery['Data'][b]["DB 21"] = str(myquery['Data'][b]["DB 21"]).replace(str(myquery['Data'][b]["DB 21"]), str(mydataframe.iloc[a]['DB21']))
            myquery['Data'][b]["DB 39"] = str(myquery['Data'][b]["DB 39"]).replace(str(myquery['Data'][b]["DB 39"]), str(mydataframe.iloc[a]['DB39']))
            myquery['Data'][b]["BIMB 1"] = str(myquery['Data'][b]["BIMB 1"]).replace(str(myquery['Data'][b]["BIMB 1"]), str(mydataframe.iloc[a]['BIMB1']))
            myquery['Data'][b]["DB 03"] = str(myquery['Data'][b]["DB 03"]).replace(str(myquery['Data'][b]["DB 03"]), str(mydataframe.iloc[a]['DB03']))
            myquery['Data'][b]["DB 27"] = str(myquery['Data'][b]["DB 27"]).replace(str(myquery['Data'][b]["DB 27"]), str(mydataframe.iloc[a]['DB27']))
            myquery['Data'][b]["HSBC 2"] = str(myquery['Data'][b]["HSBC 2"]).replace(str(myquery['Data'][b]["HSBC 2"]), str(mydataframe.iloc[a]['HSBC2']))
            myquery['Data'][b]["AP Team Remarks"] = str(myquery['Data'][b]["AP Team Remarks"]).replace(str(myquery['Data'][b]["AP Team Remarks"]), str(mydataframe.iloc[a]['APTeamRemarks']))

            # Store found HNS into a list
            mylist_a.insert(b, str(mydataframe.iloc[a]['DisbursementClaimsNo']))

            # Dump the edited JSON string back to the database
            result = json.dumps(myquery, default = str)
            cur.execute("update [cba].[disbursement_master] set content = %s where [year] = %s",(result, current_year))
            conn.commit()
          except:
            pass

        else:
          pass
        continue


    # Append new records
    print("Inserting new records into Disbursement Claim Master file.")
    c = 0
    for c in range(len(mydataframe['DisbursementClaimsNo'])):
      mylist_b.insert(c, str(mydataframe['DisbursementClaimsNo'][c]))

    new_record = list(set(mylist_b) - set(mylist_a))

    d = 0
    for d in range(len(new_record)):

      searched_df = mydataframe[mydataframe['DisbursementClaimsNo'].astype(str).str.contains("%s" % new_record[d])]
      searched_df = searched_df.reset_index().drop(columns = ['index'])

      if searched_df.empty != True:

        my_dict = []
        my_dict = {
          "Date":"", "Types of Invoice":"", "Disbursement Claims No":"", "Claims Listing No":"", "No of Cases":"", "File No":"", "Customer Name Master":"",
          "Hospital":"", "Bill No":"", "Patient":"", "Bill Amount":"", "Reasons":"", "Action":"", "Initial":"", "Bank in Date":"", "Cheque No or TT":"",
          "Cheque Amount Paid":"", "Remarks":"", "Reason for Adjustment":"", "Next Action Taken":"", "New invoice no":"", "Invoice Amount ":"", "Issue Date":"",
          "Cheque No":"", "Settlement Amount":"", "OCBC 4":"", "DB 04":"", "DB 10":"", "DB 20":"", "DB 21":"", "DB 39":"", "BIMB 1":"", "DB 03":"", "DB 27":"",
          "HSBC 2":"", "AP Team Remarks":""
        }

        my_dict["Date"] = searched_df.iloc[0]['MyDate']
        my_dict["Types of Invoice"] = searched_df.iloc[0]['TypesofInvoice']
        my_dict["Disbursement Claims No"] = searched_df.iloc[0]['DisbursementClaimsNo']
        my_dict["Claims Listing No"] = searched_df.iloc[0]['ClaimsListingNo']
        my_dict["No of Cases"] = searched_df.iloc[0]['NoofCases']
        my_dict["File No"] = searched_df.iloc[0]['FileNo']
        my_dict["Customer Name Master"] = searched_df.iloc[0]['CustomerNameMaster']
        my_dict["Hospital"] = searched_df.iloc[0]['Hospital']
        my_dict["Bill No"] = searched_df.iloc[0]['BillNo']
        my_dict["Patient"] = searched_df.iloc[0]['Patient']
        my_dict["Bill Amount"] = searched_df.iloc[0]['BillAmount']
        my_dict["Reasons"] = searched_df.iloc[0]['Reasons']
        my_dict["Action"] = searched_df.iloc[0]['MyAction']
        my_dict["Initial"] = searched_df.iloc[0]['Initial']
        my_dict["Bank in Date"] = searched_df.iloc[0]['BankinDate']
        my_dict["Cheque No or TT"] = searched_df.iloc[0]['ChequeNoorTT']
        my_dict["Cheque Amount Paid"] = ''
        my_dict["Remarks"] = searched_df.iloc[0]['Remarks']
        my_dict["Reason for Adjustment"] = searched_df.iloc[0]['ReasonforAdjustment']
        my_dict["Next Action Taken"] = searched_df.iloc[0]['NextActionTaken']
        my_dict["New invoice no"] = searched_df.iloc[0]['Newinvoiceno']
        my_dict["Invoice Amount"] = searched_df.iloc[0]['InvoiceAmount']
        my_dict["Issue Date"] = searched_df.iloc[0]['IssueDate']
        my_dict["Cheque No"] = searched_df.iloc[0]['ChequeNo']
        my_dict["Settlement Amount"] = searched_df.iloc[0]['SettlementAmount']
        my_dict["OCBC 4"] = searched_df.iloc[0]['OCBC4']
        my_dict["DB 04"] = searched_df.iloc[0]['DB04']
        my_dict["DB 10"] = searched_df.iloc[0]['DB10']
        my_dict["DB 20"] = searched_df.iloc[0]['DB20']
        my_dict["DB 21"] = searched_df.iloc[0]['DB21']
        my_dict["DB 39"] = searched_df.iloc[0]['DB39']
        my_dict["BIMB 1"] = searched_df.iloc[0]['BIMB1']
        my_dict["DB 03"] = searched_df.iloc[0]['DB03']
        my_dict["DB 27"] = searched_df.iloc[0]['DB27']
        my_dict["HSBC 2"] = searched_df.iloc[0]['HSBC2']
        my_dict["AP Team Remarks"] = searched_df.iloc[0]['APTeamRemarks']

        # Append as a new record
        myquery["Data"].append(my_dict.copy())
        result = json.dumps(myquery, default = str)
        cur.execute("UPDATE [cba].[disbursement_master] set content = %s where [year] = %s",(result, current_year))
        conn.commit()

      else:
        pass

    print("Daily update to Disbursement Claim Master file is done.")
  except Exception:
    print("Nothing")

