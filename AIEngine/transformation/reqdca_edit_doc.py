#!/usr/bin/python
# FINAL SCRIPT updated as of 8th July 2020
# Workflow - REQ/DCA

# Declare Python libraries needed for this script
import pandas as pd
import numpy as np
import time
import psycopg2
import json
import copy
from connector.connector import MySqlConnector, MSSqlConnector
from connector.dbconfig import *
from automagic.automagica_reqdca import *
from datetime import datetime as dt
from utils.audit_trail import audit_log
from utils.logging import logging
import pyodbc as db
from loading.query.query_as_df import *

# Function to update document
def update_doc(newmain_df, doc_template, email, reqdca_base, document):

  try:

    # Get the current date and time
    named_tuple = time.localtime()
    current_date = time.strftime("%d-%m-%Y", named_tuple)
    current_time = time.strftime("%H%M", named_tuple)
    current_year = dt.now().year
    curyear = time.strftime("%y", time.localtime())

    # Database connection
    conn = MSSqlConnector()
    cur = conn.cursor()
    conn_mysql = MySqlConnector()
    cur_mysql = conn_mysql.cursor()

    # Query in MARC database
    # Take note that querying development databse might not give the correct result.
    paramater = [str(email),]
    stored_proc = cur_mysql.callproc('dca_query_marc_for_user_email', paramater)
    for i in cur_mysql.stored_results():
      results = i.fetchall()
    if results != []:
      shortname = results[0][0]
    else:
      shortname = ''

    # Iterate the case ID and create and save the DCA document for each ID
    i = 0
    for i in range(len(newmain_df)):

      try:

        # Prepare variables
        caseid = int(newmain_df.iloc[i]['Case ID'])
        attention = newmain_df.iloc[i]['Attention To']
        insurer = 'Insurer'
        address = newmain_df.iloc[i]['Address']

        if document == 'DCA':
          cln = newmain_df.iloc[i]['Disbursement Listing']
          # DCA - Query database to get the JSON string containing HNS number
          conn1 = MSSqlConnector
          query = '''SELECT content FROM cba.disbursement_master WHERE year = %s'''
          params = (current_year)
          fetched_hns = mssql_get_df_by_query(query, params, reqdca_base, conn1)
          fetched_hns = fetched_hns.to_json(orient = 'columns')
          fetched_hns = fetched_hns.replace('T00:00:00.000Z', '')
          fetched_json = json.loads(fetched_hns) # JSON object
          fetched_json = fetched_json['content']['0']

          # Get the last record to fetch current HNS number
          fetched_json_df = pd.DataFrame(json.loads(fetched_json))
          fetched_json_df_last = fetched_json_df.values[-1].tolist()
          curr_hns = fetched_json_df_last[0]['Disbursement Claims No'].split("-")
          hsn_curyear = curr_hns[0].split("H&S")
          if hsn_curyear[1] != curyear:
            curr_hns[0] = "H&S" + curyear
            curr_hns[1] = "0".zfill(5)
          else:
            curr_hns[0] = curr_hns[0]
            curr_hns[1] = curr_hns[1]
          n = 1
          while (curr_hns[0] == ''):
            fetched_json_df_last = fetched_json_df.values[-n].tolist()
            curr_hns = fetched_json_df_last[0]['Disbursement Claims No'].split("-")
            if curr_hns[0] == '':
              n = n + 1
            else:
              a = 0
              b = 1
              HNS = curr_hns[a] + "-" + str(int(curr_hns[b]) + 1).zfill(5)
          else:
            a = 0
            b = 1
            HNS = curr_hns[a] + "-" + str(int(curr_hns[b]) + 1).zfill(5)
            newmain_df.loc[i, 'HNS Number'] = HNS

          subcaseid = int(newmain_df.iloc[i]['Sub Case ID_x'])
          mytype = newmain_df.iloc[i]['Type_x']
          amount = newmain_df.iloc[i]['Total']

          print("- Updating record in SQL database")
          my_dict = []
          my_dict = {
            "Date":"", "Types of Invoice":"", "Disbursement Claims No":"", "Claims Listing No":"", "No of Cases":"", "File No":"", "Customer Name Master":"",
            "Hospital":"", "Bill No":"", "Patient":"", "Bill Amount":"", "Reasons":"", "Action":"", "Initial":"", "Bank in Date":"", "Cheque No or TT":"",
            "Cheque Amount Paid":"", "Remarks":"", "Reason for Adjustment":"", "Next Action Taken":"", "New invoice no":"", "Invoice Amount ":"", "Issue Date":"",
            "Cheque No":"", "Settlement Amount":"", "OCBC 4":"", "DB 04":"", "DB 10":"", "DB 20":"", "DB 21":"", "DB 39":"", "BIMB 1":"", "DB 03":"", "DB 27":"",
            "HSBC 2":"", "AP Team Remarks":""}

          my_dict["Date"] = current_date
          my_dict["Types of Invoice"] = str(doc_template)
          my_dict["Disbursement Claims No"] = "%s" % str(HNS)
          my_dict["Claims Listing No"] = str(cln)
          my_dict["No of Cases"] = "1"

          if mytype == 'Admission':
            my_dict["File No"] = str(caseid)
          else:
            my_dict["File No"] = str(caseid) + "-" + str(subcaseid)

          my_dict["Customer Name Master"] = newmain_df.iloc[i]['Insurer Name']
          my_dict["Hospital"] = newmain_df.iloc[i]['Hospital Name']
          my_dict["Bill No"] = newmain_df.iloc[i]['Bill No']
          my_dict["Patient"] = newmain_df.iloc[i]['''Member's Name''']
          my_dict["Bill Amount"] = newmain_df.iloc[i]['Total']
          my_dict["Reasons"] = mytype # POST or Admission
          my_dict["Action"] = ""
          my_dict["Initial"] = shortname
          my_dict["Bank in Date"] = ""
          my_dict["Cheque No or TT"] = ""
          my_dict["Cheque Amount Paid"] = ""
          my_dict["Remarks"] = newmain_df.iloc[i]['Reason'] # Will appear in MARC document
          my_dict["Reason for Adjustment"] = newmain_df.iloc[i]['Remarks_x'] # Dropdown list from front-end
          my_dict["Next Action Taken"] = ""
          my_dict["New invoice no"] = ""
          my_dict["Invoice Amount"] = ""
          my_dict["Issue Date"] = ""
          my_dict["Cheque No"] = ""
          my_dict["Settlement Amount"] = ""
          my_dict["OCBC 4"] = ""
          my_dict["DB 04"] = ""
          my_dict["DB 10"] = ""
          my_dict["DB 20"] = ""
          my_dict["DB 21"] = ""
          my_dict["DB 39"] = ""
          my_dict["BIMB 1"] = ""
          my_dict["DB 03"] = ""
          my_dict["DB 27"] = ""
          my_dict["HSBC 2"] = ""
          my_dict["AP Team Remarks"] = ""

          # Updating the table in MSSQL database
          cur.execute("select content from [cba].[disbursement_master] where [year] = %s",(current_year))
          dictionary_db = json.loads(cur.fetchall()[0][0])
          list_db = dictionary_db["Data"]
          list_db.append(my_dict.copy())
          result = json.dumps(dictionary_db, default = str)
          cur.execute("update [cba].[disbursement_master] set content = %s where [year] = %s",(result, current_year))
          conn.commit()
          audit_log("REQ/DCA - Successfull detail for Case ID [ %s ] saved in SQL database." % caseid, "Completed...", reqdca_base)

        else:
          # REQ - Query database to get the JSON string containing HNS number
          conn1 = MSSqlConnector
          query = '''SELECT content FROM cba.ops_disbursement_req_master WHERE year = %s'''
          params = (current_year)
          fetched_hns = mssql_get_df_by_query(query, params, reqdca_base, conn1)
          fetched_hns = fetched_hns.to_json(orient = 'columns')
          fetched_hns = fetched_hns.replace('T00:00:00.000Z', '')
          fetched_json = json.loads(fetched_hns) # JSON object
          fetched_json = fetched_json['content']['0']

          # Get the last record to fetch current HNS number
          fetched_json_df = pd.DataFrame(json.loads(fetched_json))
          fetched_json_df_last = fetched_json_df.values[-1].tolist()
          curr_hns = fetched_json_df_last[0]['Disbursement claims no'].split("-")
          hsn_curyear = curr_hns[0].split("REQ")
          if hsn_curyear[1] != curyear:
            curr_hns[0] = "REQ" + curyear
            curr_hns[1] = "0".zfill(4)
          else:
            curr_hns[0] = curr_hns[0]
            curr_hns[1] = curr_hns[1]
          n = 1
          while (curr_hns[0] == ''):
            fetched_json_df_last = fetched_json_df.values[-n].tolist()
            curr_hns = fetched_json_df_last[0]['Disbursement claims no'].split("-")
            if curr_hns[0] == '':
              n = n + 1
            else:
              a = 0
              b = 1
              HNS = curr_hns[a] + "-" + str(int(curr_hns[b]) + 1).zfill(4)
          #else:
          #  a = 0
          #  b = 1
          #  HNS = curr_hns[a] + "-" + str(int(curr_hns[b]) + 1).zfill(4)
          #  newmain_df.loc[i, 'HNS Number'] = HNS
          else:
            a = 0
            b = 1
            HNS = newmain_df.iloc[i]['HNS Number']
            newmain_df.loc[i, 'HNS Number'] = HNS

          subcaseid = int(newmain_df.iloc[i]['Sub Case ID'])
          mytype = newmain_df.iloc[i]['Type']
          amount = newmain_df.iloc[i]['Amount']

          my_dict = []
          my_dict = {
            "Date": "", "Disbursement claims no": "", "Claims listing no": "", "SAGE ID": "", "File": "", "Bill to": "",
            "Client": "", "Patient Name": "", "Status (VIP / Non-Vip Tan chong claims) ": "", "Policy no ": "",
            "Admission date": "", "Discharge date": "", "Hospital": "", "Bill no": "", "Amounts": "", "OB Received date": "",
            "OB Registered date": "", "Reasons": "", "Cashless / Post / Fruit Basket / Reimbursement ": "", "Initial": "",
            "Bank in date": "", "Details": "", "Amounts Received": "", "AP Team Remarks": "", "Date payment issued": "",
            "Chq No. / Giro": "", "Amount (RM)": "", "BIMY 01": "", "DBMY37 (Etiqa)": "", "HLBB 2 (Axa:Chq)": "", "DBMY14 (Axa:Giro)": "",
            "DBMY15 (RHB:Giro)": "", "DBMY16 (Mpower/SS2)": "", "DBMY42 (STMB)": "", "DBMY39 (Finance) Borrowing Bank Account": "", "DBMY38 (Operation)": "",
            "Repayment Date": ""
            }

          my_dict["Date"] = str(current_date)
          my_dict["Disbursement claims no"] = "%s" % str(HNS)

          if mytype == 'Admission':
            cln = str(newmain_df.iloc[i]['Client Listing Number'])
            my_dict["File"] = str(caseid)
            my_dict["Claims listing no"] = str(cln)
          else:
            cln = str(newmain_df.iloc[i]['Client Listing Number'])
            my_dict["Claims listing no"] = str(cln)
            my_dict["File"] = str(caseid) + "-" + str(subcaseid)

          my_dict["SAGE ID"] = None
          my_dict["Bill to"] =  newmain_df.iloc[i]['Insurer Name']
          my_dict["Client"] =  newmain_df.iloc[i]['Insurer Name']
          my_dict["Patient Name"] = newmain_df.iloc[i]['Patient Name']
          my_dict["Status (VIP / Non-Vip Tan chong claims) "] = ""
          my_dict["Policy no "] = newmain_df.iloc[i]['Insurence Policy Number']
          my_dict["Admission date"] = newmain_df.iloc[i]['Admission Date']
          my_dict["Discharge date"] = newmain_df.iloc[i]['Discharged Date']
          my_dict["Hospital"] = newmain_df.iloc[i]['Hospital Name']
          my_dict["Bill no"] = newmain_df.iloc[i]['Bill. Num.']
          my_dict["Amounts"] = newmain_df.iloc[i]['Amount']
          my_dict["OB Received date"] = newmain_df.iloc[i]['OB Received Date']
          my_dict["OB Registered date"] = newmain_df.iloc[i]['OB Registered Date']
          my_dict["Reasons"] = newmain_df.iloc[i]['Reason']
          my_dict["Cashless / Post / Fruit Basket / Reimbursement "] = newmain_df.iloc[i]['Remarks']
          my_dict["Initial"] = shortname
          my_dict["Bank in date"] = None
          my_dict["Details"] = None
          my_dict["Amounts Received"] = None
          my_dict["AP Team Remarks"] = None
          my_dict["Date payment issued"] = None
          my_dict["Chq No. / Giro"] = None
          my_dict["Amount (RM)"] = None
          my_dict["BIMY 01"] = None
          my_dict["DBMY37 (Etiqa)"] = None
          my_dict["HLBB 2 (Axa] = Chq)"] = None
          my_dict["DBMY14 (Axa] = Giro)"] = None
          my_dict["DBMY15 (RHB] = Giro)"] = None
          my_dict["DBMY16 (Mpower/SS2)"] = None
          my_dict["DBMY42 (STMB)"] = None
          my_dict["DBMY39 (Finance) Borrowing Bank Account"] = None
          my_dict["DBMY38 (Operation)"] = None
          my_dict["Repayment Date"] = None
        
          # Updating the table in MSSQL database
          try:
            cur.execute("select content from cba.ops_disbursement_req_master where year = %s",(current_year))
            dictionary_db = json.loads(cur.fetchall()[0][0])
            dictionary_db["Data"].append(my_dict.copy())
            result = json.dumps(dictionary_db, default = str)
            cur.execute("update [cba].[ops_disbursement_req_master] set content = %s where [year] = %s",(result, current_year))
            conn.commit()
            audit_log("REQ/DCA - Successful detail for Case ID [ %s ] saved in SQL database." % caseid, "Completed...", reqdca_base)
          except Exception as error:
            audit_log("REQ/DCA - Detail for Case ID [ %s ] not saved in SQL database. Please contact administrator." % caseid, "Completed...", reqdca_base)
            logging("REQ/DCA - Detail for Case ID [ %s ] not saved in SQL database. Please contact administrator." % caseid, error, reqdca_base)
            pass

      except Exception as error:
        continue
      continue

    cur.close()
  except Exception as error:
    logging("REQ/DCA - extract_info", error, reqdca_base)
    pass
