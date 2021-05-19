#!/usr/bin/python
# FINAL SCRIPT updated as of 16th June 2020
# Workflow - Finance CML Update

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
from openpyxl.styles.protection import Protection
from openpyxl.worksheet.datavalidation import DataValidation
from directory.files_listing_operation import *
from directory.get_filename_from_path import *
from directory.createfolder import *

# Determine the current year and date
current_year = dt.now().year
#current_year = '2019'
current_date = time.strftime("%Y-%m-%d")

# Update dataframe with information from user input and latest date and running number
def update_dataframe(final_df, withpath_df, i, b):
  try:
    final_df.loc[b, 'SAGE ID'] = int(withpath_df['Case ID'][i])
    final_df.loc[b, 'Sub Case ID'] = int(withpath_df['Sub Case ID'][i])
    final_df.loc[b, 'Member Ref. ID'] = str(withpath_df['Member Ref. ID'][i])
    final_df.loc[b, 'Disbursement Claims'] = str(withpath_df['HNS Number'][i]) #H&S latest running number
    final_df.loc[b, 'Type'] = str(withpath_df['Type'][i])
    final_df.loc[b, 'Plan Type'] = str(withpath_df['Policy Type'][i])

    return final_df
  except Exception as error:
    print(error)

def update_cml(merged_df, sheets[a], filelocation, cmlrpa_base):
  try:

    # Database connection
    conn = MSSqlConnector()
    cur = conn.cursor()
    conn_mysql = MySqlConnector()
    cur_mysql = conn_mysql.cursor()
    final_df2 = pd.DataFrame()

    # Query in MARC database
    paramater = [str(reqdca_base.email),]
    stored_proc = cur_mysql.callproc('dca_query_marc_for_user_email', paramater)
    for i in cur_mysql.stored_results():
      results = i.fetchall()
    if results != []:
      shortname = results[0][0]
    else:
      shortname = ''

    i = 0
    for i in range(len(merged_df)):

      # Create empty main dataframe
      main_df = pd.DataFrame(columns=['Sage ID', 'Adm. No', 'File No.', 'Disbursement Listing', 'Disbursement Claims', '''Member's Name''', 'Policy No', 'Bill No',
                                      'Adm. Date', 'Dis. Date', 'Hospital', 'Diagnosis', 'Insurance', 'Variance', 'Patient', 'Total', 'Submission Date', 'OB Received Date',
                                      'OB Registered Date', 'CSU Received Date', 'Remarks', 'Payment Due Date', 'Aging (days)', 'SCAN (TAT)', 'CSU (TAT)',
                                      'PAYMENT (TAT)', 'Under / (Over) Claimed fr Client (AA-J-AM+AP)', 'Cheque Details - Bank-in Date', 'Cheque Details - Chq.No.',
                                      'Cheque Details - Amount(RM)', 'Hospital Bills Details - Date', 'Hospital Bills Details - Bill No.', 'Hospital Bills Details - Amount(RM)',
                                      'Under / (Over) Reimbursed by Client (AB-Y-AN+AQ)', 'UP/OP CASES', 'AAN Settlement Details - Date', 'AAN Settlement Details - Chq.No.',
                                      'AAN Settlement Details - Amount(RM)', 'Under / (Over) Paid to Hosp (AB - AG)', 'DBMY21', 'Designated / Borrowing Bank Account - Bank 2',
                                      'Designated / Borrowing Bank Account - Bank 3', 'Designated / Borrowing Bank Account - Bank 4', 'Designated / Borrowing Bank Account - Bank 5',
                                      'Designated / Borrowing Bank Account - Bank 6', 'Designated / Borrowing Bank Account - Bank 7', 'Repayment Date', 'TPA Billing Details - Date',
                                      'TPA Billing Details - Inv. No.', 'TPA Billing Details - Amount(RM)', 'Reimbursement Details - Date', 'Reimbursement Details - Chq.No.',
                                      'Reimbursement Details - Amount(RM)', 'Adjustment Details - Date', 'Adjustment Details - Doc.No', 'Adjustment Details - Amount(RM)',
                                      'Hospital OR Details - OR Date', 'Hospital OR Details - OR No.', 'Hospital OR Details - Amount(RM)', '''Date of OR Rec'd''', 'Subm. Date',
                                      'Unnamed: 61', 'Unnamed: 62', 'Unnamed: 63', 'Tax Invoice Details - Date', 'Tax Invoice Details - Inv No', 'Tax Invoice Details - Case Fee',
                                      'Tax Invoice Details - GST Amount', 'Tax Invoice Details - Total'])

      # Provide the name of the range of the header to be extracted
      first_header = 'Adm. No'
      #last_header = 'Remarks'

      wb = load_workbook(filelocation)
      sheets = wb.sheetnames

      filename = get_file_name(filelocation_main)
      if(os.path.isdir(str(filename[0]) + "\\TEMPORARY\\") == False):
        createfolder(str(filename[0]) + "\\TEMPORARY\\", reqdca_base, parents = True, exist_ok = True)

      myfile = list(getListOfFiles(str(filename[0]) + "\\TEMPORARY\\"))
      mylist = pd.DataFrame(list(filter(lambda x: str(filename[1]) in x, myfile)), columns = ['Filenames'])
      if mylist.empty == True:
        shutil.move(filelocation_main, str(filename[0]) + "\\TEMPORARY\\")
        myfile = list(getListOfFiles(str(filename[0]) + "\\TEMPORARY\\"))
        mylist = pd.DataFrame(list(filter(lambda x: str(filename[1]) in x, myfile)), columns = ['Filenames'])
        filelocation = mylist['Filenames'][0]
      else:
        filelocation = mylist['Filenames'][0]

      if withpath_df['Type'][i] == 'Admission':

        # Iterate the sheets available in the workbook
        a = 0
        for a in range(len(sheets)):

          # Get the cell position of the first and last header
          start_cell = search_excel(first_header, a, wb)
          end_cell = search_excel(last_header, a, wb)

          # Search for the required sheet
          if "dca" == sheets[a].lower():
            dca_df = pd.DataFrame(pd.read_excel(filelocation, sheet_name = a, skiprows = start_cell[1]-1, nrows = 0))
            dca_sheet = a
            wb.active = a
            activesheet = wb.active
            start_cell_DCA = start_cell

          # IF CASE ID IS ADMISSION:
          # Check if B2B is checked in MARC. If got checked, take Bord Num from B2B.

          elif "cashless" == sheets[a].lower() and withpath_df['B2B'][i] == 0:

            # Search the case ID in excel sheet and copy the row in searched dataframe. (get hte Bill Num from somehwere here.....)
            searched_df = pd.read_excel(filelocation, sheet_name = a, skiprows = start_cell[1]-1)
            searched_df = searched_df[pd.notnull(searched_df['Adm. No'])]
            searched_df = searched_df[searched_df['Adm. No'].astype(str).str.contains("%s" % withpath_df['Case ID'][i])]

            # Append the result in main dataframe
            main_df = main_df.append(searched_df, sort = False)

          elif "b2b" == sheets[a].lower() and withpath_df['B2B'][i] == 1:

            # Search the case ID in excel sheet and copy the row in searched dataframe
            searched_df = pd.read_excel(filelocation, sheet_name = a, skiprows = start_cell[1]-1)
            searched_df = searched_df.dropna(subset = ['Adm. No'])
            searched_df = searched_df[searched_df['Adm. No'].astype(str).str.contains("%s" % withpath_df['Case ID'][i])]

            # Append the result in main dataframe
            main_df = main_df.append(searched_df, sort = False)

          else:
            continue
      else:

        # Iterate the sheets available in the workbook
        a = 0
        for a in range(len(sheets)):

          # Get the cell position of the first and last header
          start_cell = search_excel(first_header, a, wb)
          end_cell = search_excel(last_header, a, wb)

          # Search for the required sheet
          if "dca" == sheets[a].lower():
            dca_df = pd.DataFrame(pd.read_excel(filelocation, sheet_name = a, skiprows = start_cell[1]-1, nrows = 0)
            dca_sheet = a
            wb.active = a
            activesheet = wb.active
            start_cell_DCA = start_cell

          # IF CASE ID IS POST:
          # Check in MARC. If MARC has value in Comp. Pay OR Back-2-Back is "Yes", take data from Post B2B.
          elif "post b2b" == sheets[a].lower() and withpath_df['B2B'][i] == 1 and withpath_df['Company Payment Amount'][i] != 0:

            # Search the case ID in excel sheet and copy the row in searched dataframe
            searched_df = pd.read_excel(filelocation, sheet_name = a, skiprows = start_cell[1]-1)
            searched_df = searched_df.dropna(subset = ['Adm. No'])
            searched_df = searched_df[searched_df['Adm. No'].astype(str).str.contains("%s" % withpath_df['Sub Case ID'][i])]

            # Append the result in main dataframe
            main_df = main_df.append(searched_df, sort = False)

          elif "post" == sheets[a].lower() and withpath_df['B2B'][i] == 0 and withpath_df['Company Payment Amount'][i] == 0:

            # Search the case ID in excel sheet and copy the row in searched dataframe
            searched_df = pd.read_excel(filelocation, sheet_name = a, skiprows = start_cell[1]-1)
            searched_df = searched_df.dropna(subset = ['Adm. No'])
            searched_df = searched_df[searched_df['Adm. No'].astype(str).str.contains("%s" % withpath_df['Sub Case ID'][i])]

            # Append the result in main dataframe
            main_df = main_df.append(searched_df, sort = False)

          # Check both the "Comp. Pay" and "Back-2-Back" columns. If either one is TRUE, stop and return notification to user. If both are TRUE, proceed as usual.
          else:
            continue

      main_df = pd.DataFrame(main_df.reset_index().drop(columns = "index"))
      main_df = main_df.drop_duplicates()

      # Search for latest row with value in DCA sheet
      last_row = int(search_last_row(start_cell_DCA, filelocation, int(dca_sheet), wb, first_header))
      activesheet.protection.sheet = False
      activesheet.protection.password = '560470'
      activesheet.protection.disable()

      # Iterate the main dataframe
      b = 0
      for b in range(main_df.shape[0]):

        # Compare the header and copy value in Main Dataframe into DCA Dataframe
        final_df = copy_df(main_df, dca_df, b)
        final_df.loc[b, 'Remarks'] = str(withpath_df['Remarks'][i])
        final_df.loc[b, 'Total'] = float(withpath_df['Amount'][i])
        final_df.loc[b, 'Submission Date'] = time.strftime("%d/%m/%Y")
        final_df.loc[b, 'OB Registered Date'] = main_df['OB Registered Date'][b].strftime("%d/%m-%Y")

        disbursement_listing = main_df['Disbursement Listing'][0].split(" ")
        if withpath_df['Type'][i] == 'Admission':
          final_df.loc[b, 'Disbursement Listing'] = str(disbursement_listing[b]) + " (C)"
        else:
          final_df.loc[b, 'Disbursement Listing'] = str(disbursement_listing[b]) + " (P)"

        myfinal = final_df
        myfinal = myfinal.drop(columns = "OB Registered Date")

        last_row = last_row + 1

        # Update the excel sheet row by row
        filestatus = is_locked(filelocation)
        if str(filestatus) == 'False':

          for n in range(0, 19):
            activesheet[last_row][n].protection = Protection(locked = False, hidden=False)

          activesheet[last_row][0].value = ''
          activesheet[last_row][1].value = int(myfinal['Adm. No'][0])
          activesheet[last_row][2].value = int(myfinal['File No.'][0])
          activesheet[last_row][3].value = myfinal['Disbursement Listing'][0]
          activesheet[last_row][4].value = myfinal['Disbursement Claims'][0]
          activesheet[last_row][5].value = myfinal['''Member's Name'''][0]
          activesheet[last_row][6].value = myfinal['Policy No'][0]
          activesheet[last_row][7].value = myfinal['Bill No'][0]
          activesheet[last_row][8].value = myfinal['Adm. Date'][0]
          activesheet[last_row][9].value = myfinal['Dis. Date'][0]
          activesheet[last_row][10].value = myfinal['Hospital'][0]
          activesheet[last_row][11].value = ''
          activesheet[last_row][12].value = myfinal['Insurance'][0]
          activesheet[last_row][13].value = myfinal['Patient'][0]
          activesheet[last_row][14].value = myfinal['Total'][0]
          activesheet[last_row][15].value = myfinal['Submission Date'][0]
          activesheet[last_row][16].value = myfinal['Bill received'][0]
          activesheet[last_row][17].value = myfinal['CSU Received Date'][0]
          activesheet[last_row][18].value = myfinal['Remarks'][0]
          wb.save(str(filelocation_main))
          wb.close()
          audit_log("REQ/DCA - Updating CM document completed for Case ID [ %s ]" % str(withpath_df['Case ID'][i]), "Completed...", reqdca_base)
          send(reqdca_base, reqdca_base.email, "REQ/DCA - Updating CM document completed.", "Hi, <br/><br/><br/>CM document has been updated. Please check. <br/>Reference Number: " + str(reqdca_base.guid) + "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
          audit_log("REQ/DCA - User has been informed through email.", "Completed...", reqdca_base)
        else:
          curtask = pd.DataFrame(os.popen('tasklist').readlines()).to_json(orient = 'index')
          if "excel.exe" in curtask:
            os.system("taskkill /f /im  excel.exe")
            filestatus = is_locked(filelocation)
            if str(filestatus) == 'False':
              update_entry(activesheet, myfinal, last_row, b, start_column = 2)
              wb.save(str(filelocation_main))
              wb.close()
              audit_log("REQ/DCA - Updating CM document completed for Case ID [ %s ]" % str(withpath_df['Case ID'][i]), "Completed...", reqdca_base)
              send(reqdca_base, reqdca_base.email, "REQ/DCA - Updating CM document completed.", "Hi, <br/><br/><br/>CM document has been updated. Please check. <br/>Reference Number: " + str(reqdca_base.guid) + "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
              audit_log("REQ/DCA - User has been informed through email.", "Completed...", reqdca_base)
            else:
              print("- Updating CM document: %s cannot be done. File is currently still LOCKED / USED." % filelocation)
              audit_log("REQ/DCA - Updating CM document: %s cannot be done. File is currently still LOCKED / USED." % filelocation, "Completed...", reqdca_base)
              send(reqdca_base, reqdca_base.email, "REQ/DCA - Updating CM document: %s cannot be done. File is currently still LOCKED / USED." % filelocation, "Hi, <br/><br/><br/>. <br/>Reference Number: " + str(reqdca_base.guid)+"<br/><br/><br/>Regards,<br/>Robotic Process Automation")
              audit_log("REQ/DCA - User has been informed through email.", "Completed...", reqdca_base)
              print("- User has been informed through email.")
          else:
            print("- Updating CM document: %s cannot be done. File is currently still LOCKED / USED." % filelocation)
            audit_log("REQ/DCA - Updating CM document: %s cannot be done. File is currently still LOCKED / USED." % filelocation, "Completed...", reqdca_base)
            send(reqdca_base, reqdca_base.email, "REQ/DCA - Updating CM document: %s cannot be done. File is currently still LOCKED / USED." % filelocation, "Hi, <br/><br/><br/>. <br/>Reference Number: " + str(reqdca_base.guid)+"<br/><br/><br/>Regards,<br/>Robotic Process Automation")
            audit_log("REQ/DCA - User has been informed through email.", "Completed...", reqdca_base)
            print("- User has been informed through email.")
            pass

        # Delete original file
        os.remove(filelocation)

        # Update the the value in the dataframe with value from user input
        final_df.insert(18, 'SAGE ID', "")
        final_df.insert(19, 'Sub Case ID', "")
        final_df.insert(20, 'Member Ref. ID', "")
        final_df.insert(21, 'Type', "")
        final_df.insert(22, 'Plan Type', "")
        final_merged_df = update_dataframe(final_df, withpath_df, i, b)

        # Update in the database
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
            "Designated  / Borrowing Bank AccountDBMY29.3": "", "Designated  / Borrowing Bank AccountDBMY29.4": "", "TPA Billing DetailsDate": "",
            "TPA Billing DetailsInv. No.": "", "TPA Billing DetailsAmount (RM)": "", "Reimbursement DetailsDate": "", "Reimbursement DetailsChq. No.": "",
            "Reimbursement DetailsAmount (RM)": "", "Adjustment DetailsDate": "", "Adjustment DetailsDoc. No.": "", "Adjustment DetailsAmount (RM)": "",
            "Hospital OR DetailsOR Date": "", "Hospital OR DetailsOR No.": "", "Hospital OR DetailsAmount (RM)": "", "Hospital OR DetailsDate of OR Rec'd": "",
            "Hospital OR DetailsSubm. Date": "", "Tax Invoice DetailsDate": "", "Tax Invoice DetailsInv No": "", "Tax Invoice DetailsCase Fee": "",
            "Tax Invoice DetailsGST Amount": "", "Tax Invoice DetailsTotal": ""
            }
        my_cm = { 'Data': []}

        my_dict["SAGE DOC ID"] = final_merged_df.iloc[b]['SAGE ID']
        my_dict["Adm. No"] = final_merged_df.iloc[b]['Adm. No']
        my_dict["File No."] = final_merged_df.iloc[b]['File No.']
        my_dict["Disbursement Listing"] = str(final_merged_df.iloc[b]['Disbursement Listing'])
        my_dict["Disbursement Claims"] = final_merged_df.iloc[b]['Disbursement Claims']
        my_dict['''Member's Name'''] = final_merged_df.iloc[b]['''Member's Name''']
        my_dict["Policy No"] = final_merged_df.iloc[b]['Policy No']
        my_dict["Plan Type"] = final_merged_df.iloc[b]['Plan Type']
        my_dict["Adm. Date"] = final_merged_df.iloc[b]['Adm. Date']

        if withpath_df['Type'][i] == 'Admission':
          my_dict["Dis. Date"] = final_merged_df.iloc[b]['Dis. Date'] # For ADM Case
        else:
          my_dict["Dis. Date"] = '' # For POST Case
          final_merged_df.iloc[b]['Dis. Date'] = ''

        my_dict["Hospital"] = final_merged_df.iloc[b]['Hospital']
        my_dict["Diagnosis"] = final_merged_df.iloc[b]['Diagnosis']
        my_dict["Insurance"] = ''
        my_dict["Patient"] = ''
        my_dict["Discrepancy"] = ''
        my_dict["Total"] = final_merged_df.iloc[b]['Total']
        my_dict["Submission Date"] = final_merged_df.iloc[b]['Submission Date']
        my_dict["OB Received Date"] = final_merged_df.iloc[b]['Bill received']
        my_dict["OB Registered Date"] = final_merged_df.iloc[b]['OB Registered Date']
        my_dict["CSU Received Date"] = final_merged_df.iloc[b]['CSU Received Date']
        my_dict["Primary GL No"] = ''
        my_dict["Remarks"] = ''
        my_dict["Payment Due Date"] = ''
        my_dict["Aging (days)"] = ''
        my_dict["SCAN (TAT)"] = ''
        my_dict["CSU (TAT)"] = ''
        my_dict["DB BANK CODE"] = ''
        my_dict["Under / (Over) Claimed fr Client (AA-J-AM+AP)"] = ''
        my_dict["Cheque DetailsBank-in Date"] = ''
        my_dict["Cheque DetailsChq.No."] = ''
        my_dict["Cheque DetailsAmount (RM)"] = ''
        my_dict["Hospital Bills DetailsDate"] = ''
        my_dict["Hospital Bills DetailsBill No."] = ''
        my_dict["Hospital Bills DetailsAmount (RM)"] = ''
        my_dict["Under / (Over) Reimbursed by Client(AB-Y-AN+AQ)"] = ''
        my_dict["UP/OP CASES"] = ''
        my_dict["AAN Settlement DetailsDate"] = ''
        my_dict["AAN Settlement DetailsChq.No."] = ''
        my_dict["AAN Settlement DetailsAmount (RM)"] = ''
        my_dict["Under / (Over) Paid to Hosp (AB - AG)"] = ''
        my_dict["Designated  / Borrowing Bank AccountOCBC 16"] = ''
        my_dict["Designated  / Borrowing Bank AccountBIMB 1"] = ''
        my_dict["Designated  / Borrowing Bank AccountDBMY29"] = ''
        my_dict["Designated  / Borrowing Bank AccountDBMY29.1"] = ''
        my_dict["Designated  / Borrowing Bank AccountDBMY29.2"] = ''
        my_dict["Designated  / Borrowing Bank AccountDBMY29.3"] = ''
        my_dict["Designated  / Borrowing Bank AccountDBMY29.4"] = ''
        my_dict["TPA Billing DetailsDate"] = ''
        my_dict["TPA Billing DetailsInv. No."] = ''
        my_dict["TPA Billing DetailsAmount (RM)"] = ''
        my_dict["Reimbursement DetailsDate"] = ''
        my_dict["Reimbursement DetailsChq. No."] = ''
        my_dict["Reimbursement DetailsAmount (RM)"] = ''
        my_dict["Adjustment DetailsDate"] = ''
        my_dict["Adjustment DetailsDoc. No."] = ''
        my_dict["Adjustment DetailsAmount (RM)"] = ''
        my_dict["Hospital OR DetailsOR Date"] = ''
        my_dict["Hospital OR DetailsOR No."] = ''
        my_dict["Hospital OR DetailsAmount (RM)"] = ''
        my_dict["Hospital OR DetailsDate of OR Rec'd"] = ''
        my_dict["Hospital OR DetailsSubm. Date"] = ''
        my_dict["Tax Invoice DetailsDate"] = ''
        my_dict["Tax Invoice DetailsInv No"] = ''
        my_dict["Tax Invoice DetailsCase Fee"] = ''
        my_dict["Tax Invoice DetailsGST Amount"] = ''
        my_dict["Tax Invoice DetailsTotal"] = ''

        # Updating the table in MSSQL database
        client_id = str(withpath_df['Client ID'][i])
        doctype = str('DCA')
        list_db = my_cm["Data"]
        list_db.append(my_dict.copy())
        result = json.dumps(my_cm, default = str)
        try:
          conn1 = MSSqlConnector
          query = '''SELECT * FROM cba.client_master WHERE year = %s and client_id = %s and invoice_type = %s'''
          params = (current_year, client_id, doctype)
          fetched_cm = mssql_get_df_by_query(query, params, reqdca_base, conn1)

          if fetched_cm.empty == True:
            cur.execute('''INSERT INTO [cba].[client_master] (year, client_id, invoice_type, content, createdby, modifiedby, createdon, modifiedon, url)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)''', (current_year, client_id, doctype, result, reqdca_base.email, reqdca_base.email, current_date, current_date, filelocation))
            conn.commit()
            audit_log("REQ/DCA - CM data has been updated in the database for Case ID [ %s ]. Please check." % str(withpath_df['Case ID'][i]), "Completed...", reqdca_base)
          else:
            cur.execute("SELECT content FROM cba.client_master WHERE year = %s and client_id = %s and invoice_type = %s",(current_year, client_id, doctype))
            dictionary_db = json.loads(cur.fetchall()[0][0])
            dictionary_db["Data"].append(my_dict.copy())
            result = json.dumps(dictionary_db, default = str)
            cur.execute("UPDATE cba.client_master SET content = %s, modifiedby = %s, modifiedon = %s WHERE year = %s and client_id = %s and invoice_type = %s",(result, reqdca_base.email, current_date, current_year, client_id, doctype))
            conn.commit()
        except Exception as error:
          logging("REQDCA - Unable to update CM table in database.", error, reqdca_base)
      final_df2 = final_df2.append(final_merged_df, sort = False)

    # Close database connection
    cur.close()
    return final_df2
  except Exception as error:
    pass

