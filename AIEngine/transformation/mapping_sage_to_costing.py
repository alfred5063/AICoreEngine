#!/usr/bin/python
# FINAL SCRIPT updated as of 27th Nov 2020
# Workflow - Finance-AR

import pandas as pd
import time
from openpyxl import load_workbook
from automagic.automagica_finance_update_max import *
from utils.audit_trail import audit_log, audit_logs_insert, cs_audit_log

def map_sage_to_costing(sage_excel, sage_type, max_base):
  audit_log("Finance sage to max - Mapping Process - Mapping Information in Sage File to Max", "Completed...", max_base)
  try:
    wb = load_workbook(sage_excel)
    invoice = pd.read_excel(sage_excel, sheet_name='Invoices')
    invoice = pd.DataFrame.reset_index(invoice[['CUSTPO', 'CNTITEM','IDINVC', 'DATEINVC', 'CODECURN']])
    invoice = invoice.drop(columns=['index'])

    invoiceDetail_df = pd.read_excel(sage_excel, sheet_name='Invoice_Detail_Optional_Fields')
    invoiceDetail_df = pd.DataFrame.reset_index(invoiceDetail_df[invoiceDetail_df.OPTFIELD=='COSTINGID'][['CNTITEM','VALUE', 'CNTLINE']])
    invoiceDetail_df = invoiceDetail_df.drop(columns=['index'])
    
    invoice_details = pd.read_excel(sage_excel, sheet_name='Invoice_Details')
    invoice_details = pd.DataFrame.reset_index(invoice_details[['CNTITEM', 'CNTLINE', 'BASETAX1']])
    invoice_details = invoice_details.drop(columns=['index'])
    
    df_merge_col = pd.merge(invoice, invoiceDetail_df, on = 'CNTITEM')
    df_merge_col = pd.merge(df_merge_col, invoice_details,  how='left', left_on=['CNTITEM', 'CNTLINE'], right_on = ['CNTITEM', 'CNTLINE'])
    
    error = []
    record_status = []
    status = []
    browser_max = login_max(max_base)
    for i in range(len(df_merge_col['CUSTPO'])):

      try:

        try:
          audit_log("Finance sage to max - Preprocessing - Mapping Information Ticket ID Found.", "Completed...", max_base)
          tixid1 = df_merge_col['CUSTPO'][i]
          audit_log("Finance sage to max - Preprocessing - Mapping Information Costing ID Found.", "Completed...", max_base)
          costingid1 = df_merge_col['VALUE'][i]
          audit_log("Finance sage to max - Preprocessing - Mapping Information Date Invoices Found.", "Completed...", max_base)
          dateinv1 = df_merge_col['DATEINVC'][i]
          audit_log("Finance sage to max - Preprocessing - Mapping Information Number Invoices Found.", "Completed...", max_base)
          numinv1 = df_merge_col['IDINVC'][i]
          audit_log("Finance sage to max - Preprocessing - Mapping Currency Found.", "Completed...", max_base)
          codecurnrc = df_merge_col['CODECURN'][i]
          audit_log("Finance sage to max - Preprocessing - Mapping Amount Found.", "Completed...", max_base)
          amtduehc = df_merge_col['BASETAX1'][i]
        except Exception as e:
          logging("Process Finance AR - Extract CostingID from sheet and grouped by the CNTITEM", e, max_base)
        
        #print(tixid1)
        #print(costingid1)
        #print(numinv1)
        #print(amtduehc)

        try:
          status, map_sage = update_max(browser_max, tixid1, costingid1, dateinv1, numinv1, codecurnrc, amtduehc, max_base)
          try:
            element_checkexist = '/html/body/div[1]/div[3]/form/div[4]/div[9]/div/div[6]/div/div[3]/div[2]/span/div[2]/div[2]/table/tbody/tr/td[4]/table/tbody/tr[6]/td[2]/input'
            tick_isbilled = browser_max.find_element_by_xpath(element_checkexist).get_attribute("checked")
            if(tick_isbilled != "true"):
              try:
                status, map_sage = update_max(browser_max, tixid1, costingid1, dateinv1, numinv1, codecurnrc, amtduehc, max_base)
              except Exception as error:
                logging("Finance sage to max - Automagica - reupdate max.", error, base)
            else:
              print('Bill is checked.')
              pass
          except Exception as error:
            logging("Finance sage to max - Automagica - checking tick_isbilled.", error, base)
        except Exception as e:
          logging("Process Finance AR - Update MAX", e, max_base)

        if str(status) == "Fail":
          error.append(cs_audit_log.audit_log('Error log',map_sage, max_base))
          record_status.append('Ticket id: {0} not update completely'.format(tixid1))
        else:
          for item in status:
            record_status.append('Ticket id: {0} update completed. Cost id status: {1}'.format(tixid1, item))

      except Exception as e:
        logging("Process Finance AR - Extract CostingID from sheet and grouped by the CNTITEM", e, max_base)
    time.sleep(10)
    close_max(map_sage)
    audit_log("Finance AR - Preprocessing - Map the Information into Max.", "Completed...", max_base)
    return 'Completed', record_status
  except Exception as e:
    logging("Process Finance AR - Error in Mapping Process", e, max_base)
    return 'Failed, error: {0}'.format(e)
