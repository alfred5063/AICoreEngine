#!/usr/bin/python
# FINAL SCRIPT updated as of 19th May 2020
# Workflow - Finance SOA
# Version 1

# Declare Python libraries needed for this script
import sys
import pandas as pd
import time
from automagica import *
from utils.Session import session
from connector.dbconfig import read_db_config
from utils.audit_trail import audit_log
from utils.logging import logging
from automagic.automagica_sql_setting import *
from utils.application import *
from automagic.automagica_common import *

#base = soarpa_base

def grab_inpatient(browser):
  app = get_application_url('marc_search_inpatient')
  gl_no = client = invoice_amount = release_payment = register_date = remarks = ''
  try:
    Wait(1)
    gl_no = browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'case_id')).get_attribute('value')
    client = browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'client')).get_attribute("value")
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'ws_tab')).click()
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'doc_info')).click()
    invoice_amount = browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'invoice_amount')).get_attribute("value")
    register_date = browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'register_date')).get_attribute("value")
  except Exception as error:
    remarks = 'Bill not received. Please provide CTC'
    pass

  return gl_no, client, invoice_amount, register_date, remarks

def grab_outpatient(browser):
  app = get_application_url('marc_search_outpatient')
  case_id = subcase_id = client = invoice_amount = release_payment = register_date = remarks = ''
  try:
    case_id = browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'case_id')).get_attribute("value")
    subcase_id = browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'subcase_id')).get_attribute("value")
    gl_no = "%s-%s" %(case_id,subcase_id)
    client = browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'client')).get_attribute("value")
    invoice_amount = browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'invoice_amount')).get_attribute("value")
    pending_payment = browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'pending_payment')).is_displayed()
    release_payment = browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'release_payment')).is_displayed()
    if pending_payment and release_payment:
      register_date = browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'register_date')).get_attribute("value")
    else:
      remarks = 'Pending Medical Enquiry'
  except Exception as error:
    remarks = 'Bill not received. Please provide CTC'
    pass

  return gl_no, client, invoice_amount, register_date, remarks

def query_Marc(soa_df, base):
  audit_log('Query Marc', 'Begin Querying Marc for information', base)
  print('Begin Querying Marc for information')
  try:
    workflow = 'SOA'
    browser = initBrowser(base, workflow)
    bill_not_found = ''
    count = 0
    count_exist = 0
    i = 0
    for i in range(soa_df.shape[0]):
      if soa_df.iloc[i]['Remarks'] == None or pd.isnull(soa_df.iloc[i]['Remarks']) and pd.isnull(soa_df.iloc[i]['GL No']) and pd.isnull(soa_df.iloc[i]['GL Ref']) == True:
        print("TRUE")
        count +=1
        app = search_obnum_outpatient(browser, base, soa_df.iloc[i]['Invoice ID'])
        if(app == -1):
          app = search_obnum_inpatient(browser, base, soa_df.iloc[i]['Invoice ID'])
          if(app == -1):
            soa_df.loc[i,'Remarks'] = 'Bill not received. Please provide CTC'
            print(soa_df.iloc[i]['Invoice ID']+" Doesnt Exist in Marc")
          else:
            count_exist+=1
            print(soa_df.iloc[i]['Invoice ID'] + " Found in Inpatient")
            soa_df.loc[i, 'GL No'], soa_df.loc[i, 'Client'], soa_df.loc[i, 'Invoice Amount'], soa_df.loc[i, 'Register Date'], soa_df.loc[i, 'Remarks'] = grab_inpatient(browser)
        else:
          count_exist+=1
          print(soa_df.iloc[i]['Invoice ID'] + " Found in Outpatient")
          soa_df.loc[i, 'GL No'], soa_df.loc[i, 'Client'], soa_df.loc[i, 'Invoice Amount'], soa_df.loc[i, 'Register Date'], soa_df.loc[i, 'Remarks'] = grab_outpatient(browser)

        if(soa_df.loc[i, 'Remarks'] == 'Bill not received. Please provide CTC'):
          bill_not_found = bill_not_found + soa_df.iloc[i]['Invoice ID'] + ' ,'
          #log_invoice = str(soa_df.iloc[i]['Invoice ID'])
          #audit_log('%s not found in MARC', log_invoice, base)

    print("Total Bill :%s" % count)
    print("Bill Found :%s" % count_exist)
    browser.close()
    bill_not_found = bill_not_found + '\b not found in Marc'
    audit_log('Query Marc', 'Completed...', base)
    audit_log('Bill Not Updated', bill_not_found, base)
    print('Query Marc Completed...')
    print(bill_not_found)
  except Exception as error:
    browser.close()
    logging('Bill Not Updated', error, base)
    print('Bill Not Updated')
    print(error)
  return soa_df


