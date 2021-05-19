#!/usr/bin/python
# FINAL SCRIPT updated as of 8th July 2020
# Workflow - REQ/DCA

# Declare Python libraries needed for this script
import sys
import json
import pandas as pd
import numpy as np
import os
import shutil
from datetime import datetime
from connector.connector import MySqlConnector
from utils.logging import logging
from utils.audit_trail import audit_log

# Function to check the existance of Case ID and its' type (Post of Admission)
def check_validity(jsondf, reqdca_base, document):

  try:

    # Create temporary empty dataframes
    myjsondf = pd.DataFrame(jsondf)
    validated_df = pd.DataFrame(columns = ['Confirmed Type', 'Validity', 'Case ID', 'Client', 'Arrangement'])
    nonvalidated_df = pd.DataFrame(columns = ['Confirmed Type', 'Validity', 'Case ID', 'Client', 'Arrangement'])
    maindf = pd.DataFrame(columns = ['Confirmed Type', 'Validity', 'Case ID', 'Client', 'Arrangement'])

    # Create a cursor for MARC database connection
    conn = MySqlConnector()
    cur = conn.cursor()

    t = 0
    mycaseid = myjsondf['Case ID']
    for t in range(len(mycaseid)):
      caseid = int(myjsondf.iloc[t]['Case ID'])
      #type = str(myjsondf.iloc[t]['Type'])

      # Calling Stored Procedure
      parameters = [str(caseid),]
      #stored_proc = cur.callproc('query_marc_for_outpatient_validity', parameters)
      #for testing
      a = 0
      query_list = list()
      for a in range(len(parameters)):

        query = '''SELECT DISTINCT
                  cse.`inptCaseId`, cse.`inptCaseB2B`, cse.`inptCaseSpecArrg`,cse.`inptCaseHospitalName`, cse.`inptCaseAdmDate`,
                  inpt.`inptCaseMmExtRefId`, inpt.`inptCasemmPolicyNum`, inpt.`inptCaseMmPlanExtCode`, inpt.`inptCaseMmPrgSvcId`, inpt.`inptCasemmServiceName`, inpt.`inptCaseMmPrgId`, inpt.`inptCaseMmPrgName`, inpt.`inptCaseMmPlanName`, inpt.`inptCaseMmPolType`, inpt.`inptCaseMmNRIC`, inpt.`inptCaseMmName`,  inpt.`inptCaseMmClientName`, inpt.`inptCaseMmInsurerName`, ob.`iactwsCompPayAmt`, ob.`iactwsBillNum`,
                  inptcase.`inptCaseBillRegDate`, inptcase.`inptCaseBillReceivedDate`, inptcase.`inptCaseDiscDate`, subcase.`inptCasePpId`
                  FROM inpt_case cse
                  INNER JOIN `inpt_case_member` inpt ON inpt.`inptCaseMmId` = cse.`inptCaseMemberId`
                  INNER JOIN `client` cli ON cli.`clientId` = inpt.`inptCaseMmInsurerId`
                  INNER JOIN `inpt_case_csu` subcase ON subcase.`inptCasePpInptCaseId` = cse.`inptCaseId`
                  INNER JOIN `inpt_case_worksheet_value` ob ON ob.`iactwsId` = cse.`inptCaseInptCaseWorksheetValueId`
                  INNER JOIN `inpt_case` inptcase ON inptcase.`inptCaseMemberId` = cse.`inptCaseMemberId`
                  WHERE subcase.`inptCasePpId` = {0}'''.format(parameters[a])
        query_list.append(query)

      i = 0
      results = list()
      fetched_result1 = pd.DataFrame()
      for i in range(len(query_list)):
        cur.execute(query_list[i])
        result = cur.fetchall()
        results.append(result)
        fetched_result1 = fetched_result1.append(results[i])

      #End of testing code

      #for i in cur.stored_results():
      #  results = i.fetchall()
      #  print(results)

      parameters = [str(caseid),]
      #stored_proc = cur.callproc('query_marc_for_cln', parameters)
      #for j in cur.stored_results():
      #  results_cln = j.fetchall()

      #for testing 2
      b = 0
      query_list2 = list()
      for b in range(len(parameters)):
        query = '''SELECT DISTINCT acc.`inptCaseAccInptCaseId`, acc.`inptCaseAccChqNum`
                  FROM inpt_case_accounting acc
                  WHERE acc.inptCaseAccType = 'BD' AND acc.`inptCaseAccInptCaseId` = {0}'''.format(parameters[b])

                  
        query_list2.append(query)
      
      i = 0
      results_cln = pd.DataFrame()
      for i in range(len(query_list2)):
        cur.execute(query_list2[i])
        result = cur.fetchall()
        results_cln = results_cln.append(result)

      #End of testing code

      #fetched_result1 = pd.DataFrame(results)
      #results_cln = pd.DataFrame(results_cln)
      if fetched_result1.empty != True:
        fetched_result1.columns = ['CaseId', 'CaseB2B', 'CaseSpecArrg','CaseHospitalName', 'CaseAdmDate', 'CaseMmExtRefId', 'CasemmPolicyNum', 'CaseMmPlanExtCode', 'CaseMmPrgSvcId',
                                   'CasemmServiceName', 'CaseMmPrgId', 'CaseMmPrgName', 'CaseMmPlanName', 'CaseMmPolType', 'CaseMmNRIC', 'CaseMmName',  'CaseMmClientName',
                                   'CaseMmInsurerName', 'CompPayAmt', 'BillNum', 'inptCaseBillRegDate', 'inptCaseBillReceivedDate', 'inptCaseDiscDate', 'CasePpId']
        results_cln.columns = ['Sub Case ID', 'Client Listing Number']
        #results_cln = pd.DataFrame()

        if results_cln.empty != True:
          fetched_result1['CaseAccChqNum'] = results_cln['Client Listing Number']
        else:
          fetched_result1['CaseAccChqNum'] = "N/A"

        validated_df.loc[t, 'Case ID'] = caseid
        validated_df.loc[t, 'Confirmed Type'] = 'POST'
        validated_df.loc[t, 'Validity'] = "Valid Case ID"
        validated_df.loc[t, 'Client'] = fetched_result1['CaseMmInsurerName'].iloc[t]

        if fetched_result1['CaseSpecArrg'].iloc[0] == 0:
          validated_df.loc[t, 'Arrangement'] = 'No'
        else:
          validated_df.loc[t, 'Arrangement'] = 'Yes'

        audit_log("REQ/DCA - Case ID [ %s ] is valid." % caseid, "Completed...", reqdca_base)
      else:

        # Calling Stored Procedure
        #stored_proc = cur.callproc('query_marc_for_inpatient_validity', parameters)

        #for testing
        a = 0
        query_list = list()
        for a in range(len(parameters)):

          query = '''SELECT DISTINCT cse.`inptCaseId`, cse.`inptCaseB2B`, cse.`inptCaseSpecArrg`,cse.`inptCaseHospitalName`, cse.`inptCaseAdmDate`,
          inpt.`inptCaseMmExtRefId`, inpt.`inptCasemmPolicyNum`, inpt.`inptCaseMmPlanExtCode`, inpt.`inptCaseMmPrgSvcId`, inpt.`inptCasemmServiceName`, inpt.`inptCaseMmPrgId`, inpt.`inptCaseMmPrgName`, inpt.`inptCaseMmPlanName`, inpt.`inptCaseMmPolType`, inpt.`inptCaseMmNRIC`, inpt.`inptCaseMmName`,  inpt.`inptCaseMmClientName`, inpt.`inptCaseMmInsurerName`, acc.`inptCaseAccChqNum`, ob.`iactwsCompPayAmt`, ob.`iactwsBillNum`, inptcase.`inptCaseBillRegDate`, inptcase.`inptCaseBillReceivedDate`, inptcase.`inptCaseDiscDate`
          FROM inpt_case cse
          INNER JOIN `inpt_case_member` inpt ON inpt.`inptCaseMmId` = cse.`inptCaseMemberId`
          INNER JOIN `client` cli ON cli.`clientId` = inpt.`inptCaseMmInsurerId`
          INNER JOIN `inpt_case_accounting` acc ON acc.`inptCaseAccInptCaseId` = cse.`inptCaseId`
          INNER JOIN `inpt_case_worksheet_value` ob ON ob.`iactwsId` = cse.`inptCaseInptCaseWorksheetValueId`
          INNER JOIN `inpt_case` inptcase ON inptcase.`inptCaseMemberId` = cse.`inptCaseMemberId`
          WHERE acc.inptCaseAccType = 'BD' AND cse.`inptCaseId` = {0}'''.format(parameters[a])
          query_list.append(query)

        i = 0
        results = list()
        fetched_result2 = pd.DataFrame()
        for i in range(len(query_list)):
          cur.execute(query_list[i])
          result = cur.fetchall()
          results.append(result)
          fetched_result2 = fetched_result2.append(results[i])


        #End of testing code

        #for i in cur.stored_results():
        #  results = i.fetchall()
        #fetched_result2 = pd.DataFrame(results)

        if fetched_result2.empty != True:
          fetched_result2.columns = ['CaseId', 'CaseB2B', 'CaseSpecArrg', 'CaseHospitalName', 'CaseAdmDate', 'CaseMmExtRefId', 'CasemmPolicyNum', 'CaseMmPlanExtCode', 'CaseMmPrgSvcId',
                                     'CasemmServiceName', 'CaseMmPrgId', 'CaseMmPrgName', 'CaseMmPlanName', 'CaseMmPolType', 'CaseMmNRIC', 'CaseMmName', 'CaseMmClientName',
                                     'CaseMmInsurerName', 'CaseAccChqNum', 'CompPayAmt', 'BillNum', 'inptCaseBillRegDate', 'inptCaseBillReceivedDate', 'inptCaseDiscDate']
          fetched_result2['CasePpId'] = "0"
          validated_df.loc[t, 'Case ID'] = caseid
          validated_df.loc[t, 'Confirmed Type'] = 'Admission'
          validated_df.loc[t, 'Validity'] = "Valid Case ID"
          validated_df.loc[t, 'Client'] = fetched_result2['CaseMmInsurerName'][0]

          if fetched_result2['CaseSpecArrg'][0] == 0:
            validated_df.loc[t, 'Arrangement'] = 'No'
          else:
            validated_df.loc[t, 'Arrangement'] = 'Yes'

          audit_log("REQ/DCA - Case ID [ %s ] is valid." % caseid, "Completed...", reqdca_base)
        else:
          nonvalidated_df.loc[t, 'Case ID'] = caseid
          nonvalidated_df.loc[t, 'Confirmed Type'] = 'N/A'
          nonvalidated_df.loc[t, 'Validity'] = "Invalid Case ID. Cannot be found in MARC. Please check."
          nonvalidated_df.loc[t, 'Client'] = "Unknown. Please check."
          nonvalidated_df.loc[t, 'Arrangement'] = "Unknown. Please check."
          audit_log("REQ/DCA - Case ID [ %s ] cannot be found in MARC." % caseid, "Completed...", reqdca_base)

    # close the communication with the MARC database
    cur.close()
    conn.close()

    # Consolidated two dataframes
    maindf = validated_df.append(nonvalidated_df, ignore_index = True, sort = False)
    maindf = maindf.drop_duplicates()
    myjsondf['Case ID'] = pd.to_numeric(myjsondf['Case ID'])
    maindf['Case ID'] = pd.to_numeric(maindf['Case ID'])
    maindf_details = pd.merge(left = maindf, right = myjsondf, left_on = maindf['Case ID'], right_on = myjsondf['Case ID'])
    maindf_details = maindf_details.drop(columns = ['Case ID_x', 'Case ID_y', 'Type'])
    maindf_details = maindf_details[['key_0', 'Amount', 'Reason', 'Remarks', 'Confirmed Type', 'Validity', 'Client', 'Arrangement']]
    maindf_details.columns = ['Case ID', 'Amount', 'Reason', 'Remarks', 'Type', 'Validity', 'Client', 'Arrangement']
    pd.set_option("max_columns", None)
    audit_log("REQ/DCA - Case ID validation check.", "Completed...", reqdca_base)
    return maindf_details

  except Exception as error:
    logging("REQDCA - process_REQDCA", error, reqdca_base)
