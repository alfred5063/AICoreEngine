from loading.query.query_as_df import execute_query_preprod
import pandas as pd
import numpy as np
from connector.connector import MySqlConnector_Prod2, MySqlConnector
from utils.audit_trail import audit_log
from utils.logging import logging
from extraction.crypto.crypto import decrypt_pass

def get_claim_infor(base, claim_status=None, start_visit_date=None, end_visit_date=None,
                    claim_type=None, case_type=None, start_approval_date=None, end_approval_date=None):
  #session = Session(task_id, sessionid)
  try:
    #audit_log('get_marc_access' , 'extract key from rpa to decryp process for marc to authenticate.')
    connector = MySqlConnector_Prod2
    criteria=""

    #validate visit date
    if start_visit_date !="":
      if end_visit_date !="":
        criteria = "WHERE (a.outptClaimVisitDate BETWEEN CAST('{0}' AS DATE)".format(start_visit_date)
        criteria = criteria + " AND CAST('{0}' AS DATE))".format(end_visit_date)
      else:
        criteria = "WHERE a.outptClaimVisitDate = '{0}'".format(start_visit_date)
        
    #validate claim status
    if claim_status !=None:
      #split_claim_status = claim_status.split(";")
      claim_params = ""

      for claim_item in claim_status:
        print('claims status: {0}'.format(claim_item))
        split_claim_status = claim_item.split(";")
        count = 1
        for item in split_claim_status:
          print('claim item: {0}'.format(item))
          print('count: {0}, claim_count:{1}'.format(count, len(split_claim_status)))
          if item !="":
            if count == len(split_claim_status)-1: 
              claim_params = claim_params + "'"+item+"'"
            else:
              claim_params = claim_params + "'"+item+"',"
            count = count + 1

    if criteria !="":
      print('Claim status: [{0}]'.format(claim_params))
      if claim_params !="": 
        criteria = criteria + " AND c.outptstatuseclaimname in ({0})".format(claim_params)
    else:
      if claim_params !="": 
        criteria = "WHERE c.outptstatuseclaimname in ({0})".format(claim_params)

    #validate approval date
    if criteria !="":
      if start_approval_date !="": 
        if end_approval_date !="":
          criteria = criteria + " AND (a.outptClaimApprovalDate BETWEEN CAST('{0}' AS DATE)".format(start_approval_date)
          criteria = criteria + " AND CAST('{0}' AS DATE))".format(end_approval_date)
        else:
          criteria = criteria + " AND a.outptClaimApprovalDate = '{0}'".format(start_approval_date)
    else:
      if start_approval_date !="": 
        if end_approval_date !="":
          criteria = "WHERE (a.outptClaimApprovalDate BETWEEN CAST('{0}' AS DATE)".format(start_approval_date)
          criteria = criteria + " AND CAST('{0}' AS DATE))".format(end_approval_date)
        else:
          criteria = "WHERE a.outptClaimApprovalDate = '{0}'".format(start_approval_date)

    #validate claim type
    if claim_type !=None:
      #split_claim_status = claim_status.split(";")
      claim_type_params = ""

      for claim_type_item in claim_type:
        print('claims type: {0}'.format(claim_type_item))
        split_claim_type = claim_type_item.split(";")
        count = 1
        for item in split_claim_type:
          print('claim type item: {0}'.format(item))
          item = get_claim_type_value(item)
          print('claim type item value: {0}'.format(item))
          if item !="":
            if count == len(split_claim_type)-1: 
              claim_type_params = claim_type_params + "'"+item+"'"
            else:
              claim_type_params = claim_type_params + "'"+item+"',"
            count = count + 1

    if criteria !="":
      print('Claim type: [{0}]'.format(claim_type_params))
      if claim_type_params !="": 
        criteria = criteria + " AND a.outptClaimType in ({0})".format(claim_type_params)
    else:
      if claim_type_params !="": 
        criteria = "WHERE a.outptClaimType in ({0})".format(claim_type_params)

    #validate case type
    if case_type !=None:
      #split_claim_status = claim_status.split(";")
      case_type_params = ""

      for case_type_item in case_type:
        print('case_type: {0}'.format(case_type_item))
        split_case_type = case_type_item.split(";")
        count = 1
        for item in split_case_type:
          print('case type item: {0}'.format(item))
          item = get_case_type_value(item)
          print('case type item value: {0}'.format(item))
          if item !="":
            if count == len(split_case_type)-1: 
              case_type_params = case_type_params + "'"+item+"'"
            else:
              case_type_params = case_type_params + "'"+item+"',"
            count = count + 1

    if criteria !="":
      print('case type: [{0}]'.format(case_type_params))
      if case_type_params !="": 
        criteria = criteria + " AND b.outptCaseType in ({0})".format(case_type_params)
    else:
      if case_type_params !="": 
        criteria = "WHERE b.outptCaseType in ({0})".format(case_type_params)

    query = '''SELECT a.outptClaimOutptCaseId Case_Id,
               a.outptclaimprefixid Claim_Id,
               b.outptCaseMmFullName Mm_Name,
               b.outptCaseMmNric Mm_National_Id,
               b.outptCaseMmOtherIc Mm_Other_National_Id, 
               b.outptCaseClientName Company_Or_Client_Name,
               f.clientfullname parent_company,
               b.outptCaseInsurerName Insurer_Name, 
               CASE WHEN a.outptClaimType = 'OA' THEN "Outpt - AHS" 
                    WHEN a.outptClaimType = 'OD' THEN "Outpt - Dental"
                    WHEN a.outptClaimType = 'OG' THEN "Outpt - GP"
                    WHEN a.outptClaimType = 'OM' THEN "Outpt - Maternity"
                    WHEN a.outptClaimType = 'OS' THEN "Outpt - SP"
                    WHEN a.outptClaimType = 'OT' THEN "Outpt - TCM"
                    WHEN a.outptClaimType = 'IA' THEN "Inpatient - Admission"
                    WHEN a.outptClaimType = 'ID' THEN "Inpatient - Daycare"
                    WHEN a.outptClaimType = 'IC' THEN "Inpatient - Cash Benefit"
                    WHEN a.outptClaimType = 'OP' THEN "Outpt - Physio"
                    WHEN a.outptClaimType = 'OPR' THEN "Outpt - Pre-Hospitalization"
                    WHEN a.outptClaimType = 'OPO' THEN "Outpt - Post-Hospitalization"
                    WHEN a.outptClaimType = 'OO' THEN "Outpt - Others"
                    WHEN a.outptClaimType = 'OI' THEN "Outpt - IMM"
                    WHEN a.outptClaimType = 'OA' THEN "Outpt - AHS"
                    WHEN a.outptClaimType = 'OOP' THEN "Outpt - Optical"
                    ELSE a.outptClaimType
                    END AS Claim_Type, 
                 CASE WHEN b.outptCaseType='C' THEN 'Cashless'
                                WHEN b.outptCaseType='E' THEN 'External'
                                WHEN b.outptCaseType='R' THEN 'Reimbursement'
                                WHEN b.outptCaseType='CD' THEN 'Cancel - Duplicate/Error'
                                WHEN b.outptCaseType='CC' THEN 'Cashless Chronic'
                                WHEN b.outptCaseType='RC' THEN 'Reimbursement - Chronic'
                                WHEN b.outptCaseType='CS' THEN 'Cashless - Special Arrangement'
                                WHEN b.outptCaseType='RS' THEN 'Reimbursement - Special Arrangement'
                                ELSE b.outptCaseType
                 END AS Case_Type, 
                 c.outptstatuseclaimname as Claim_Status,
                 a.outptClaimVisitDate as Visit_Date,
                 d.medprvfullname as Clinic_Or_Hospital,
                 a.outptClaimOtherMedPrvName as Other_Clinic,
                 a.outptClaimInvCurrTotalAmt as Total_Bill_Amt,
                 a.outptClaimInvCurrTotalInsAmt as Insurance_Amt,
                 a.outptClaimInvCurrTotalPtAmt as Patient_Amt,
                 a.outptClaimInvCurrTotalAsoAmt as ASO_Amt,
                 b.outptCaseCreatedDate as Case_Created_Date, 
                 a.outptClaimCreatedDate as Claim_Created_Date,
                 a.outptClaimBillReceivedDate as Bill_Received_Date,
                 a.outptClaimPrvStatementNum as Prev_Statement_No, 
                 CASE WHEN a.outptClaimApprovalType='I' THEN 'Initial GL'
                      WHEN a.outptClaimApprovalType='D' THEN 'Defer'
                      WHEN a.outptClaimApprovalType='F' THEN 'Final GL'
                      WHEN a.outptClaimApprovalType='A' THEN 'Approved'
                      WHEN a.outptClaimApprovalType='R' THEN 'Reject'
                      WHEN a.outptClaimApprovalType='C' THEN 'Canceled'
                      WHEN a.outptClaimApprovalType='X' THEN 'Canceled'
                      ELSE a.outptClaimApprovalType
                 END AS Approval_Type,
                 a.outptClaimApprovalDate as Approval_Date,
                 a.outptClaimApprovalBillCurrInsurerAmt as Approval_Amount,
                 a.outptClaimCreatedByAbbvName as Created_By,
                 a.outptClaimLastEdittedDate as Last_Edited_By,
                 a.outptclaimapprovalbyabbvname
                FROM outpt_claim a LEFT JOIN marcmy.outpt_case b ON a.outptClaimOutPtCaseId = b.outptCaseId
                LEFT OUTER JOIN marcmy.outpt_status c ON c.outptStatusid = a.outptClaimOutptStatusId
                LEFT OUTER JOIN marcmy.medical_provider d ON d.medprvid = a.outptClaimMedicalProviderId
                LEFT OUTER JOIN marcmy.client e ON e.clientid = b.outptCaseClientId
                LEFT OUTER JOIN marcmy.client f ON f.clientid = e.clientparentid {0}
                ORDER BY a.outptClaimVisitDate, a.outptclaimprefixid DESC'''.format(criteria)
    print('Query: {0}'.format(query))

    result_df = execute_query_preprod(query, base, connector)
    return result_df
  except Exception as error:
    logging('get_claim_infor', error, base)
    raise error


def get_claim_type_value(claim_type):
  if claim_type == "Outpt-AHS":
    return "OA"
  elif claim_type == "Outpt-Dental":
    return "OD"
  elif claim_type == "Outpt-Maternity":
    return "OM"
  elif claim_type == "Outpt-GP":
    return "OG"
  elif claim_type == "Outpt-SP":
    return "OS"
  elif claim_type == "Out-TCM":
    return "OT"
  elif claim_type == "Inpatient-Adminission":
    return "IA"
  elif claim_type == "Inpatient-Daycare":
    return "IC"
  elif claim_type == "Outpt-Physio":
    return "OP"
  elif claim_type == "Outpt-Pre-Hospitalization":
    return "OPR"
  elif claim_type == "Outpt-Post-Hospitalization":
    return "OPO"
  elif claim_type == "Outpt-Others":
    return "OO"
  elif claim_type == "Outpt-IMM":
    return "OI"
  elif claim_type == "Outpt-AHS":
    return "OA"
  elif claim_type == "Outpt-Optical":
    return "OOP"
  else:
    return claim_type

def get_case_type_value(case_type):
  if case_type=="Cashless":
    return "C"
  elif case_type=="External":
    return "E"
  elif case_type=="Reimbursement":
    return "R"
  elif case_type=="Cancel-Duplicate-Error":
    return "CD"
  elif case_type=="Cashless-Chronic":
    return "CC"
  elif case_type=="Reimbursement-Chronic":
    return "RC"
  elif case_type=="Cashless-SpecialArrangement":
    return "CS"
  elif case_type=="Reimbursement-SpecialArrangement":
    return "RS"
  else:
    return case_type

def get_staff_productivity(base, start_approval_date=None, end_approval_date=None):
  #session = Session(task_id, sessionid)
  try:
    #audit_log('get_marc_access' , 'extract key from rpa to decryp process for marc to authenticate.')
    connector = MySqlConnector_Prod2
    criteria=""
    criteris_spgl = ""

    #validate approval date
    if start_approval_date !="": 
        if end_approval_date !="":
          criteria = "WHERE (a.outptClaimApprovalDate BETWEEN CAST('{0}' AS DATE)".format(start_approval_date)
          criteria = criteria + " AND CAST('{0}' AS DATE))".format(end_approval_date)
          criteria_spgl = "WHERE (a.outptCaseDocCreatedDate BETWEEN CAST(  '{0}' AS DATE)".format(start_approval_date)
          criteria_spgl = criteria_spgl + "AND CAST(  '{0}' AS DATE))".format(end_approval_date)
        else:
          criteria = "WHERE a.outptClaimApprovalDate = '{0}'".format(start_approval_date)
      

    query = '''SELECT  CASE WHEN b.outptCaseType='C' THEN 'Cashless'
                                WHEN b.outptCaseType='E' THEN 'External'
                                WHEN b.outptCaseType='R' THEN 'Reimbursement'
                                WHEN b.outptCaseType='CD' THEN 'Cancel - Duplicate/Error'
                                WHEN b.outptCaseType='CC' THEN 'Cashless Chronic'
                                WHEN b.outptCaseType='RC' THEN 'Reimbursement - Chronic'
                                WHEN b.outptCaseType='CS' THEN 'Cashless - Special Arrangement'
                                WHEN b.outptCaseType='RS' THEN 'Reimbursement - Special Arrangement'
                                ELSE b.outptCaseType
                 END AS 'Case Type',
                 CASE WHEN a.outptClaimType = 'OA' THEN "Outpt - AHS" 
                    WHEN a.outptClaimType = 'OD' THEN "Outpt - Dental"
                    WHEN a.outptClaimType = 'OG' THEN "Outpt - GP"
                    WHEN a.outptClaimType = 'OM' THEN "Outpt - Maternity"
                    WHEN a.outptClaimType = 'OS' THEN "Outpt - SP"
                    WHEN a.outptClaimType = 'OT' THEN "Outpt - TCM"
                    WHEN a.outptClaimType = 'IA' THEN "Inpatient - Admission"
                    WHEN a.outptClaimType = 'ID' THEN "Inpatient - Daycare"
                    WHEN a.outptClaimType = 'IC' THEN "Inpatient - Cash Benefit"
                    WHEN a.outptClaimType = 'OP' THEN "Outpt - Physio"
                    WHEN a.outptClaimType = 'OPR' THEN "Outpt - Pre-Hospitalization"
                    WHEN a.outptClaimType = 'OPO' THEN "Outpt - Post-Hospitalization"
                    WHEN a.outptClaimType = 'OO' THEN "Outpt - Others"
                    WHEN a.outptClaimType = 'OI' THEN "Outpt - IMM"
                    WHEN a.outptClaimType = 'OA' THEN "Outpt - AHS"
                    WHEN a.outptClaimType = 'OOP' THEN "Outpt - Optical"
                    ELSE a.outptClaimType
                    END AS 'Claim Type',
                    a.outptClaimApprovalType 'Approval Type',
                     a.outptClaimOutptStatusId,
                 DATE_FORMAT(a.outptClaimApprovalDate,'%d-%m-%Y') as 'Approval Date',
                 c.userfullname as 'Staff Name',
                 a.outptclaimapprovalbyabbvname 'User ID',
                 CASE WHEN a.outptClaimInvCurrTotalAsoAmt >0 Then 1 ELSE 0 END as B2B
                FROM outpt_claim a
                LEFT JOIN marcmy.outpt_case b ON a.outptClaimOutPtCaseId = b.outptCaseId
                LEFT JOIN marcmy.user c ON c.usershortname = a.outptclaimapprovalbyabbvname
                {0}
                ORDER BY a.outptClaimApprovalDate DESC'''.format(criteria)
    print('Query: {0}'.format(query))

    query_spgl = '''SELECT c.outptClaimApprovalByAbbvName,
          c.outptClaimApprovalDate, a.outptCaseDocOutptClaimStateId, a.outptCaseDocCreatedDate,
          a.outptCaseDocCreatedByAbbvName, d.userfullname, c.outptClaimApprovalType, b.outptcaseType, c.outptClaimOutptBenefitName, c.outptClaimOutptBenefitLocalName, a.outptcasedocname
          FROM outpt_case_doc a
          JOIN outpt_case b ON a.outptCaseDocOutptCaseId=b.outptCaseId
          JOIN outpt_claim c ON c.outptClaimOutptCaseId = b.outptCaseId
          JOIN marcmy.user d ON d.usershortname = a.outptCaseDocCreatedByAbbvName
          {0}
          AND b.outptcaseType = 'C'
          AND (c.outptClaimApprovalType = 'A' OR c.outptClaimApprovalType = 'R')'''.format(criteria_spgl)

    result_df = execute_query_preprod(query, base, connector)
    result_spgl_df = execute_query_preprod(query_spgl, base, connector)
    return result_df, result_spgl_df
  except Exception as error:
    #logging('get_claim_infor', error, base)
    raise error  

