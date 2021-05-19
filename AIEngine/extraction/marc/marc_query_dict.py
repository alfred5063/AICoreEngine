from extraction.marc.query_marc_df import query_as_df, sp_query_as_df
from transformation.manageRow import remove_duplicate
from utils.audit_trail import audit_log
from utils.logging import logging
import sys, os

parameter_mapper = {
  'NRIC': ['Policy Num'],
  'OtherIC': ['Policy Num'],
  'Policy Num': ['Policy Num'],
  'policy num':['policy num'],
  'Principal': ['Policy Num'],
  'Plan Expiry Date': ['Policy Num'],
  'DOB': ['Policy Num'],
  'Employee ID': ['Policy Num'],
  'Relationship': ['Policy Num'],
}

def get_uniqueue_list_from_column(df, column_to_validate, base):
  audit_log('get_uniqueue_list_from_column', 'Get a unique list from excel column for comparison.', base)
  try:
    print("Select unique value from column: %s" % column_to_validate)
    audit_log('Processing data', 'Get unique value from column: %s' % column_to_validate, base)
    not_null = df[df[column_to_validate].notnull() == True][column_to_validate]
    unique_df = not_null.dt.date.tolist() if not_null.dtype == 'datetime64[ns]' else not_null.unique().tolist()
    return unique_df
  except Exception as error:
    logging('get_uniqueue_list_from_column', error, base)
    raise error

def get_uniqueue_list_from_multiple_columns(df, columns, base):
  audit_log('get_uniqueue_list_from_multiple_columns', 'Get multiple columns to get unique value.', base)
  try:
    list = []
    for column in columns:
      list.append(get_uniqueue_list_from_column(df, column, base))
  except Exception as error:
    logging('get_uniqueue_list_from_multiple_columns', error, base)
  return list

initial_query = '''SELECT impp.`imppId`, mm.`mmId`, mm.`mmFullName`, mm.`mmNRIC`, mm.`mmOtherIc`, mm.`mmDOB`,
                 impp.`imppRelationship`, impp.`imppExternalRefId`, impp.`imppEmployeeId`, impp.`imppFirstJoinDate`,
                 impp.`imppPlanAttachDate`, impp.`imppPlanExpiryDate`, impp.`imppCancelDate`, pol.`inptPolPolicyNum`,
                 `inptPolOwner`, plan.`inptPlanExtCode`, `imppInptPolicyPlanId`, polplan.`ippInptPlanId`
                FROM member mm INNER JOIN `inpt_member_policyplan` impp ON impp.`imppMemberId` = mm.mmid
                INNER JOIN `inpt_policy_plan` polplan ON polplan.`ippId` = impp.`imppInptPolicyPlanId`
                INNER JOIN inpt_plan plan ON plan.`inptPlanId` = polplan.`ippInptPlanId`
                INNER JOIN `inpt_policy` pol ON pol.`inptPolId` = polplan.`ippInptPolicyId`
                {}
                '''

def base_query(params, base):
  audit_log('base_query', 'Retrieve base query for membership data from MARC.', base)
  try:
    query = initial_query
    query1 = query.format('WHERE pol.`inptPolPolicyNum` IN ({})  AND mm.`mmNRIC` IN ({}) AND mm.`mmFullName` in ({})')
    query2 = query.format('WHERE pol.`inptPolPolicyNum` IN ({})  AND mm.`mmOtherIc` in ({}) AND mm.`mmFullName` in ({})')
    valid = params[3] and len(params[3]) > 0
    result_query = query1 + ' UNION ' + query2 if valid else query1
    param_list = [params[0], params[1], params[2], params[0], params[1], params[3]] if valid else [params[0], params[1], params[2]]
  except Exception as error:
    logging('base_query', error, base)
  return (result_query, param_list)

def dob_query(params, base):
  audit_log('base_query', 'Retrieve base query for membership data from MARC.', base)
  try:
    query = initial_query
    query1 = query.format('WHERE pol.`inptPolPolicyNum` IN ({})  AND mm.`mmNRIC` IN ({}) AND mm.`mmFullName` in ({})')
    query2 = query.format('WHERE pol.`inptPolPolicyNum` IN ({})  AND mm.`mmOtherIc` in ({}) AND mm.`mmFullName` in ({})')
    valid = params[3] and len(params[3]) > 0
    result_query = query1 + ' UNION ' + query2 if valid else query1
    param_list = [params[0], params[2], params[0], params[1], params[3]] if valid else [params[0], params[1], params[2]]
  except Exception as error:
    logging('base_query', error, base)
  return (result_query, param_list)


def query_by_policy_num(params, base):
  audit_log('query_by_policy_num', 'Filter based on NRIC.', base)
  try:
    
    query = initial_query
    result_query = query.format('WHERE pol.`inptPolPolicyNum` IN ({})')
  except Exception as error:
    logging('query_by_policy_num', error, base)
  return (result_query, params)

def sp_query_by_policy_num(params, base):
  audit_log('query_by_policy_num', 'Filter based on NRIC.', base)
  try:
    
    query = 'query_by_policy_num'
    result_query = query
  except Exception as error:
    logging('query_by_policy_num', error, base)
  return (result_query, params)

def sp_query_principle_by_policy_num(params, base):
  audit_log('query_principle_by_policy_num', 'Filter based on NRIC.', base)
  try:
    query = 'query_principle_by_policy_num'
    result_query = query
  except Exception as error:
    logging('query_principle_by_policy_num', error, base)
  return (result_query, params)

def query_principle_by_policy_num(params, base):
  audit_log('query_principle_by_policy_num', 'Filter based on NRIC.', base)
  try:
    query = '''SELECT
      CASE WHEN impp.`imppPrincipalId` IS NULL THEN impp.`imppId` ELSE impp.`imppPrincipalId` END AS "imppPrincipalId",
      CASE WHEN impp.`imppPrincipalId` IS NULL THEN mm.`mmFullName` ELSE prpmm.mmFullName END AS "Principal Name MARC",
      CASE WHEN impp.`imppPrincipalId` IS NULL THEN mm.`mmNRIC` ELSE prpmm.mmNRIC END AS "Principal IC",
      CASE WHEN impp.`imppPrincipalId` IS NULL THEN mm.`mmOtherIc` ELSE prpmm.mmOtherIc END AS "Principal Other IC",
      CASE WHEN impp.`imppPrincipalId` IS NULL THEN impp.`imppExternalRefId` ELSE prpim.imppExternalRefId END AS "Principal Ext Ref",
      mm.`mmNRIC`, mm.`mmOtherIc`,
      impp.`imppId`, mm.`mmId`,impp.`imppEmployeeId`, mm.`mmFullName`, mm.`mmDOB`, mm.`mmGender`, impp.`imppRelationship`, impp.`imppExternalRefId`, impp.`imppVIP`,
      impp.`imppFirstJoinDate`, impp.`imppPlanAttachDate`, impp.`imppPlanExpiryDate`, impp.`imppCancelDate`, pol.`inptPolPolicyNum`, pol.`inptPolOwner`, polplan.ippInptpolicyId, `inptPlanExtCode`,`inptPlanName`, polplan.ippInptPlanId, impp.`imppInptPolicyPlanId`,impp.`imppCreatedDate`, impp.`imppSpecialConditions`, impp.`imppAnnualIndUtil`, impp.imppImportBatchId, impp.`imppExtensionExpiryDate`
      FROM member mm INNER JOIN `inpt_member_policyplan` impp ON impp.`imppMemberId` = mm.mmid
      INNER JOIN `inpt_policy_plan` polplan ON polplan.`ippId` = impp.`imppInptPolicyPlanId`
      INNER JOIN inpt_plan plan ON plan.`inptPlanId` = polplan.`ippInptPlanId`
      INNER JOIN `inpt_policy` pol ON pol.`inptPolId` = polplan.`ippInptPolicyId`
      LEFT OUTER JOIN `inpt_member_policyplan` prpim ON impp.imppPrincipalId = prpim.`imppId`
      LEFT OUTER JOIN member prpmm ON prpmm.mmId = prpim.imppMemberId
      {}
      '''
    result_query = query.format('WHERE pol.`inptPolPolicyNum` IN ({})')
  except Exception as error:
    logging('query_principle_by_policy_num', error, base)
  return (result_query, params)

def policy_no_query(params, base):
  audit_log('policy_no_query', 'Filter based on NRIC union Other IC or only NRIC.', base)
  try:
    query = initial_query
    query1 = query.format('WHERE mm.`mmNRIC` in ({}) AND mm.`mmFullName` IN ({})')
    query2 = query.format('WHERE mm.`mmOtherIc` in ({}) AND mm.`mmFullName` IN ({})')
    valid = params[2] and len(params[2]) > 0
    result_query = query1 + ' UNION ' + query2 if valid else query1
    param_list = [params[0], params[2], params[1],params[2]] if valid else  [params[0], params[2]]
    
  except Exception as error:
    logging('policy_no_query', error, base)
  return (result_query, param_list)

def plan_expiry_query(params, base):
  audit_log('plan_expiry_query', 'Filter based on NRIC union Other IC or only NRIC.', base)
  try:
    query = initial_query
    query1 = query.format('WHERE pol.`inptPolPolicyNum` IN ({})  AND mm.`mmNRIC` in ({}) AND mm.`mmFullName` IN ({})')
    query2 = query.format('WHERE pol.`inptPolPolicyNum` IN ({})  AND mm.`mmOtherIc` in ({}) AND mm.`mmFullName` IN ({})')
    valid = params[2] and len(params[2]) > 0
    result_query = query1 + ' UNION ' + query2 if valid else query1
    param_list = [params[0], params[1], params[0], params[2]] if valid else [params[0], params[1]]
  except Exception as error:
    logging('plan_expiry_query',error, base)
  return (result_query, param_list)

def principal_query(params, base):
  audit_log('principal_query', 'Filter based on NRIC union Other IC or only NRIC for principal.', base)
  try:
    initial_query ='''SELECT
      CASE WHEN impp.`imppPrincipalId` IS NULL THEN impp.`imppId` ELSE impp.`imppPrincipalId` END AS "imppPrincipalId",
      CASE WHEN impp.`imppPrincipalId` IS NULL THEN mm.`mmFullName` ELSE prpmm.mmFullName END AS "Principal Name MARC",
      CASE WHEN impp.`imppPrincipalId` IS NULL THEN mm.`mmNRIC` ELSE prpmm.mmNRIC END AS "Principal IC",
      CASE WHEN impp.`imppPrincipalId` IS NULL THEN mm.`mmOtherIc` ELSE prpmm.mmOtherIc END AS "Principal Other IC",
      CASE WHEN impp.`imppPrincipalId` IS NULL THEN impp.`imppExternalRefId` ELSE prpim.imppExternalRefId END AS "Principal Ext Ref",
      mm.`mmNRIC`, mm.`mmOtherIc`,
      impp.`imppId`, mm.`mmId`,impp.`imppEmployeeId`, mm.`mmFullName`, mm.`mmDOB`, mm.`mmGender`, impp.`imppRelationship`, impp.`imppExternalRefId`, impp.`imppVIP`,
      impp.`imppFirstJoinDate`, impp.`imppPlanAttachDate`, impp.`imppPlanExpiryDate`, impp.`imppCancelDate`, pol.`inptPolPolicyNum`, pol.`inptPolOwner`, polplan.ippInptpolicyId, `inptPlanExtCode`,`inptPlanName`, polplan.ippInptPlanId, impp.`imppInptPolicyPlanId`,impp.`imppCreatedDate`, impp.`imppSpecialConditions`, impp.`imppAnnualIndUtil`, impp.imppImportBatchId, impp.`imppExtensionExpiryDate`
      FROM member mm INNER JOIN `inpt_member_policyplan` impp ON impp.`imppMemberId` = mm.mmid
      INNER JOIN `inpt_policy_plan` polplan ON polplan.`ippId` = impp.`imppInptPolicyPlanId`
      INNER JOIN inpt_plan plan ON plan.`inptPlanId` = polplan.`ippInptPlanId`
      INNER JOIN `inpt_policy` pol ON pol.`inptPolId` = polplan.`ippInptPolicyId`
      LEFT OUTER JOIN `inpt_member_policyplan` prpim ON impp.imppPrincipalId = prpim.`imppId`
      LEFT OUTER JOIN member prpmm ON prpmm.mmId = prpim.imppMemberId
      {}
      '''
    query1 = initial_query.format('WHERE pol.`inptPolPolicyNum` IN ({}) AND mm.`mmNRIC` IN ({}) AND mm.`mmFullName` IN ({}) AND impp.`imppPlanExpiryDate` >= NOW() AND impp.imppCancelDate IS NULL')
    query2 = initial_query.format('WHERE pol.`inptPolPolicyNum` IN ({}) AND mm.`mmOtherIc` in ({}) AND mm.`mmFullName` IN ({}) AND impp.`imppPlanExpiryDate` >= NOW() AND impp.imppCancelDate IS NULL')
    valid = params[3] and len(params[3]) > 0
    result_query = query1 + ' UNION ' + query2 if valid else query1
    param_list = [params[0], params[1], params[3], params[0], params[2], params[3]] if valid else [params[0], params[1], params[3]]
  except Exception as error:
    logging('principal_query', error, base)
  return (result_query, param_list)

def get_member_by_NRIC(params, base):
 
  query_params = params[0]
  print(query_params)
  excel_column = params[1]
  print(excel_column)
  
  if len(excel_column) == 0:
    raise Exception(f'Column {params[1]} on input is empty')
  audit_log('get_member_by_NRIC', 'Get NRIC value from table.', base)
  try:
    marc_info = excel_db_mapper[excel_column, base]
    marc_column = marc_info['marc_column']
    #limit = len(params[1])
    query, prepared_params = marc_info['query'](query_params)
    
    #query += f'LIMIT {limit}'
    marc_result = query_as_df(query, prepared_params, base)
    marc_result['imppPlanExpiryDate'] = marc_result['imppPlanExpiryDate'].astype('str')
    marc_result['mmDOB'] = marc_result['mmDOB'].astype('str')
    #marc_result = remove_duplicate(marc_result, 'mmNRIC', keep='first')
    return marc_result, marc_column
  except Exception as error:
    exc_type, exc_obj, exc_tb = sys.exc_info()
    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
    error_message = str(error) + str(fname) + str(exc_tb.tb_lineno)
    print(error_message)
    logging('get_member_by_NRIC', error, base)
    raise Exception(error_message)

def sp_get_member_by_NRIC(params, base):
 
  query_params = params[0]
  print('query value: {0}'.format(query_params))
  excel_column = params[1]
  print('excel column: {0}'.format(excel_column))
  
  if len(excel_column) == 0:
    raise Exception(f'Column {params[1]} on input is empty')
  audit_log('get_member_by_NRIC', 'Get NRIC value from table.', base)
  try:
    marc_info = excel_db_mapper[excel_column]
    print('marc_info: {0}'.format(marc_info))
    marc_column = marc_info['marc_column']
    print('marc_column: {0}'.format(marc_column))
    #limit = len(params[1])
    query, prepared_params = marc_info['query'](query_params, base)
    print('query: {0} & parameters: {1}'.format(query, prepared_params[0]))

    columns = ['imppId','mmId','mmFullName','mmNRIC','mmOtherIc','mmDOB','imppRelationship','imppExternalRefId','imppEmployeeId','imppFirstJoinDate','imppPlanAttachDate','imppPlanExpiryDate', 'imppCancelDate','inptPolPolicyNum','inptPolOwner','inptPlanExtCode','imppInptPolicyPlanId','ippInptPlanId']
    marc_result = sp_query_as_df(query, prepared_params, columns, base)
    marc_result['imppPlanExpiryDate'] = marc_result['imppPlanExpiryDate'].astype('str')
    marc_result['mmDOB'] = marc_result['mmDOB'].astype('str')
    #marc_result = remove_duplicate(marc_result, 'mmNRIC', keep='first')
    return marc_result, marc_column
  except Exception as error:
    exc_type, exc_obj, exc_tb = sys.exc_info()
    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
    error_message = str(error) + str(fname) + str(exc_tb.tb_lineno)
    print(error_message)
    logging('get_member_by_NRIC', error, base)
    raise Exception(error_message)


excel_db_mapper = {
  'NRIC': { 'marc_column': 'mmNRIC', 'query': sp_query_by_policy_num },
  'OtherIC': { 'marc_column': 'mmOtherIc', 'query': sp_query_by_policy_num },
  'Policy Num': { 'marc_column': 'inptPolPolicyNum', 'query': sp_query_by_policy_num },
  'policy num': { 'marc_column': 'inptPolPolicyNum', 'query': sp_query_by_policy_num },
  'Principal': { 'marc_column': 'mmNRIC', 'query': sp_query_principle_by_policy_num },
  'Plan Expiry Date': { 'marc_column': 'imppPlanExpiryDate', 'query': sp_query_by_policy_num },
  'DOB': { 'marc_column': 'mmDOB', 'query': sp_query_by_policy_num },
  'Employee ID': { 'marc_column': 'imppEmployeeId', 'query': sp_query_by_policy_num },
  'Relationship': { 'marc_column': 'imppRelationship', 'query': sp_query_by_policy_num }
}


