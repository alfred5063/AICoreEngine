#!/usr/bin/python
# FINAL SCRIPT updated as of 3rd Dec 2020
# Workflow - STP-DM

# Declare Python libraries needed for this script
from utils.Session import session
from utils.audit_trail import *
from utils.logging import logging
from enum import Enum
from datetime import datetime
from connector.dbconfig import read_db_config
import os
from pathlib import Path

current_path = os.path.dirname(os.path.abspath(__file__))
dti_path = str(Path(current_path).parent)
config = read_db_config(dti_path+r'\config.ini', 'debug')
mode = config['mode']

def generate_matrix_code_by_row(dm_base, dm_properties, row, marc_result, marc_principal):
  try:
    print('Processing data for no {0} started'.format(row["no"]))
    audit_log('generate_matrix_code_by_row', 'processing row validation', dm_base, status=mode)

    #define new row
    row["action_code"] = ""
    row["fields_update"] = ""
    row["action"] = ""
    row["search_criteria"] = ""
    row["error"] = ""
    row["record_status"] = ""
    row["status"]=""

    print('define row matrix code & check import type.')
    if row["import type"] == "nan":
      logging("Import type is Null", dm_base.filename, dm_base, status=mode)
      row["error"]="Import type is null for record {0}".format(row["no"])
    else:
      print('Processing row with valid import type.')

      #convert column to string
      #row["nric"] = row_convert_to_str(dm_base, row, "nric")
      #fixed nric 12 digit format
      #row["nric"] = row_fix_nric_digit(dm_base, row, "nric")
      #row["principal nric"] = row_fix_nric_digit(dm_base, row, "principal nric")

      action_code = 0

      action_code, marc_with_policy = get_match_policynum(dm_base, row, marc_result)
      print('get policy action_code: {0}'.format(action_code))

      if action_code == 100:
        #drill down to filter plan expiry date with MARC result
        plan_expiry_action_code, marc_with_plan_expiry = get_match_expirydate(dm_base, row, marc_with_policy)
        
        action_code = action_code + plan_expiry_action_code
        print('get plan expiry action_code: {0}'.format(action_code))

        print('nric number: {0}'.format(row["nric"]))

        if row["nric"] != "":
          print('Verified nric: {0}'.format(row["nric"]))
          #drill down to filter nric with MARC result and plan expiry date
          nric_action_code, marc_with_nric = get_match_nric(dm_base, row, marc_with_plan_expiry)
          action_code = action_code + nric_action_code
          #validate otheric
          nric_otheric_action_code, marc_with_nric_otheric = get_match_otheric(dm_base, row, marc_with_nric)
          action_code = action_code + nric_otheric_action_code
          #validate dob
          nric_dob_action_code, marc_with_nric_dob = get_match_dob(dm_base, row, marc_with_nric)
          action_code = action_code + nric_dob_action_code
          #validate employee id
          nric_employeeid_action_code, marc_with_nric_employeeid = get_match_employeeid(dm_base, row, marc_with_nric)
          action_code = action_code + nric_employeeid_action_code
          #validate relationship
          nric_relationship_action_code, marc_with_nric_relationship = get_match_relationship(dm_base, row, marc_with_nric)
          action_code = action_code + nric_relationship_action_code
        else:
          print('Verified otheric: {0}'.format(row["otheric"]))
          #drill down to filter otheric with MARC result and plan expiry date
          otheric_action_code, marc_with_otheric = get_match_otheric(dm_base, row, marc_with_plan_expiry)  
          action_code = action_code + otheric_action_code
          #validate dob
          otheric_dob_action_code, marc_with_otheric_dob = get_match_dob(dm_base, row, marc_with_otheric)
          action_code = action_code + otheric_dob_action_code
          #validate employee id
          otheric_employeeid_action_code, marc_with_otheric_employeeid = get_match_employeeid(dm_base, row, marc_with_otheric)
          action_code = action_code + otheric_employeeid_action_code
          #validate relationship
          nric_relationship_action_code, marc_with_nric_relationship = get_match_relationship(dm_base, row, marc_with_otheric)
          action_code = action_code + nric_relationship_action_code
        #validate principal
        principal_action_code, marc_with_principal = get_match_principal(dm_base, row, marc_principal)
        
        action_code = action_code + principal_action_code
        print('action_code: {0}'.format(action_code))
        row["action_code"] = action_code
      else:
        row["action_code"] = 0

    print('processing data {0} ended with action_code {1}'.format(row["no"], row["action_code"]))
    audit_log('generate_matrix_code_by_row',
              'Processing data: {0} completed with action_code: {1}'.format(row["no"], row["action_code"]), dm_base, status=mode)
  except Exception as error:
    print('Error generate matrix code: {0}'.format(error))
    row["error"] = row["error"] + '<{0}>'.format(error)
  return row

#convert row to str format
def row_convert_to_str(dm_base, row, column_name):
  #print('Convert to str format: {0}'.format(row[column_name]))
  audit_log('row_convert_to_str', 'converting row format to string', dm_base, status=mode)
  row[column_name] = str(row[column_name])
  if row[column_name] != "nan":
     index = row[column_name].index('.')
     row[column_name] = row[column_name][0:index]
  return row[column_name]

def row_fix_nric_digit(dm_base, row, column_name):
  audit_log('row_fix_nric_digit', 'convert nric number to 12 digit format', dm_base, status=mode)
  if row[column_name] != "nan":
     row[column_name] = '{:0>12}'.format(str(row[column_name]))

  return row[column_name]

def get_match_policynum(dm_base, row, marc_result):
  audit_log('get_match_policynum', 'validate policy number with marc', dm_base, status=mode)
  sum = 0
  marc_policy = marc_result[(marc_result.inptPolPolicyNum == row["policy num"])]
  if len(marc_policy.index) >= 1:
    policy = matrixCode.policy
    print('Policy matrix code: {0}'.format(policy.value))
    if row["import type"] == "X":
       sum = int(policy.value * 2)
    else:
       sum = int(policy.value)
  return sum, marc_policy

def get_match_expirydate(dm_base, row, marc_result):
  audit_log('get_match_expirydate', 'validate plan expiry date with marc', dm_base, status=mode)
  sum = 0
  try:
    plan_expiry_date = datetime.strptime(row["plan expiry date"], '%d/%m/%Y')
    marc_plan_expiry = marc_result[(marc_result.imppPlanExpiryDate == str(plan_expiry_date.date()))]
    if len(marc_plan_expiry.index) >= 1:
      planexpiry = matrixCode.planexpiry
      print('Plan expiry matrix code: {0}'.format(planexpiry.value))
      if row["import type"] == "X":
         sum = int(planexpiry.value * 2)
      else:
         sum = int(planexpiry.value)
  except Exception as error:
    logging('get_match_expirydate', error, dm_base, status=mode)
  return sum, marc_plan_expiry

def get_match_nric(dm_base, row, marc_result):
  audit_log('get_match_nric', 'validate nric with marc', dm_base, status=mode)
  sum = 0
  try:
    marc_nric = marc_result[(marc_result.mmNRIC == row["nric"])]
    if len(marc_nric.index) >= 1:
      nric = matrixCode.nric
      print('nric matrix code: {0}'.format(nric.value))
      if row["import type"] == "X":
         sum = int(nric.value * 2)
      else:
         sum = int(nric.value)
  except Exception as error:
    logging('get_match_nric', error, dm_base, status=mode)
  return sum, marc_nric

def get_match_otheric(dm_base, row, marc_result):
  audit_log('get_match_otheric', 'validate otheric with marc', dm_base, status=mode)
  sum = 0
  try:
    marc_otheric = marc_result[(marc_result.mmOtherIc == row["otheric"])]
    if len(marc_otheric.index) >= 1:
      otheric = matrixCode.otheric
      print('otheric matrix code: {0}'.format(otheric.value))
      if row["import type"] == "X":
         sum = int(otheric.value * 2)
      else:
         sum = int(otheric.value)
  except Exception as error:
    logging('get_match_otheric', error, dm_base, status=mode)
  return sum, marc_otheric

def get_match_dob(dm_base, row, marc_result):
  audit_log('get_match_dob', 'validate dob with marc', dm_base, status=mode)
  sum = 0
  try:
    dob = datetime.strptime(row["dob"], '%d/%m/%Y')
    marc_dob = marc_result[(marc_result.mmDOB == str(dob.date()))]
    if len(marc_dob.index) >= 1:
      dob = matrixCode.dob
      print('dob matrix code: {0}'.format(dob.value))
      if row["import type"] == "X":
         sum = int(dob.value * 2)
      else:
         sum = int(dob.value)
  except Exception as error:
    logging('get_match_dob', error, dm_base, status=mode)
  return sum, marc_dob

def get_match_employeeid(dm_base, row, marc_result):
  audit_log('get_match_employeeid', 'validate employee id with marc', dm_base, status=mode)
  sum = 0
  try:
    marc_employee_id = marc_result[(marc_result.imppEmployeeId == row["employee id"])]
    if len(marc_employee_id.index) >= 1:
      employeeid = matrixCode.employeeid
      print('employeeid matrix code: {0}'.format(employeeid.value))
      if row["import type"] == "X":
         sum = int(employeeid.value * 2)
      else:
         sum = int(employeeid.value)
  except Exception as error:
    logging('get_match_employeeid', error, dm_base, status=mode)
  return sum, marc_employee_id

def get_match_relationship(dm_base, row, marc_result):
  audit_log('get_match_relationship', 'validate relationship with marc', dm_base, status=mode)
  sum = 0
  try:
    marc_relationship = marc_result[(marc_result.imppRelationship == row["relationship"])]
    if len(marc_relationship.index) >= 1:
      relationship = matrixCode.relationship
      print('relationship matrix code: {0}'.format(relationship.value))
      if row["import type"] == "X":
         sum = int(relationship.value * 2)
      else:
         sum = int(relationship.value)
  except Exception as error:
    logging('get_match_relationship', error, dm_base, status=mode)
  return sum, marc_relationship

def get_match_principal(dm_base, row, marc_principal):
  audit_log('get_match_principal', 'validate principal with marc', dm_base, status=mode)
  marc_principal.rename(columns={"Principal IC":"principal_ic", "Principal Name MARC":"principal_name"
                                , "Principal Other IC":"principal_otheric", "Principal Ext Ref":"principal_ext_ref"}, inplace=True)
  sum = 0
  isValid = False
  try:
    marc_principal = None
    plan_expiry_date = datetime.strptime(row["plan expiry date"], '%d/%m/%Y')
    if row["principal nric"] !="nan":
      marc_principal = marc_principal[(marc_principal.principal_ic==row["principal nric"]) & (marc_principal.mmNRIC==row["nric"]) &
                                       (marc_principal.imppPlanExpiryDate == plan_expiry_date.date()) &
                                       (marc_principal.inptPolPolicyNum == row["policy num"])]
      if len(marc_principal.index) >= 1:
        isValid = True

    else:
      if row["principal other ic"] != "nan":
        marc_principal = marc_principal[(marc_principal.principal_otheric==row["principal other ic"]) &
                                        (marc_principal.mmOtherIc==row["otheric"]) &
                                       (marc_principal.imppPlanExpiryDate == plan_expiry_date.date()) &
                                       (marc_principal.inptPolPolicyNum == row["policy num"])]
        if len(marc_principal.index) >= 1:
          isValid = True
      elif row["principal ext ref id (aka client)"] !="nan":
        marc_principal = marc_principal[(marc_principal.principal_ext_ref == row["principal ext ref id (aka client)"]) &
                                        (marc_principal.imppEmployeeId) & row["employee id"]
                                        (marc_principal.imppPlanExpiryDate == plan_expiry_date.date()) &
                                       (marc_principal.inptPolPolicyNum == row["policy num"])]
        if len(marc_principal.index) >= 1:
          isValid = True

    #validate all principal if valid then assign matrix code
    if isValid == True:
     principal = matrixCode.principal
     print('principal matrix code: {0}'.format(principal.value))
     if row["import type"] == "X":
         sum = int(principal.value * 2)
     else:
         sum = int(principal.value)

  except Exception as error:
    logging('get_match_principal', error, dm_base, status=mode)
  return sum, marc_principal

class matrixCode(Enum):
    nric = 1
    otheric = 10
    policy = 100
    principal=1000
    planexpiry=10000
    dob=100000
    employeeid=1000000
    relationship=10000000

