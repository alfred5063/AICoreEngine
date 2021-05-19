from utils.audit_trail import audit_log

def check_principal(column_to_validate):
  audit_log('check_principal', 'check principal value to validate for get action code.')
  return 'NRIC' if column_to_validate == 'Principal' else column_to_validate
