# Take a set of matrix codes and data and update marc accordingly.
# Written by David Yang

# Speculatively importing postgres, if matrix lookup is stored on db.
from selenium_interface.engine import RPA
from utils.logging import logging
from utils.audit_trail import audit_log


# Speculative shell for now, to be revised when formats are decided.
def update_marc(table, base, username=None, password=None):
  try:
    
    audit_log('update_marc' , 'Start Selenium', base)
    
    rpa = RPA()
    
    if rpa is None:
      return "Could not access marc"
    
    rpa.login(base, uid=username, pw=password)

    table, result = rpa.ms_update_member(table, base)
    # Check output table for new columns, status.
    rpa.logout()
  except Exception as error:
    logging('update_marc', error, base)
    raise error
  return result, table

def login_marc(base, username=None, password=None):
  try:
    audit_log('Login marc' , 'Start Selenium', base)
    rpa = RPA()
    
    if rpa is None:
      rpa.logout()
      return "Could not access marc"
    rpa.login(base, uid=username, pw=password)
  except Exception as error:
    logging('login_marc', error, base)
    raise error
  return rpa

def logout_marc(rpa):
  rpa.logout()
