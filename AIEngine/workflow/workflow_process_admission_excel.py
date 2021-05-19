import traceback
from automagic.marc_adm import *
import pandas as pd
from transformation.adm_updatedoc import get_client_detail_to_excel,update_doc
from extraction.marc.authenticate_marc_access import get_marc_access
from utils.audit_trail import audit_log
from utils.logging import logging
from utils.Session import session
from pandas.io.json import json_normalize
def workflow_process_admission_excel(myjson, guid, stepname, email, taskid,username,password):
  adm_base = session.base(taskid, guid, email, stepname, username=username, password=password)
  step_data=myjson["step_parameters"]
  CM_path=step_data[0]["value"]
  DCM_path=step_data[1]["value"]
  BDX_file = (myjson['execution_parameters_txt']["parameters"][0]["Excel_File"])
  if password != "":
    adm_base.set_password(get_marc_access(adm_base))
  ##phase 5 update to client master and disbursement client master
  local_CMFile=CM_path+r"\\"
  try:
    browser=login_Marc(adm_base)
    bdx_type,Cliend_ID=get_client_detail_to_excel(browser,adm_base,BDX_file)
    update_doc(BDX_file,local_CMFile,DCM_path,Cliend_ID,bdx_type,adm_base)

    close_Marc(browser)
    audit_log("Update DCM and CM", "Completed...", adm_base)
    audit_log(adm_base.stepname, 'Completed', adm_base)
    return "Completed"
  except Exception as e:
    close_Marc(browser)
    logging("ADM-process admission",traceback.format_exc(),adm_base)
    audit_log("Failed-Update DCM and CM", "Completed...", adm_base)

