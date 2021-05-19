# Declare Python libraries needed for this script
#!/usr/bin/python
# FINAL SCRIPT updated as of 12th March 2021

from extraction.marc.authenticate_marc_access import get_marc_access
from utils.Session import session
from utils.audit_trail import audit_log
from utils.logging import logging
from automagic.marc import *
from utils.notification import send
from automagic.automagica_sql_setting import *

def import_dm_to_marc(taskid, guid, step_name=None, source=None, destination=None, email=None, content=None, username=None, password=None, debug=False):
  dm_base = session.base(taskid, guid,email,step_name,username=username, password=password, filename=source)
  dm_properties = session.data_management(source, destination)
  audit_log('Initiate import process for DM to MARC', 'Completed...', dm_base)
  if debug==False:
    try:
      if password != "":
        dm_base.set_password(get_marc_access(dm_base))
    except:
      pass

  global browser

  #run import automatically into MARC
  audit_log('Import to MARC', 'Trigger browser to import the final result excel file. ', dm_base)
  app = get_application_url('marc_import')
  app_url = app.application.appUrl
  if app_url == ' ':
    app_url = None
  
  browser, flag = initBrowser(dm_base, url=app_url, username=dm_base.username, password=dm_base.password)
  audit_log('Login to MARC', 'Completed...', dm_base)
  #login to marc
  try:
    browser, import_status = upload_import(dm_base, browser, dm_properties.source, debugmode=True)
  except Exception as error:
    print(error)
    logging('upload_import', error, dm_base)
  browser.close()
  audit_log('Email notified', 'Completed...', dm_base)

  if dm_base.email != None and content == None:
    send(dm_base, dm_base.email, "RPA Task Execution Completed.", "Hi, <br/><br/><br/>You task has been completed. "+
                      "<br/>Reference Number: " + str(guid) + "<br/>"+
                      "<br/>MARC Import statust as follow: <br/>"+
                      import_status[0]+"<br/>"+
                      import_status[1]+"<br/>"+
                      import_status[2]+"<br/>"+
                      "<br/><br/>"+
                      "<br/>Regards,<br/>Robotic Process Automation")
  if dm_base.email != None and content != None:
    send(dm_base, dm_base.email, "RPA Task Execution Completed.", "Hi, <br/><br/><br/>You task has been completed. "+
                        "<br/>Reference Number: " + str(guid) + "<br/>"+
                        "<br/>MARC Import statust as follow: <br/>"+
                        import_status[0]+"<br/>"+
                        import_status[1]+"<br/>"+
                        import_status[2]+"<br/>"+
                        "<br/><br/>"+ content+
                        "<br/>Regards,<br/>Robotic Process Automation") 
  audit_log('Task Complete Process', 'Completed', dm_base)

  return 'success'

