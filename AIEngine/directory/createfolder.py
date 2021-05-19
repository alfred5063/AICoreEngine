from pathlib import Path
from utils.logging import logging
from utils.audit_trail import audit_log
#  """ Purpose: This class is use for directory create folder
#      Author: Yeoh Eik Den
#      Created on: 15 May 2019"""

def createfolder(path, base, **kwargs):
  audit_log('createfolder', 'Creating folder.', base)
  try:
    Path(path).mkdir(**kwargs)
  except OSError as error:
    print ("Creation of the directory %s failed" % path)
    logging('createfolder', error, base)
  else:
    print ("Successfully created the directory %s " % path)

#example code:
#path ='C:\\Asia-Assistance\\DM\\New\\Folder\\'
#createfolder(path)
