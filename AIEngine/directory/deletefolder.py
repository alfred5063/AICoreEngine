import os
import sys
import shutil
from utils.logging import logging
from utils.audit_trail import audit_log
#  """ Purpose: This function is to delete folder
#      Author: Yeoh Eik Den
#      Created on: 17 May 2019"""

def deletefolder(path, base):
  audit_log('deletefolder', 'Delete folder.', base)
  try:
    shutil.rmtree(path)
  except OSError as e:
    logging('deletefolder', e, base)
    print("Error: %s - %s." % (e.filename, e.strerror))

#example:
#folderpath = "C:\\Asia-Assistance\\DM\\new\\Folder\\"
#deletefolder(folderpath)

