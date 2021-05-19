#!/usr/bin/python
# FINAL SCRIPT updated as of 4th June 2020
# Workflow - CML, REQ and DCM files updates

# Declare Python libraries needed for this script
import pandas as pd
import json as json
import requests
import os, time
from datetime import datetime as dt
from transformation.update_dcmdb import update_dcmdb
#from transformation.update_reqdb import update_reqdb
#from transformation.update_cmldb import update_cmldb


def process_updatefiles():

  current_year = dt.now().year

  path = '\\dtisvr2\CBA_UAT\Result\REQDCA_Paths\1. DISBURSEMENT CLAIMS INVOICE LOG FILE'

  #CMLfile = r'\\dtisvr2\CBA_UAT\Result\REQDCA_Paths\MASTER'
  #update_cmldb(DCMfile)

  #REQfile = path + "\%s" % str(current_year) + "\Ops disbursement_REQ_%s.xlsx" % str(current_year)
  #update_reqdb(REQfile)

  DCMfile = path + "\%s" % str(current_year) + "\Disbursement Claims Master %s.xlsx" % str(current_year)
  update_dcmdb(DCMfile)
