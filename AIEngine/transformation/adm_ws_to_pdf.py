#!/usr/bin/python
# FINAL SCRIPT updated as of 31th October 2019
# Workflow - CBA/ADMISSION
# Version 1
from automagic.marc_adm import *
from transformation.adm_connect_db import *
from transformation.adm_bdx_manipulation import get_Lastest_File,get_all_File
#from transformation.adm_excel_to_pdf import excel_to_pdf
from transformation.adm_cba_status import adm
import os
import shutil
from datetime import date,datetime,timedelta
from directory.createfolder import createfolder
from shutil import copyfile


def ws_to_pdf(object_list,browser,download_path,print_path):
  object_list.append(adm(344169))
  for id in object_list:
        i=id.caseid
        go_Inpatient_Admission_View(browser)
        key_InID(str(i),browser)
        click_PaitientName(browser)
        Click_Calc(browser)
        Download_WS(browser)
        try:
          file=get_Lastest_File(download_path,"xlsm")
        except IndexError:
          continue
        list=get_all_File(download_path,"xlsm")
        for xlsm in list:
          if str(i) in xlsm:
            move_excel_to_print(xlsm)

            #id.Ac_ws=str(i)+" Actual Ws"
            #id.Dr_Fee=str(i)+" Dr Fee(Surg)"
            #id.Dr_Fee_Non=str(i)+" Dr Fee(Non-Surg)"
            #excel_to_pdf(file,sheet_name="Actual Ws",path=print_path+"\\"+id.Ac_ws)
            #excel_to_pdf(file,sheet_name="Dr Fee(Surg)",path=print_path+"\\"+id.Dr_Fee)
            #excel_to_pdf(file,sheet_name="Dr Fee(Non-Surg)",path=print_path+"\\"+id.Dr_Fee_Non)

            file=move_excel_to_print(xlsm)
            id.ws_path=file.replace(download_path,"")
            print(id.ws_path)

            if str(344169) == str(i):
              os.remove(file)
              id.ws_path=""
              continue
          else:
            print("no worksheet for", i)
            continue
  #os.remove(print_path+"\\"+str(344169)+" Actual Ws.pdf")
  #os.remove(print_path+"\\"+str(344169)+" Dr Fee(Surg).pdf")
  #os.remove(print_path+"\\"+str(344169)+" Dr Fee(Non-Surg).pdf")

def move_to_storage(List_file,destination=None,path=None):
  if destination !=None:
    for i in List_file:
      try:
        if os.path.exists(i.replace(path,destination)):
          os.remove(i.replace(path,destination))
        # Move a file by renaming it's path
        os.rename(i, i.replace(path,destination))
      except:
        continue

def move_excel_to_print(file):
  if os.path.exists(file.replace("New","Print",1)):
    os.remove(file.replace("New","Print",1))
        # Move a file by renaming it's path
  os.rename(file, file.replace("New","Print",1))
  return file.replace("New","Print",1)


def pdf_move_to_storage(List_file,destination=None,path=None):
  if destination !=None:
    for i in List_file:
      try:
        if os.path.exists(i.replace(path,destination)):
          os.remove(i.replace(path,destination))
        # Move a file by renaming it's path
        os.rename(i, i.replace(path,destination))
      except:
        continue

def move_to_bord_listing(BDFile,bord_lisiting_path,first_adm_object,base):
  client_name=first_adm_object.client_name
  policy_type=first_adm_object.policy_type
  Year=datetime.now().year
  invoice_type=first_adm_object.invoice_type

  if os.path.exists(bord_lisiting_path) != True:
    createfolder(bord_lisiting_path,base)

  client_name_path=bord_lisiting_path+"\\{}".format(client_name)
  if os.path.exists(client_name_path) != True:
    createfolder(client_name_path,base)

  policy_type_path=client_name_path+"\\{}".format(policy_type)
  if os.path.exists(policy_type_path) != True:
    createfolder(policy_type_path,base)

  Year_path=policy_type_path+"\\{}".format(Year)
  if os.path.exists(Year_path) != True:
    createfolder(Year_path,base)

  invoice_type_path=Year_path+"\\{}".format(invoice_type)
  if os.path.exists(invoice_type_path) != True:
    createfolder(invoice_type_path,base)

  new_bord_listing_path=invoice_type_path+"\\{}".format(BDFile.split("\\")[-1])
  copyfile(BDFile,new_bord_listing_path)
