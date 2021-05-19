import pandas as pd
import math
import numpy
from xlrd import *
from xlwt import *
import datetime
from transformation.medi_bdx_manipulation import find_total_index,count_BdxL
from transformation.medi_crm_bdx_manipulation import find_total_index as find_total_index_crm
from transformation.medi_crm_bdx_manipulation import count_BdxL as count_BdxL_crm

from xlutils.copy import copy
from utils.audit_trail import audit_log

def build_rb(file,formatting_info=False):
  if formatting_info==True:
    rb  = open_workbook(file,formatting_info=True) # where vis.xls is your test file
  else:
    rb  = open_workbook(file)
  return rb

def build_wb(file,formatting_info=False):
  if formatting_info==True:
    rb  = open_workbook(file,formatting_info=True) # where vis.xls is your test file
  else:
    rb  = open_workbook(file)
  wb=copy(rb)
  return wb

def savefile(wb,path):
  wb.save(path)

def get_dcm_fill_index(dcmfile):
  data=pd.read_excel(dcmfile,skiprows = 2 , na_values = "Missing")
  Bord_No_list = pd.DataFrame(data, columns=['Bord No']).values.tolist()
  counter=len(Bord_No_list)-1
  try:
    while True:
      math.isnan(Bord_No_list[counter][0])
      counter-=1
  except:
    a=None
  fill_index=counter+1
  return fill_index

def get_latest_running_no(dcmfile):
  data=pd.read_excel(dcmfile,skiprows = 2 , na_values = "Missing")
  return (data.iloc[get_dcm_fill_index(dcmfile),1])

def get_latest_index(dcmfile):
  data=pd.read_excel(dcmfile,skiprows = 2 , na_values = "Missing")
  return get_dcm_fill_index(dcmfile)+3

def update_debit_note_No(BDFile,value):
  rb=build_rb(BDFile,formatting_info=True)
  wb=build_wb(BDFile,formatting_info=True)
  readerSheet = rb.sheet_by_index(0)
  write_sheet = wb.get_sheet(0)
  row=0
  for i in range(20):
    print(readerSheet.cell(i,0).value)
    if "debit" in readerSheet.cell(i,0).value.lower():
      row=i
      break
  style3 = XFStyle()
  font3 = Font()
  Alignment3=Alignment()
  #Alignment3.horz=Alignment.HORZ_CENTER
  Alignment3.vert=Alignment.VERT_CENTER
  style3.alignment.wrap = 5
  style3.alignment=Alignment3
  borders3= Borders()
  borders3.left=Borders.THIN
  borders3.right=Borders.THIN
  borders3.top=Borders.THIN
  borders3.bottom=Borders.THIN
  borders3.left_colour = Style.colour_map["gray25"]
  borders3.right_colour =  Style.colour_map["gray25"]
  borders3.top_colour =  Style.colour_map["gray25"]
  borders3.bottom_colour =  Style.colour_map["gray25"]
  style3.borders=borders3
  font3.name = 'Arial'
  font3.height = 20*10
  font3.bold=True
  style3.font = font3
  pattern3 = Pattern()
  pattern3.pattern = Pattern.SOLID_PATTERN
  pattern3.pattern_fore_colour = Style.colour_map['white']
  style3.pattern = pattern3

  write_sheet.write(row,3,value,style3)
  savefile(wb,BDFile)


def update_to_dcm(BDFile,dcm_path,ocm=False):
  dcmfile=dcm_path
  path=dcmfile
  bdx_type=""
  if "WEB" in BDFile:
    bdx_type="web"
  elif "MANUAL" in BDFile:
    bdx_type="manual"
  if ocm==True:
    bdx_type=""
  rb=open_workbook(BDFile,formatting_info=True)
  readerSheet = rb.sheet_by_index(0)
  wb=build_wb(dcmfile,formatting_info=True)
  write_sheet = wb.get_sheet(0)

  if bdx_type.lower()=="manual":
    print("manual")
    bord_date=readerSheet.cell(7,3).value
    print(bord_date)
    #bord_date = datetime.datetime(*xldate_as_tuple(bord_date, rb.datemode))
    #bord_date = bord_date.strftime("%d/%m/%y").replace("0","")
    file_no=BDFile.split("\\")[-1]
    file_no=file_no.replace("MANUAL.xls","")
    bord_number=readerSheet.cell(6,3).value
    client=readerSheet.cell(3,3).value
    row,column=find_total_index(BDFile,bdx_type)
    ammount=readerSheet.cell(row,column).value
    reason="O/P MEDICAL CLAIMS"
    name_row=""
    for name_row in range(count_BdxL(BDFile)+40):
      if readerSheet.cell(name_row,column-1).value=="Name:":
        break
    try:
      initial=readerSheet.cell(name_row,column).value.split()[-1]
    except:
      initial=""
      pass
    category="OP"
    total_case=count_BdxL(BDFile)
    ammount=readerSheet.cell(row,column).value
  elif bdx_type.lower()=="web":
    print("web")
    bord_date=readerSheet.cell(8,3).value
    #bord_date = datetime.datetime(*xldate_as_tuple(bord_date, rb.datemode))
    #bord_date = bord_date.strftime("%d/%m/%y").replace("0","")
    file_no=BDFile.split("\\")[-1]
    file_no=file_no.replace("WEB.xls","")
    bord_number=readerSheet.cell(7,3).value
    client=readerSheet.cell(3,3).value
    row,column=find_total_index(BDFile,bdx_type)
    ammount=readerSheet.cell(row,column).value
    reason="O/P MEDICAL CLAIMS"
    name_row=""
    for name_row in range(count_BdxL(BDFile)+40):
      if readerSheet.cell(name_row,column-1).value=="Name:":
        break
    print(readerSheet.cell(name_row,column).value)
    try:
      initial=readerSheet.cell(name_row,column).value.split()[-1]
    except:
      initial=""
      pass
    category="OP"
    total_case=str(count_BdxL(BDFile))
  if ocm==True:
    for i in range(11):
      if "Submission Date"  in readerSheet.cell(i,0).value:
        bord_date= readerSheet.cell(i,2).value
        break
    file_no=BDFile.split("\\")[-1]
    file_no=file_no.replace("(W).xls","")
    file_no=file_no.replace("(M).xls","")

    for i in range(11):
      if "Borderaux"  in readerSheet.cell(i,0).value:
        bord_number= readerSheet.cell(i,2).value
        break

    for i in range(11):
      if "Corporate"  in readerSheet.cell(i,0).value:
        client= readerSheet.cell(i,2).value
        break
    reason="O/P MEDICAL CLAIMS"
    row,column=find_total_index_crm(BDFile)
    ammount=readerSheet.cell(row,column).value
    initial=readerSheet.cell(row+8,column).value
    category="OP"
    total_case=count_BdxL_crm(BDFile)



  style = XFStyle() 
  font = Font()
  font.name = 'Calibri'
  alignment=Alignment()
  alignment.horz=Alignment.HORZ_CENTER
  alignment.vert=Alignment.VERT_CENTER
  style.alignment.wrap = 5
  style.alignment=alignment
  font.height = 20*9
  borders= Borders()
  borders.left=Borders.THIN
  borders.right=Borders.THIN
  borders.top=Borders.THIN
  borders.bottom=Borders.THIN
  style.borders=borders
  font.colour_index=Style.colour_map['red']
  style.font = font
  pattern = Pattern()
  pattern.pattern = Pattern.SOLID_PATTERN
  pattern.pattern_fore_colour = Style.colour_map['white']
  style.pattern = pattern
  try:
    if client=="Self-Funded":
      client=readerSheet.cell(4,3).value
  except:
    pass

  row=get_latest_index(dcmfile)
  rb2=build_rb(dcmfile)
  readerSheet=rb2.sheet_by_index(0)
  running_no=readerSheet.cell(row,1).value
  write_sheet.write(row,0,bord_date,style)
  write_sheet.write(row,2,file_no,style)
  write_sheet.write(row,3,bord_number,style)
  write_sheet.write(row,4,client,style)
  write_sheet.write(row,6,ammount,style)
  write_sheet.write(row,7,reason,style)
  write_sheet.write(row,8,initial,style)
  write_sheet.write(row,9,category,style)
  write_sheet.write(row,17,total_case,style)
  savefile(wb,path)
  update_debit_note_No(BDFile,running_no)
  return bord_number,ammount,bord_date,running_no,bdx_type


#path=r"C:\Users\junshou.leow\Desktop\mediclinic\Disbursement Claims Running No 2019(edit).xls"
#BDFile=r"C:\Users\junshou.leow\Desktop\mediclinic\Dev\08 AUGUST\ONEMILL-0002-201908 WEB.xls"
#update_to_dcm(BDFile,path)


#file_path=r"C:\Users\junshou.leow\Desktop\ss MANUAL.xls"
#dcm_path=r"C:\Users\junshou.leow\Desktop\Client\Disbursement Claims Running No 2019.xls"
#bord_number,ammount,bord_date,running_no,bdx_type=update_to_dcm(file_path,dcm_path,ocm=True)
