#!/usr/bin/python
# FINAL SCRIPT updated as of 5th June 2020
# Workflow - CBA/ADMISSION
# Version 1
# adm_bdx_manipulation.py

import pandas as pd
import numpy as np
from xlutils.copy import copy #http://pypi.python.org/pypi/xlutils
from xlrd import *
from xlwt import *
import pandas as pd
from PIL import Image
from io import BytesIO
from automagic.marc import *
from transformation.adm_connect_db import *
import win32com.client as win32
import glob
import os
import json
import win32api
import pythoncom
import traceback
from utils.audit_trail import audit_log
#from automagica import Wait
from connector.connector import MySqlConnector, MSSqlConnector
from loading.query.query_as_df import *
from pathlib import Path

class fileo:
  def __init__(self):
    self._Raw_Path=""
    self._Ori_Path=""
    self._Save_Path=""

  def set_Save_Path(self,path):
    self._Save_Path=path
  def get_Save_Path(self):
    return self._Save_Path

  def set_Raw_Path(self,path):
    self._Raw_Path=path
  def get_Raw_Path(self):
    return self._Raw_Path

  def set_Ori_Path(self,path):
    self._Ori_Path=path
  def get_Ori_Path(self):
    return self._Ori_Path


def build_rb(file, formatting_info = False):
  """
  use xlrd library to open xls excel file
  return readable workbook
  """
  if formatting_info == True:
    rb  = open_workbook(file, formatting_info = True) # where vis.xls is your test file
  else:
    rb  = open_workbook(file)
  return rb


def build_wb(file,formatting_info=False):
  """
  use xlrd library to open xls excel file
  return writable workbook
  """
  if formatting_info==True:
    rb  = open_workbook(file,formatting_info=True) # where vis.xls is your test file
  else:
    rb  = open_workbook(file)
  wb=copy(rb)
  return wb

def savefile(wb,path):
  """
  save writable workbook to the path
  """
  wb.save(path)


def reinsert_image(wb):
    """
    the image path is hard coded
    this function reinsert image back to the bord listing file because library do not support image 
    """

    current_path = os.path.dirname(os.path.abspath(__file__))
    #current_path = os.path.dirname(os.path.realpath(__file__))
    image_path=current_path+"\\aa_logo.jpg"
    img = Image.open(image_path)
    image_parts = img.split()
    r = image_parts[0]
    g = image_parts[1]
    b = image_parts[2]
    img = Image.merge("RGB", (r, g, b))
    fo = BytesIO()
    img.save(fo, format='bmp')
    w_sheet = wb.get_sheet(0)
    w_sheet.insert_bitmap_data(fo.getvalue(),0,0)
    w_sheet = wb.get_sheet(1)
    w_sheet.insert_bitmap_data(fo.getvalue(),0,0)
    img.close()


def check_Pol_Type(policy_type):
  Policy_type=policy_type
  if "GROUP" in Policy_type.upper():
    Pol_Type_short="GRP"
  elif "INDIVIDUAL" in Policy_type.upper():
    Pol_Type_short="IND"
  return Pol_Type_short,Policy_type.upper()


def get_specific_insurer_detail_new(adm_obj, adm_base):

  caseid = adm_obj[0].caseid
  Insurance = adm_obj[0].client_name
  Insurance_name = Insurance.split(" ")
  
  # Query MSSQL for Insurer's Address
  #connector = MSSqlConnector
  #query = '''SELECT * FROM cba.insurer_details'''
  #database = mssql_get_df_by_query_without_param(query, adm_base, connector)
  database = pd.read_excel(r'\\dtisvr2\CBA_UAT\Insurer\Cleaned ALL Client Address May 2020.xlsx')
  database = pd.DataFrame(database.drop_duplicates().reset_index().drop(columns = ['index']))
  database['MAILING ADDRESS'].replace('', np.nan, inplace = True)
  database.dropna(subset = ['MAILING ADDRESS'], inplace = True)

  try:
    k = 0
    for k in range(len(Insurance_name)):
      searched = database[database['COMPANY NAME'].str.lower().astype(str).str.contains(Insurance_name[k].lower())]
      if searched.empty != True and searched['COMPANY NAME'].count() > 1:
        database = searched.drop_duplicates().reset_index().drop(columns = ['index'])
      elif searched.empty != True and searched['COMPANY NAME'].count() == 1:
        final = searched.drop_duplicates().reset_index().drop(columns = ['index'])
        final = final[['COMPANY NAME', 'INSURER', 'ATTEN #1', 'MAILING ADDRESS', 'ID']].values.astype(str).tolist()
        audit_log("Admission - Insurance details successfully found.", "Completed...", adm_base)
        break;
      else:
        print("teda")
        database = database
        pass
      continue

    if searched.empty != True and searched['COMPANY NAME'].count() > 1:
      audit_log("Admission - Too many insurance details found. Unable to specify correctly. Please check.", "Completed...", adm_base)
    elif searched.empty != True and searched['COMPANY NAME'].count() < 1:
      audit_log("Admission - Insurance details cannot be found.", "Completed...", adm_base)

    return final

  except Exception as e:
    audit_log("Admission - Searching for insurence details failed. Please check.", "Completed...", adm_base)



def get_specific_insurer_detail(adm_obj, adm_base):

  #caseid = '344169'
  #Insurance = 'AXA Affin General Insurance Berhad'

  caseid = adm_obj[0].caseid
  Insurance = adm_obj[0].client_name
  Insurance_name = Insurance.split(" ")
  database = pd.DataFrame(get_insurer_info())
  extact_paring = pd.DataFrame()
  final = pd.DataFrame()

  try:
    k = 0
    for k in range(len(database)):
      if database[0][k] is not None and Insurance.lower() == database[0][k].lower():
        print("%s YES" % k)
        extact_paring = extact_paring.append(pd.DataFrame(database.loc[k:k, 0:5]))
        final = final.append(extact_paring).reset_index().drop(columns = ['index'])
        break
      else:
        print("NO")
        continue

    if final.empty != True and len(final) > 1:
      list_for_all_same_first_name = pd.DataFrame()
      list_for_all_same_second_name = pd.DataFrame()
      i = 0
      for i in range(len(database)):
        if database[0][i] is not None and Insurance_name[0].lower() in database[0][i].lower():
          list_for_all_same_first_name = pd.DataFrame(list_for_all_same_first_name.append(database.loc[i:i, 0:5]))
        elif database[0][i] is not None and Insurance_name[1].lower() in database[0][i].lower():
          list_for_all_same_second_name = pd.DataFrame(list_for_all_same_second_name.append((database.loc[i:i, 0:5])))

      final = pd.DataFrame(list_for_all_same_first_name.append(list_for_all_same_second_name))

      if len(list_for_all_same_second_name) > 1:
        list_for_all_same_second_name.columns = ['comp_name', 'insurer', 'recepient', 'address', 'client_id', 'type']
        list_for_all_same_second_name = list_for_all_same_second_name.reset_index().drop(columns = ['index'])
        Pol_Type, policy_type = check_Pol_Type(adm_obj[0].policy_type)
        g = 0
        for g in range(len(list_for_all_same_second_name)):
          if list_for_all_same_second_name['type'][g] is not None and list_for_all_same_second_name['type'][g] == Pol_Type:
            final.append(list_for_all_same_second_name[g])
            break
          else:
            pass

          if len(final) is not 1:
            final.append(list_for_all_same_second_name[0])
      else:
        final = get_specific_insurer_detail_new(adm_obj, adm_base)
    elif final.empty != True and len(final) == 1:
      final = final.values.astype(str).tolist()
      print("HERE")
    else:
      final = get_specific_insurer_detail_new(adm_obj, adm_base)

    audit_log("Insurance detail found successfuly", "Completed...", adm_base)
  except Exception:
    try:
      final = get_specific_insurer_detail_new(adm_obj, adm_base)
      if pd.DataFrame(final).empty == True:
        final = [(str(Insurance), str(Insurance_name[0]), 'N/A', 'UNABLE TO BE IDENTIFIED', 0, None)]
      else:
        final = final
        final.columns = ['insurer', 'comp_name', 'recepient', 'address', 'client_id', 'type']
    except Exception as e:
      audit_log("Insurance detials failed to find ".format(e), "Completed...", adm_base)
      final = [(str(Insurance), str(Insurance_name[0]), 'N/A', 'UNABLE TO BE IDENTIFIED', 0, None)]
      audit_log("Insurance detail Not found ,Check to LONPAC automatically", "Completed...", adm_base)
  return final


def Get_Address(adm_obj, adm_base):
  att = None
  add = None
  Cliend_id = None
  enlisted = None
  final = get_specific_insurer_detail(adm_obj, adm_base)
  test_final = pd.DataFrame(final)
  if test_final.empty != True:
    try:
      formatted_Address = split_address(final['comp_name'][0], final['address'][0])
      PREVIEW_INFO = [final['recepient'][0], final['address'][0]]
      return PREVIEW_INFO[0], PREVIEW_INFO[1], final['client_id'][0], final
    except Exception:
      formatted_Address = split_address(final[0][0], final[0][3])
      PREVIEW_INFO = [final[0][2], final[0][3]]
      return PREVIEW_INFO[0], PREVIEW_INFO[1], final[0][4], final
  else:
    att = None
    add = None
    Cliend_id = None
    enlisted = None
    return att, add, Cliend_id, enlisted

def read_BdxL(file):
  """
  This function read the excel for all listing available and return all content with headrow dataframe

        :param str file:                   Bordereaux file location
        :rtype:                            Dataframe
  """
  rb = build_rb(file)
  readerSheet = rb.sheet_by_index(0)
  skiprow=0
  while "No" != readerSheet.cell(skiprow,0).value:
    skiprow+=1
  data=pd.read_excel(file,skiprows = skiprow , na_values = "Missing")
  data1=(data.loc[data['No'].isin([i for i in range(1000)])]).copy()
  return data1

def count_BdxL(file):
  """
  This function read the excel and return the bordreaux content row amount

        :param str file:                   Bordereaux file location
        :rtype:                            int
  """
  rb=build_rb(file)
  readerSheet = rb.sheet_by_index(0)
  skiprow=0
  while "No" != readerSheet.cell(skiprow,0).value:
    skiprow+=1
  data=pd.read_excel(file,skiprows = skiprow , na_values = "Missing")
  data1=(data.loc[data['No'].isin([i for i in range(1000)])]).copy()
  return len(data1.index)

def count_BdxL_Colomn(file):
  """
  This function read the excel and return the bordreaux content column amount

        :param str file:                   Bordereaux file location
        :rtype:                            int
  """
  rb=build_rb(file)
  readerSheet = rb.sheet_by_index(0)
  skiprow=0
  while "No" != readerSheet.cell(skiprow,0).value:
    skiprow+=1
  data=pd.read_excel(file,skiprows = skiprow , na_values = "Missing")  
  data1=(data.loc[data['No'].isin([i for i in range(1000)])]).copy().T
  return len(data1.index)

def autoAdjustColumns(workbook,path ,writerSheet, writerSheet_index, extraCushion,bdx=True):
    """
      this function adjust the excel column for disburment listing automatically
        :param xlwt.workbook workbook:                   workbook that going to write
        :param str file:                                 Bordereaux file location
        :param xlwt.wordsheet writerSheet:               workbook that which activated existing worksheet for edit
        :param int writerSheet_index:                    index number for writersheet to copy
        :param int extraCushion:                         set addtional range for bordreaux width

        :rtype:                            int
    """
    if bdx:
      starting_point=15
      ending_point=17+count_BdxL(path)
    else:
      starting_point=0
      ending_point=100
    no_of_col=count_BdxL_Colomn(path)
    readerSheet = open_workbook(path).sheet_by_index(writerSheet_index)
    for col in range(no_of_col):
          biggest=0
          for row in range(starting_point,ending_point):
            thisCell = readerSheet.cell(row, col)
            neededWidth = int((1 + len(str(thisCell.value))) * 256)
            if neededWidth>=biggest:
              biggest=neededWidth
          writerSheet.col(col).width = biggest+extraCushion


def get_Lastest_File(file_path, file_type):
  """
  get latest file by the lastest time it download
  """
  if file_type == "xls":
    list_of_files = glob.glob(file_path + "\*.xls") # * means all if need specific format then *.csv
  elif file_type == "xlsm":
    list_of_files = glob.glob(file_path+"\*.xlsm") # * means all if need specific format then *.csv
  latest_file = sorted(list_of_files, key = os.path.getctime)[-1]
  return latest_file

def get_all_File(file_path,file_type):
  """
  get all file from a floder
  """
  if file_type=="xls":
    list_of_files = glob.glob(file_path+"\*.xls") # * means all if need specific format then *.csv
  elif file_type=="pdf":
    list_of_files = glob.glob(file_path+"\*.pdf") # * means all if need specific format then *.csv
  elif file_type=="xlsm":
    list_of_files = glob.glob(file_path+"\*.xlsm") # * means all if need specific format then *.csv
  return list_of_files

def get_Second_Lastest_File(file_path,file_type):
  """
  get latest file by the second lastest time it download
  """
  if file_type=="xls":
    list_of_files = glob.glob(file_path+"\*.xls") # * means all if need specific format then *.csv
  elif file_type=="xlsm":
    list_of_files = glob.glob(file_path+"\*.xlsm") # * means all if need specific format then *.csv
  second_latest_file = sorted(list_of_files, key=os.path.getctime)[-2]
  return second_latest_file



def col_to_num(col_str):
    """
    Convert base26 column string to number
    """
    expn = 0
    col_num = 0
    for char in reversed(col_str):
        col_num += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1
    col_index=col_num-1
    return col_index


def copy_Coloum(coloum_index,row_ammount,copy_coloum_index,read_sheet,write_sheet):
  style = XFStyle()
  font = Font()
  font.name = 'Calibri'
  font.height = 20*12
  borders= Borders()
  borders.left=Borders.THIN
  borders.right=Borders.THIN
  borders.top=Borders.THIN
  borders.bottom=Borders.THIN
  style.borders=borders
  font.colour_index=Style.colour_map['black']
  style.font = font
  pattern = Pattern()
  pattern.pattern = Pattern.SOLID_PATTERN
  pattern.pattern_fore_colour = Style.colour_map['white']
  style.pattern = pattern

  style1 = XFStyle()
  font1 = Font()
  font1.name = 'Calibri'
  font1.height = 20*12
  borders1= Borders()
  borders1.left=Borders.THIN
  borders1.right=Borders.THIN
  borders1.top=Borders.THIN
  borders1.bottom=Borders.THIN
  style1.borders=borders1
  font1.colour_index=Style.colour_map['white']
  style1.font = font1
  pattern1 = Pattern()
  pattern1.pattern = Pattern.SOLID_PATTERN
  pattern1.pattern_fore_colour = Style.colour_map['black']
  style1.pattern = pattern1

  style2 = XFStyle()
  style2=Style.default_style

  style3 = XFStyle()
  font3 = Font()
  font3.name = 'Calibri'
  font3.height = 20*12
  font3.bold=True
  style3.font = font3
  pattern3 = Pattern()
  pattern3.pattern = Pattern.SOLID_PATTERN
  pattern3.pattern_fore_colour = Style.colour_map['white']
  style3.pattern = pattern3

  style4 = XFStyle()
  font4 = Font()
  Borders4=Borders()
  Borders4.bottom=Borders.DOUBLE
  font4.name = 'Calibri'
  font4.height = 20*12
  font4.bold=True
  style4.font = font4
  style4.borders=Borders4
  pattern4 = Pattern()
  pattern4.pattern = Pattern.SOLID_PATTERN
  pattern4.pattern_fore_colour = Style.colour_map['white']
  style4.pattern = pattern4
  skiprow=0
  while "No" != read_sheet.cell(skiprow,0).value:
    skiprow+=1
  for i in range(skiprow,row_ammount+17+8):
    if i==skiprow:
      write_sheet.write(i,coloum_index,read_sheet.cell(i,copy_coloum_index).value,style1)
      continue
    if i>row_ammount+skiprow:
      if i==row_ammount+skiprow+5 or i==row_ammount+skiprow+8:
        try:
          temp=int(read_sheet.cell(i,copy_coloum_index).value)
          write_sheet.write(i,coloum_index,read_sheet.cell(i,copy_coloum_index).value,style4)
          continue
        except ValueError:
          write_sheet.write(i,coloum_index,read_sheet.cell(i,copy_coloum_index).value,style3)
          continue
      write_sheet.write(i,coloum_index,read_sheet.cell(i,copy_coloum_index).value,style2)
      continue
    style.font = font
    write_sheet.write(i,coloum_index,read_sheet.cell(i,copy_coloum_index).value,style)

def delete_Coloum(coloum_index,row_ammount,write_sheet,skiprow):
  """
  delete REMAINING column at a certain CELL which is table of bordereaux
  """
  style = XFStyle()
  style=Style.default_style
  for i in range(skiprow,row_ammount+skiprow+1):
    write_sheet.write(i,coloum_index,'',style)


def delete_bdx_Coloum(string, rb, wb, file):
  """
      this function is to delete column and copy the next index to current one
        :param str string:                   workbook that going to write
        :param xlrd.workbook rb:             Bordereaux file location
        :param xlwt.wordbook wb:             workbook that which activated existing worksheet for edit
        :param str file:                     bordreaux file location
  """
  r_sheet = rb.sheet_by_index(0)
  skiprow = 0

  while "No" != r_sheet.cell(skiprow,0).value:
    skiprow+=1

  w_sheet = wb.get_sheet(0)
  delete = string.upper()
  delete_char = delete.split(",")
  delete_Index_List = []
  for i in range(len(delete_char)):
    col_to_num(delete_char[i])
    delete_Index_List.append(col_to_num(delete_char[i]))

  # Get the list of column indexes
  #allcolumns = []
  #allcolumns = [*range(count_BdxL_Colomn(file))]
  #delete_these_actually = list(set(allcolumns) - set(delete_Index_List))

  minimun_index = min(delete_Index_List)
  current_Index = 0 #index that walk to next coloum to keep
  for i in range(0, count_BdxL_Colomn(file)):
    if i in delete_Index_List:
      print("yes %s" %i)
      continue
    else:
      copy_Coloum(current_Index, count_BdxL(file), i, r_sheet, w_sheet)
    current_Index+=1

  for i in range(len(delete_Index_List)):
  #i = 1
  #for i in range(len(delete_these_actually)):
    delete_Coloum(current_Index,count_BdxL(file)+8,w_sheet,skiprow)
    current_Index+=1

def b2b_Check_Bdx_Total_Amount(file):
  """
  check the total amount is matach or not
  """ 
  df=read_BdxL(file).T
  total=0
  for i in df.loc["B2B"]:
    total+=i
  for i in df.loc["Insurance"]:
    total+=i
  output=round(total,2)
  column=0
  for i in (df.index):
    if i == "Insurance":
      break
    column+=1
  row=17+count_BdxL(file)+8
  readerSheet = open_workbook(file).sheet_by_index(0)
  thisCell = readerSheet.cell(row-1, column).value
  if thisCell==output:
    return True
  else:
    return False

def Check_Bdx_Total_Amount(file,rb,bdx_type):
  """
  pick the type to check
  """ 
  if "CASHLESS" in bdx_type:
    return cashless_Check_Bdx_Total_Amount(file)
  if "B2B" in bdx_type:
    return b2b_Check_Bdx_Total_Amount(file)

def get_Amount_Due(file):
  df=read_BdxL(file).T
  column=0
  for i in (df.index):
    if i == "Insurance":
      break
    column+=1
  row=17+count_BdxL(file)+8
  readerSheet = open_workbook(file).sheet_by_index(0)
  Amount_Due = readerSheet.cell(row-1, column).value
  return Amount_Due


def split_address(company_name,address):
  address=address.split(',')
  formatted_address=[]
  formatted_address.append(company_name)
  Next=False
  for i in range(len(address)):
    if Next:
      Next=False
      continue
    if len(address[i])<15 and  len(address)!=i+1:
      formatted_address.append(address[i]+","+address[i+1])
      Next=True
      continue
    formatted_address.append(address[i])
  String=""
  for i in range(len(formatted_address)):
    String=String+formatted_address[i]
    if len(formatted_address)!=i+1:
      String=String+"\n"
  return String

def cashless_Check_Bdx_Total_Amount(file):
  """
  check the total amount is matach or not
  """ 
  df=read_BdxL(file).T
  total=0
  for i in df.loc["Insurance"]:
    total+=i
  output=round(total,2)
  column=0
  for i in (df.index):
    if i == "Insurance":
      break
    column+=1
  row=17+count_BdxL(file)+8
  readerSheet = open_workbook(file).sheet_by_index(0)
  thisCell = readerSheet.cell(row-1, column).value
  if thisCell==output:
    return True
  else:
    return False


def acknowledgement(rb, wb, file, enlisted, adm_base):
  r_sheet = rb.sheet_by_index(0)
  Insurance = r_sheet.cell(10, 1).value.upper()
  Bord_No = r_sheet.cell(12, 1).value
  Bord_No = Bord_No.split("/")
  Bord_No[0] = Bord_No[0].zfill(5)
  Bord_No = Bord_No[0] + "/" + Bord_No[1].replace('(',' (')
  Insurance_name = Insurance.split(" ")

  try:
    First_Case_Id = str(int(r_sheet.cell(17, 1).value))
  except:
    First_Case_Id = str(r_sheet.cell(17, 1).value)

  if len(First_Case_Id) > 6:
    Case_ID_List = First_Case_Id.split("-")
    First_Case_Id = Case_ID_List[1]
  else:
    pass

  Amount_Due = get_Amount_Due(file)
  r_sheet = rb.sheet_by_index(1)
  Grand_total = r_sheet.cell(22, 3).value
  enlisted = enlisted
  #differentiate insurance policy type to get a unique and correct insurance information
  formatted_Address = split_address(enlisted[0][0], enlisted[0][3])
  height = formatted_Address.count('\n') + 1

  w_sheet = wb.get_sheet(1)
  style = XFStyle()
  style.alignment.wrap = 5
  font = Font()
  font.name = 'Calibri'
  font.height = 20*12
  style.font = font
  w_sheet.write(8, 1, formatted_Address, style)
  w_sheet.row(8).height_mismatch = True
  w_sheet.row(8).height = (40 + 256) * height

  style2 = XFStyle() #for row 15
  style2.alignment.wrap = 5
  font2 = Font()
  font2.name = 'Calibri'
  font2.height = 20*12
  style2.font = font2
  borders2 = Borders()
  borders2.left = Borders.THIN
  borders2.right = Borders.THIN
  borders2.top = Borders.THIN
  borders2.bottom = Borders.THIN
  style2.borders = borders2

  w_sheet.write(10, 1, enlisted[0][2], style) # 'NEO YIN YIN (CLAIMS DEPT)'
  w_sheet.write(14, 2, Bord_No, style2) # '00001/20 (P)'
  #reinsert_image(wb)
  Cliend_Id = enlisted[0][4] # '48'

  r_sheet = rb.sheet_by_index(1)
  w_sheet.write(14, 2, Bord_No, style2) # '00001/20 (P)'
  w_sheet.write(14, 0, r_sheet.cell(14,0).value, style2) # 'REFER ITEMISED'
  w_sheet.write(14, 1, r_sheet.cell(14,1).value, style2) # 'H&S20-00043'
  w_sheet.write(14, 3, r_sheet.cell(14,3).value, style2) # '04/06/2020'

  style1 = XFStyle() # index 13
  font1 = Font()
  font1.name = 'Calibri'
  font1.height = 20*12
  borders1 = Borders()
  borders1.left = Borders.THIN
  borders1.right = Borders.THIN
  borders1.top = Borders.THIN
  borders1.bottom = Borders.THIN
  style1.borders = borders1
  font1.colour_index = Style.colour_map['white']
  style1.font = font1
  pattern1 = Pattern()
  Alignment1 = Alignment()
  Alignment1.horz = Alignment.HORZ_CENTER
  Alignment1.vert = Alignment.VERT_CENTER
  style1.alignment = Alignment1
  pattern1.pattern = Pattern.SOLID_PATTERN
  pattern1.pattern_fore_colour = Style.colour_map['black']
  style1.pattern = pattern1
  w_sheet.write(13, 0, r_sheet.cell(13,0).value, style1) # 'File No'
  w_sheet.write(13, 1, r_sheet.cell(13,1).value, style1) # 'Disbursement Claims No'
  w_sheet.write(13, 2, r_sheet.cell(13,2).value, style1) # 'Disbursement Listing'
  w_sheet.write(13, 3, r_sheet.cell(13,3).value, style1) # 'Date'

  w_sheet.write(16, 3, r_sheet.cell(16,3).value, style1) # 'Amount'
  w_sheet.write(16, 0, r_sheet.cell(16,0).value, style1) # 'Particulars'

  style3 = XFStyle()
  style3.alignment.wrap = 5
  font3 = Font()
  font3.bold = True
  font3.name = 'Calibri'
  font3.height = 20*12
  style3.font = font3
  borders3 = Borders()
  borders3.left = Borders.NO_LINE
  borders3.right = Borders.NO_LINE
  borders3.top = Borders.NO_LINE
  borders3.bottom = Borders.NO_LINE
  style3.borders = borders3

  w_sheet.col(0).width = 5911
  w_sheet.col(1).width = 6423
  w_sheet.col(2).width = 5515
  w_sheet.col(3).width = 10007
  # fix at the first place, the column adjustment function will auto fix the size again
  w_sheet.write(6, 0, r_sheet.cell(6,0).value, style3) # 'DISBURSEMENT CLAIMS'
  w_sheet.write(8, 0, r_sheet.cell(8,0).value, style3) # 'Bill To : '
  w_sheet.write(10, 0, r_sheet.cell(10,0).value, style3) # 'ATTN : '
  w_sheet.write(28, 0, r_sheet.cell(28,0).value, style3) # 'ASIA ASSISTANCE NETWORK(M) SDN BHD'
  w_sheet.write(37, 0, r_sheet.cell(37,0).value, style3) # 'ASIA ASSISTANCE NETWORK(M) SDN BHD'

  style_normal = XFStyle()
  style_normal.alignment.wrap = 5
  font_normal = Font()
  font_normal.name = 'Calibri'
  font_normal.height = 20*12
  style_normal.font = font_normal
  borders_normal= Borders()
  borders_normal.left = Borders.NO_LINE
  borders_normal.right = Borders.NO_LINE
  borders_normal.top = Borders.NO_LINE
  borders_normal.bottom = Borders.NO_LINE
  style_normal.borders = borders_normal
  w_sheet.write(25, 0, r_sheet.cell(25,0).value, style_normal) # 'Payment Term: 14 days upon disbursement claims submitted'
  w_sheet.write(27, 0, r_sheet.cell(27,0).value, style_normal) # 'All Cheque must be crossed and made payable to:'
  w_sheet.write(30, 0, r_sheet.cell(30,0).value, style_normal) # 'Yours truly,'
  w_sheet.write(36, 0, r_sheet.cell(36,0).value, style_normal) # 'HO'

  style_17_21 = XFStyle()
  style_17_21.alignment.wrap = 5
  font_style_17_21 = Font()
  font_style_17_21.name = 'Calibri'
  font_style_17_21.height = 20*12
  style_17_21.font = font_normal
  borders_style_17_21 = Borders()
  borders_style_17_21.left = Borders.THIN
  borders_style_17_21.right = Borders.THIN
  style_17_21.borders = borders_style_17_21

  for r in range(17, 22):
    for c in range(0, 4):
      w_sheet.write(r, c, r_sheet.cell(r,c).value, style_17_21)

  style_22 = XFStyle()
  style_22.alignment.wrap = 5
  font_style_22 = Font()
  font_style_22.name = 'Calibri'
  font_style_22.height = 20*12
  style_22.font = font_style_22
  borders_style_22 = Borders()
  Alignment_22 = Alignment()
  Alignment_22.horz = Alignment.HORZ_CENTER
  Alignment_22.vert = Alignment.VERT_CENTER
  style_22.alignment = Alignment_22
  borders_style_22.left = Borders.THIN
  borders_style_22.right = Borders.THIN
  borders_style_22.top = Borders.THIN
  borders_style_22.bottom = Borders.THIN
  style_22.borders = borders_style_22
  w_sheet.write(22, 0, r_sheet.cell(22,0).value, style_22) # 'Grand Total:'

  style_22_D = XFStyle()
  style_22_D.alignment.wrap = 5
  font_style_22_D = Font()
  font_style_22_D.name = 'Calibri'
  font_style_22_D.height = 20*12
  font_style_22_D.bold=True
  style_22_D.font = font_style_22_D
  borders_style_22_D = Borders()
  Alignment_22_D = Alignment()
  Alignment_22_D.horz = Alignment.HORZ_CENTER
  Alignment_22_D.vert = Alignment.VERT_CENTER
  style_22_D.alignment = Alignment_22_D
  borders_style_22_D.left = Borders.THIN
  borders_style_22_D.right = Borders.THIN
  borders_style_22_D.top = Borders.THIN
  borders_style_22_D.bottom = Borders.THIN
  style_22_D.borders = borders_style_22_D
  w_sheet.write(22, 3, r_sheet.cell(22,3).value, style_22_D)

  style_30_38 = XFStyle()
  style_30_38.alignment.wrap = 5
  font_style_30_38 = Font()
  font_style_30_38.name = 'Calibri'
  font_style_30_38.height = 20*12
  font_style_30_38.bold = True
  style_30_38.font = font_style_30_38
  borders_style_30_38 = Borders()
  borders_style_30_38.left = Borders.THICK
  borders_style_30_38.right = Borders.THICK
  style_30_38.borders = borders_style_30_38

  style_31 = XFStyle()
  borders_style_31 = Borders()
  borders_style_31.top = Borders.THICK
  borders_style_31.left = Borders.THICK
  borders_style_31.right = Borders.THICK
  style_31.borders = borders_style_31

  style_39 = XFStyle()
  borders_style_39 = Borders()
  borders_style_39.bottom = Borders.THICK
  borders_style_39.left = Borders.THICK
  borders_style_39.right = Borders.THICK
  style_39.borders = borders_style_39
  try:
    for r in range(30,39):
      if r==30:
        w_sheet.write_merge(r,r,2,3,r_sheet.cell(r,2).value,style_31)
        r_sheet.cell(r,2).value
        continue
      if r==38:
        w_sheet.write_merge(r,r,2,3,r_sheet.cell(r,2).value,style_39)
        r_sheet.cell(r,2).value
        continue
      w_sheet.write_merge(r,r,2,3,r_sheet.cell(r,2).value,style_30_38)
      r_sheet.cell(r,2).value
  except:
    pass

  return Cliend_Id

def savefile(wb,path):
  reinsert_image(wb)
  wb.save(path)


def bdx_automation(fileobj, bdx_type, enlisted, adm_base, download_path):

  if "/" in fileobj.get_Ori_Path():
    temp = fileobj.get_Ori_Path()
    fileobj.set_Ori_Path(temp.replace("/"," "))
  else:
    pass
  ori_path = fileobj.get_Ori_Path()
  raw_path = fileobj.get_Raw_Path()

  try:
    build_rb(raw_path)
    build_path = raw_path
  except NotImplementedError:
    fix_excel(raw_path, ori_path, download_path, adm_base)
    audit_log("Fix excel xls format", "Completed...", adm_base)
    build_path=ori_path

  wb = build_wb(build_path)
  rb = build_rb(build_path)
  Cliend_ID = acknowledgement(rb, wb, build_path, enlisted, adm_base)

  # This is where the types of columns to be processed for individual Bord Listing
  Colomn_Action_Type = get_Colomn_Action_Type(Cliend_ID)
  if Colomn_Action_Type == 'DELETE':
    Colomn_Action_String = get_Colomn_Involved(Cliend_ID)
    if bdx_type != 'Cashless':
      delete_bdx_Coloum(Colomn_Action_String, rb, wb, build_path)
    else:
      Colomn_Action_String = Colomn_Action_String.replace('N,', '')
      delete_bdx_Coloum(Colomn_Action_String, rb, wb, build_path)

  w_sheet = wb.get_sheet(0)
  autoAdjustColumns(wb, build_path, w_sheet, 0, 500)
  if fileobj.get_Save_Path()=="":
    fileobj.set_Save_Path(ori_path)
  savefile(wb,fileobj.get_Save_Path())

  file2=fileobj.get_Save_Path()
  wb2=build_wb(file2,formatting_info=True)
  w_sheet2 = wb2.get_sheet(0)
  rb2=build_rb(file2,formatting_info=True)

  autoAdjustColumns(wb2, file2,w_sheet2, 0, 100)
  w_sheet1 = wb2.get_sheet(1) #print in one page for ACK
  w_sheet1.fit_num_pages=1 #print in one page for ACK
  savefile(wb2,file2)
  return Check_Bdx_Total_Amount(file2,rb2,bdx_type),Cliend_ID
