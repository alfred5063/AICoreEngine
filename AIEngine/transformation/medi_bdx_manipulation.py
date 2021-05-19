#!/usr/bin/python
# CLEAN SCRIPT updated as of 11 SEP 2019
# Task 273 - adjust column width
# Area RPA3
# Iteration RPA3\Sprint 8
# Scriptor - Leow Jun Shou
import pandas as pd
from utils.audit_trail import audit_log
import os
from xlutils.copy import copy
from xlrd import *
from xlwt import *
import pandas as pd
from PIL import Image
from io import BytesIO
import win32com.client as win32
from datetime import date,datetime,timedelta
import datetime as dt
from dateutil.relativedelta import relativedelta
from connector.connector import MSSqlConnector
import glob
import pandas as pd
from pathlib import Path
from connector.dbconfig import read_db_config
from connector.connector import MySqlConnector
def get_policy_date_list(first_case_id):
  current_path = os.path.dirname(os.path.abspath(__file__))
  dti_path = str(Path(current_path).parent)
  config = read_db_config(dti_path+r'\config.ini', 'mosdb')
  connector = MySqlConnector(config)
  cursor = connector.cursor()
  qry ="SELECT a. serno, b.mmid, b.mmeff, b.mmexp FROM medicase a, member b WHERE a.serno = '{}' AND b.mmid = a.mmid".format(str(first_case_id))
  cursor.execute(qry)
  records = cursor.fetchall()
  cursor.close()
  connector.close()
  return records
def num_to_date(number):
  day1=number[6]+number[7]
  month1=number[4]+number[5]
  year1=number[0]+number[1]+number[2]+number[3]
  return date(day=int(day1), month=int(month1), year=int(year1)).strftime('%d %B %Y')
def get_first_case_id(file):
  rb=build_rb(file)
  readerSheet = rb.sheet_by_index(0)
  row=0
  column=0
  while "No" != readerSheet.cell(row,0).value:
    row+=1
  while "Serial No" != readerSheet.cell(row,column).value:
    column+=1
  return readerSheet.cell(row+1,column).value
def get_EventDate():
  connector = MSSqlConnector()
  cursor = connector.cursor()
  qry="SELECT date FROM cba.calendar_events"
  cursor = connector.cursor()
  cursor.execute(qry)
  records = cursor.fetchall()
  return records
def get_Date():
  """
  This function get today date and compare with calender event in database
  return a formatted date to bdx 
  """
  Date=dt.date.today()

  return Date

def fill_date():
  Date=get_Date()+timedelta(days=1)
  dateList=[]
  for i in get_EventDate():
    temp=list(i)
    dateList.append(temp[0])
  while Date in dateList:
    Date=Date+timedelta(days=1)
  formatted_Date = Date.strftime("%d/%m/%Y")
  return formatted_Date

def get_skiprows(file):
  rb=build_rb(file)
  readerSheet = rb.sheet_by_index(0)
  skiprow=0
  while "No" != readerSheet.cell(skiprow,0).value:
    skiprow+=1
  return skiprow


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


def read_BdxL(file):
  """
  This function read the excel for all listing available

  Arguement is location of excel file

  End by return a dataframe of the listing
  """
  data=pd.read_excel(file,skiprows = get_skiprows(file) , na_values = "Missing")
  
#code can be improved by flaxible by change the column and row incase marc change formula
  data1=(data.loc[data['No'].isin([i for i in range(1000)])]).copy()
  return data1

def read_BdxL_template(file,row):
  """
  This function read the excel for all listing available

  Arguement is location of excel file

  End by return a dataframe of the listing
  """
  data=pd.read_excel(file,skiprows = row , na_values = "Missing")
  
#code can be improved by flaxible by change the column and row incase marc change formula
  data1=(data.loc[data['No'].isin([i for i in range(1000)])]).copy()
  return data1

def count_BdxL(file):
  """
  This function read the excel for all listing available

  Arguement is location of excel file

  End by return dataframe index number
  """
  data=pd.read_excel(file,skiprows = get_skiprows(file) , na_values = "Missing")
  
#code can be improved by flaxible by change the column and row incase marc change formula

  data1=(data.loc[data['No'].isin([i for i in range(1000)])]).copy()
  row=0
  for i in data['No'].tolist():
    if (type(i)!=type("")):
      break
    row+=1
  return row
def count_BdxL_Colomn(file):
  """
  This function read the excel for all listing available

  Arguement is location of excel file

  End by return dataframe index number
  """
  data=pd.read_excel(file,skiprows = get_skiprows(file) , na_values = "Missing")
#code can be improved by flaxible by change the column and row incase marc change formula
  
  data1=(data.loc[data['No'].isin([i for i in range(1000)])]).copy().T

  return len(data1.index)

def template_skiprows(template):
  print(type(template),"is template")
  if os.path.exists(template) != True:
    print(template)
  rb=build_rb(template)
  readerSheet = rb.sheet_by_index(0)
  for No_row in range(30):
    if "No" == readerSheet.cell(No_row,0).value:
      break
  return No_row

def autoAdjustColumns(path , extraCushion,bdx_type):
    """
    this function adjust the excel column for disburment listing automatically
  """
    rd=build_rb(path,formatting_info=True)
    wb=build_wb(path,formatting_info=True)
    print("adjustting colomn")
    if bdx_type.lower()=="manual":
      starting_point= get_skiprows(path)-1
      ending_point=get_skiprows(path)+1+count_BdxL(path)
      no_of_col=count_BdxL_Colomn(path)
    elif bdx_type.lower()=="web":
      starting_point= get_skiprows(path)-1
      ending_point=get_skiprows(path)+1+count_BdxL(path)
      no_of_col=count_BdxL_Colomn(path)
    print(starting_point)
    print(ending_point)
    readerSheet = rd.sheet_by_index(0)
    writerSheet = wb.get_sheet(0)
    for col in range(no_of_col):
      biggest=0
      for row in range(starting_point,ending_point):
        thisCell = readerSheet.cell(row, col)
        neededWidth = int((1 + len(str(thisCell.value))) * 256)
        if neededWidth>=biggest:
          biggest=neededWidth
      writerSheet.col(col).width = biggest+extraCushion
    savefile(wb,path)



def col_to_num(col_str):
    """ Convert base26 column string to number. """
    expn = 0
    col_num = 0
    for char in reversed(col_str):
        col_num += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1
    col_index=col_num-1

    return col_index




def find_total_index(file,bdx_type):
  read_BdxL(file)
  row=get_skiprows(file)+count_BdxL(file)+3
  column=read_BdxL(file).columns.get_loc("Nett Total (RM)")
  return row,column



def copy_Coloum(coloum_index,row_ammount,copy_coloum_index,read_sheet,write_sheet,file,net_colomn):
  style = XFStyle() 
  font = Font()
  font.name = 'Arial'
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
  font.colour_index=Style.colour_map['black']
  style.font = font
  pattern = Pattern()
  pattern.pattern = Pattern.SOLID_PATTERN
  pattern.pattern_fore_colour = Style.colour_map['white']
  style.pattern = pattern


  style1 = XFStyle() # for eACH labels
  font1 = Font()
  font1.name = 'Arial'
  font1.height = 20*9
  borders1= Borders()
  Alignment1=Alignment()
  Alignment1.horz=Alignment.HORZ_CENTER
  Alignment1.vert=Alignment.VERT_CENTER
  style1.alignment.wrap = 5
  style1.alignment=Alignment1
  borders1.left=Borders.THIN
  borders1.right=Borders.THIN
  borders1.top=Borders.THIN
  borders1.bottom=Borders.THIN
  style1.borders=borders1
  font1.bold=True
  font1.colour_index=Style.colour_map['white']
  style1.font = font1
  pattern1 = Pattern()
  pattern1.pattern = Pattern.SOLID_PATTERN
  pattern1.pattern_fore_colour = Style.colour_map['dark_teal']
  style1.pattern = pattern1

  style2 = XFStyle()
  style2=Style.default_style

  style3 = XFStyle()# for normal everting
  font3 = Font()
  Borders3=Borders()
  Borders3.bottom=Borders.THIN
  Borders3.top=Borders.THIN
  Borders3.left=Borders.THIN
  Borders3.right=Borders.THIN
  Borders3.bottom_colour = Style.colour_map["gray25"]
  Borders3.left_colour = Style.colour_map["gray25"]
  Borders3.right_colour =  Style.colour_map["gray25"]
  Borders3.top_colour =  Style.colour_map["gray25"]
  style3.borders=Borders3
  font3.name = 'Arial'
  font3.height = 20*9
  font3.bold=True
  style3.font = font3
  pattern3 = Pattern()
  pattern3.pattern = Pattern.SOLID_PATTERN
  pattern3.pattern_fore_colour = Style.colour_map['white']
  style3.pattern = pattern3

  style4 = XFStyle()
  font4 = Font() #for total ammount
  Borders4=Borders()
  Alignment4=Alignment()
  Alignment4.horz=Alignment.HORZ_CENTER
  Alignment4.vert=Alignment.VERT_CENTER
  style4.alignment.wrap = 5
  style4.alignment=Alignment1
  Borders4.bottom=Borders.DOUBLE
  Borders4.top=Borders.THIN
  Borders4.left=Borders.THIN
  Borders4.right=Borders.THIN
  Borders4.left_colour = Style.colour_map["gray25"]
  Borders4.right_colour =  Style.colour_map["gray25"]
  Borders4.top_colour =  Style.colour_map["gray25"]
  font4.name = 'Arial'
  font4.height = 20*9
  font4.bold=True
  style4.font = font4
  style4.borders=Borders4
  pattern4 = Pattern()
  pattern4.pattern = Pattern.SOLID_PATTERN
  pattern4.pattern_fore_colour = Style.colour_map['white']
  style4.pattern = pattern4
  

  for i in range(get_skiprows(file),row_ammount+get_skiprows(file)+2):
    if i==get_skiprows(file):
      write_sheet.write(i,coloum_index,read_sheet.cell(i,copy_coloum_index).value,style1)
      if read_sheet.cell(i,copy_coloum_index).value=="Nett Total (RM)":
        net_colomn=coloum_index
      continue

    if i>row_ammount+get_skiprows(file):
      
      if i==row_ammount+get_skiprows(file)+1:
        try:
          temp=float(read_sheet.cell(i,copy_coloum_index).value)#for generating error
          write_sheet.write(i,coloum_index,read_sheet.cell(i,copy_coloum_index).value,style4)
          continue
          ##go to bottom
        except ValueError:
          write_sheet.write(i,coloum_index,read_sheet.cell(i,copy_coloum_index).value,style3)
          continue

    write_sheet.write(i,coloum_index,read_sheet.cell(i,copy_coloum_index).value,style)
  return net_colomn
def delete_Coloum(coloum_index,row_ammount,write_sheet,file):
  style = XFStyle()
  style=Style.default_style
  for i in range(0,row_ammount+get_skiprows(file)+20):
    write_sheet.write(i,coloum_index,'',style)
def delete_bdx_Coloum(delete_list,rb,wb,file,bdx_type):
  #delete_list = [x+1 for x in delete_list]
  print("deleting colomn")
  r_sheet=rb.sheet_by_index(0)
  w_sheet = wb.get_sheet(0)
  net_colomn=0
  
  current_Index=0# index that walk to next coloum to keep
  for i in range(0,count_BdxL_Colomn(file)):
    if i in delete_list:
      continue
    net_colomn=copy_Coloum(current_Index,count_BdxL(file),i,r_sheet,w_sheet,file,net_colomn)
    current_Index+=1
  last_index=current_Index
  print("last_index",last_index)
  for i in range(len(delete_list)):
    delete_Coloum(current_Index,count_BdxL(file),w_sheet,file)
    current_Index+=1


  style3 = XFStyle()
  font3 = Font()
  Alignment3=Alignment()
  Alignment3.horz=Alignment.HORZ_CENTER
  Alignment3.vert=Alignment.VERT_CENTER
  #style3.alignment.wrap = 5
  style3.alignment=Alignment3
  borders3= Borders()
  borders3.left=Borders.MEDIUM
  borders3.right=Borders.MEDIUM
  borders3.top=Borders.MEDIUM
  borders3.bottom=Borders.MEDIUM
  #borders3.left_colour = Style.colour_map["gray25"]
  #borders3.right_colour =  Style.colour_map["gray25"]
  #borders3.top_colour =  Style.colour_map["gray25"]
  #borders3.bottom_colour =  Style.colour_map["gray25"]
  style3.borders=borders3
  font3.name = 'Arial'
  font3.height = 20*11
  font3.bold=True
  style3.font = font3
  pattern3 = Pattern()
  pattern3.pattern = Pattern.SOLID_PATTERN
  pattern3.pattern_fore_colour = Style.colour_map['gray25']
  style3.pattern = pattern3
  if bdx_type=="manual":
    w_sheet.write_merge(2,2,last_index-3,last_index-1,r_sheet.cell(3,14).value+"- MANUAL",style3)
  elif bdx_type=="web":
    w_sheet.write_merge(2,2,last_index-3,last_index-1,r_sheet.cell(3,14).value+"- WEB",style3)
  return net_colomn


def move_header(rb,wb,bdx_type,new_client,first_case_id):
  try:
    start=num_to_date(get_policy_date_list(first_case_id)[0][2])
    end=num_to_date(get_policy_date_list(first_case_id)[0][3])
    policy_date=start+" - "+end
  except:
    policy_date="Cant found in DB"
  print("copying header")
  read_sheet=rb.sheet_by_index(0)
  write_sheet = wb.get_sheet(0)
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

  style4 = XFStyle()
  font4 = Font()
  Alignment4=Alignment()
  #Alignment3.horz=Alignment.HORZ_CENTER
  Alignment4.vert=Alignment.VERT_CENTER
  style4.alignment.wrap = 5
  style4.alignment=Alignment4
  borders4= Borders()
  borders4.left=Borders.THIN
  borders4.right=Borders.THIN
  borders4.top=Borders.THIN
  borders4.bottom=Borders.THIN
  borders4.left_colour = Style.colour_map["gray25"]
  borders4.right_colour =  Style.colour_map["gray25"]
  borders4.top_colour =  Style.colour_map["gray25"]
  borders4.bottom_colour =  Style.colour_map["gray25"]
  style4.borders=borders4
  font4.name = 'Arial'
  font4.height = 20*10
  font4.bold=True
  font4.colour_index=Style.colour_map['light_blue']
  style4.font = font4
  pattern4 = Pattern()
  pattern4.pattern = Pattern.SOLID_PATTERN
  pattern4.pattern_fore_colour = Style.colour_map['white']
  style4.pattern = pattern4

  write_sheet.write_merge(1,1,0,5,read_sheet.cell(1,0).value,style3)
  write_sheet.write_merge(3,3,0,2,read_sheet.cell(3,0).value,style3)
  write_sheet.write_merge(4,4,0,2,read_sheet.cell(4,0).value,style3)
  write_sheet.write_merge(5,5,0,2,read_sheet.cell(5,0).value,style4)
  write_sheet.write_merge(6,6,0,2,read_sheet.cell(6,0).value,style3)
  write_sheet.write_merge(7,7,0,2,read_sheet.cell(7,0).value,style3)
  write_sheet.write_merge(8,8,0,2,read_sheet.cell(8,0).value,style3)

  write_sheet.write_merge(3,3,3,5,read_sheet.cell(3,4).value,style3)
  write_sheet.write_merge(4,4,3,5,new_client,style3)
  write_sheet.write_merge(5,5,3,5,policy_date,style4)
  write_sheet.write_merge(6,6,3,5,read_sheet.cell(6,4).value,style3)
  write_sheet.write_merge(7,7,3,5,fill_date(),style3)
  write_sheet.write_merge(8,8,3,5,read_sheet.cell(8,4).value,style3)

  if bdx_type=="web":
    write_sheet.write_merge(9,9,3,5,read_sheet.cell(9,4).value,style3)
def move_bottom(rb,wb,net_colomn,file,bdx_type,medi_base):
  list_label=list(read_BdxL(file).T.index.values)
  Case_index=int(list_label.index("Case Date"))
  print("copying bottom")
  read_sheet=rb.sheet_by_index(0)
  write_sheet = wb.get_sheet(0)
  bottom=get_skiprows(file)+count_BdxL(file)
  
  style3 = XFStyle()
  font3 = Font()
  Alignment3=Alignment()
  Alignment3.horz=Alignment.HORZ_RIGHT
  Alignment3.vert=Alignment.VERT_CENTER
  style3.alignment.wrap = 5
  style3.alignment=Alignment3
  font3.name = 'Arial'
  font3.height = 20*10
  font3.bold=True
  style3.font = font3

  style2 = XFStyle()
  font2 = Font()
  Alignment2=Alignment()
  Alignment2.horz=Alignment.HORZ_CENTER
  Alignment2.vert=Alignment.VERT_CENTER
  style2.alignment.wrap = 5
  Borders2=Borders()
  Borders2.bottom=Borders.DOUBLE
  style2.borders=Borders2
  style2.alignment=Alignment2
  font2.name = 'Arial'
  font2.height = 20*11
  font2.bold=True
  style2.font = font2

  style4 = XFStyle()
  font4 = Font()
  Alignment4=Alignment()
  Alignment4.horz=Alignment.HORZ_RIGHT
  Alignment4.vert=Alignment.VERT_CENTER
  style2.alignment.wrap = 5
  style4.alignment=Alignment4
  font4.name = 'Arial'
  font4.height = 20*10
  style4.font = font4

  style5 = XFStyle()
  borders5= Borders()
  borders5.bottom=Borders.THIN
  style5.borders=borders5

  style6 = XFStyle()
  font6 = Font()
  Alignment6=Alignment()
  Alignment6.horz=Alignment.HORZ_CENTER
  Alignment6.vert=Alignment.VERT_CENTER
  style2.alignment.wrap = 5
  style6.alignment=Alignment6
  font6.name = 'Arial'
  font6.height = 20*10
  style6.font = font6


  style7 = XFStyle()
  style7=Style.default_style
  net_colomn2=net_colomn
  write_sheet.write_merge(bottom+3,bottom+3,0,7,read_sheet.cell(bottom+3,0).value,style7)
  write_sheet.write_merge(bottom+3,bottom+3,net_colomn2-2,net_colomn2-1,read_sheet.cell(bottom+3,Case_index).value,style3)
  write_sheet.write_merge(bottom+9,bottom+9,net_colomn2-2,net_colomn2-1,read_sheet.cell(bottom+8,Case_index).value,style4)
  write_sheet.write_merge(bottom+15,bottom+15,net_colomn2-2,net_colomn2-1,read_sheet.cell(bottom+15,Case_index).value,style4)
  write_sheet.write(bottom+9,net_colomn2,'',style5)
  write_sheet.write(bottom+15,net_colomn2,'',style5)



  write_sheet.write(bottom+3,net_colomn2,read_sheet.cell(bottom+3,Case_index+2).value,style2)
  write_sheet.write(bottom+11,net_colomn2-1,read_sheet.cell(bottom+10,Case_index+1).value,style4)
  write_sheet.write(bottom+12,net_colomn2-1,read_sheet.cell(bottom+11,Case_index+1).value,style4)
  write_sheet.write(bottom+11,net_colomn2,read_sheet.cell(bottom+10,Case_index+2).value,style6)

  email=medi_base.email
  name=email.split("@")[0]
  first_name=email.split(".")[0]
  write_sheet.write(bottom+11,net_colomn2,first_name.upper(),style6)
  write_sheet.write(bottom+12,net_colomn2,str(get_Date()),style6)
  #write_sheet.write(bottom+12,net_colomn2,read_sheet.cell(bottom+11,Case_index+2).value,style6)
  write_sheet.write(bottom+17,net_colomn2-1,read_sheet.cell(bottom+17,Case_index+1).value,style4)
  #write_sheet.write(bottom+17,net_colomn2-1,read_sheet.cell(bottom+17,Case_index+2).value,style4)
  write_sheet.write(bottom+18,net_colomn2-1,read_sheet.cell(bottom+18,Case_index+1).value,style4)
  return read_sheet.cell(bottom+17,Case_index+2).value

def getListOfFiles(dirName):
    # create a list of file and sub directories 
    # names in the given directory 
    listOfFile = os.listdir(dirName)
    #print(listOfFile)
    allFiles = list()
    #client_name_split=Client.split(" ")
    #short_client_name=""
    #for i in client_name_split:
    #  short_client_name+=i[0]
    #print(short_client_name)

    # Iterate over all the entries
    for entry in listOfFile:
        # Create full path
        fullPath = os.path.join(dirName, entry)
         #If entry is a directory then get the list of files in this directory 
        if os.path.isdir(fullPath):
            allFiles = allFiles + getListOfFiles(fullPath)
        else:
            allFiles.append(fullPath)
    #    allFiles.append(fullPath)
    #print(fullPath)
    return allFiles
##print(getListOfFiles(get_env()))
#file_object  = open("data11.txt", "r")
#string=file_object.read()
#list=string[1:-1].split(",")
#for i in list:
#  if "AEONASIA"+"-" in i:
#    print(i)
def get_all_File(file_path,file_type):
  if file_type=="xls":
    list_of_files = glob.glob(file_path+"\*.xls") # * means all if need specific format then *.csv
  elif file_type=="pdf":
    list_of_files = glob.glob(file_path+"\*.pdf") # * means all if need specific format then *.csv
  elif file_type=="zip":
    list_of_files = glob.glob(file_path+"\*.zip") # * means all if need specific format then *.csv
  return list_of_files  
def index_to_delete(template,target,bdx_type):
  print(template)
  old_df=read_BdxL_template(template,template_skiprows(template))
  new_df=read_BdxL(target)
  index_counter=0
  index_list=[]
  old_label=old_df.columns.values
  new_label=new_df.columns.values
  print(len(old_label))
  print(new_label)
  for index in range(len(new_label)):
    index_list.append(index)
    if index_counter<len(old_label) and old_label[index_counter]==new_label[index] and index_counter<=len(old_label):
      index_list.pop()
      index_counter+=1
  return index_list
def hasNumbers(inputString):
    if all(char.isdigit() for char in inputString):
      return False
    else:
      return True

def fix_excel(file_name,file_path):
  file_name=file_name
  file_path=file_path
  download_path=file_path.replace(file_name,'')
  df = pd.read_table(file_path)
  df.to_excel(download_path+"html_format.xls")

  rb=open_workbook(download_path+"html_format.xls")
  read_sheet=rb.sheet_by_index(0)
  strin=""
  try:
    for i in range(100):
      strin=strin+read_sheet.cell(i,1).value
  except:
    pass

  table = pd.read_html(strin)[0]
  Total_row=len(table.index)
  table.to_excel(download_path+"raw_version.xls",index=False)
  rb=open_workbook(download_path+"raw_version.xls",formatting_info=True,ragged_rows=True)
  wb=copy(rb)
  writerSheet = wb.get_sheet(0)
  No_index=0
  readerSheet = rb.sheet_by_index(0)
  while "Kindly note that the payment term is 14 (fourteen" not in readerSheet.cell(No_index,0).value:
    No_index+=1
  last_column=0
  try:
    for i in range(100):
      if readerSheet.cell(0,i).value !="":
        last_column=i
      writerSheet.write(0,i," ")
  except:
    pass
  dup_list=[]
  for i in range(Total_row):
    for j in range(last_column):
      if readerSheet.cell(i,j).value!="" and readerSheet.cell(i,j).value!="0.00" and j<last_column-1 and readerSheet.cell(i,j).value==readerSheet.cell(i,j+1).value and readerSheet.cell(i,j).value not in dup_list and readerSheet.cell(i,j).value!="0" and hasNumbers(readerSheet.cell(i,j).value):
        duplicate=0
        dup_list.append(readerSheet.cell(i,j).value)
        for dup in range(20):
          if readerSheet.cell(i,j).value==readerSheet.cell(i,j+dup).value:
            duplicate+=1
          else:
            break
        for remove in range(duplicate):
          writerSheet.write(i,j+remove," ")
        writerSheet.write_merge(i,i,j,j+duplicate-1,readerSheet.cell(i,j).value)
  os.remove(download_path+"html_format.xls")
  os.remove(download_path+"raw_version.xls")
  wb.save(file_path)

def bdx_automation(file_name,BDFile,template_path,medi_base,path=None):
  if path==None:
    path=BDFile
  target=BDFile
  try:
    rb=open_workbook(BDFile,formatting_info=True)
  except:
    fix_excel(file_name,BDFile)
    rb=open_workbook(BDFile,formatting_info=True)
  readerSheet = rb.sheet_by_index(0)
  for Corporate_row in range(30):
    if "Corporate" in readerSheet.cell(Corporate_row,0).value:
      break
  bdx_type=""
  template=""
  file_name_search=file_name.split("-")[0]
  if os.path.exists(template_path+r"\file_dir.txt") != True:
    print("gg")
    all_path=getListOfFiles(template_path)
    file_object  = open(template_path+r"\file_dir.txt", "w")
    file_object.write(str(all_path))
    file_object.close()

  file_dir  = open(template_path+r"\file_dir.txt", "r")
  string=file_dir.read()
  list=string[1:-1].split(", ")
  print(file_name_search)
  if "WEB" in BDFile:
    bdx_type="web"
    bdx_type2="(W)"
  elif "MANUAL" in BDFile:
    bdx_type="manual"
    bdx_type2="(M)"

  for i in list:
    if file_name_search+"-" in i:
      template=(i[1:-1]).replace("\\\\",'\\')
      if bdx_type.upper() in i or bdx_type2 in i:
        template=(i[1:-1]).replace("\\\\",'\\')
        break
  audit_log("Template for client excel file found", "Completed...", medi_base)
  client=readerSheet.cell(Corporate_row,4).value
  dc_template_path=os.path.dirname(template)
  list_of_current_template_floder=get_all_File(dc_template_path,"xls")
  for i in list_of_current_template_floder:
    if "MCL" in i:
      dc_template_path=i
      if  bdx_type.upper() in i or  bdx_type2 in i:
        dc_template_path=i
        break
  audit_log("Template for disbursement listing file found", "Completed...", medi_base)


  delete_list=index_to_delete(template,target,bdx_type)
  rb=build_rb(target)
  wb=build_wb(target)
  net_colomn=delete_bdx_Coloum(delete_list,rb,wb,target,bdx_type)
  move_header(rb,wb,bdx_type,client,get_first_case_id(target))
  name=move_bottom(rb,wb,net_colomn,target,bdx_type,medi_base)
  audit_log("bdx Excel is updated", "Completed...", medi_base)
  savefile(wb,path)
  autoAdjustColumns(path,500,bdx_type)
  return name,dc_template_path


#bdx_automation("SPIRITAERO-0017 MANUAL.xls",r"\\dtisvr2\mediclinic\New\Admission_354\New\SPIRITAERO-0017 MANUAL.xls","\\\\dtisvr2\\mediclinic\\Client",r"\\dtisvr2\mediclinic\New\Admission_354\New\Sok.xls")

#try:
#with open(xls, 'r') as upxls:
#	dbx.files_upload(file_path.read(),'/test.xlsx', mode=dropbox.files.WriteMode("overwrite"))
# except Exception as e:
#	  print e	
