
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
#from transformation.medi_read_medi_config import get_skiprows
#from transformation.file_manipulation import get_excel_path
#from transformation.medi_mask_mapping import mask_mapping
from datetime import date,datetime,timedelta
import datetime as dt
from dateutil.relativedelta import relativedelta
from connector.connector import MSSqlConnector
import glob
import pandas as pd
import traceback
def get_EventDate():
  connector = MSSqlConnector()
  cursor = connector.cursor()
  qry="SELECT date FROM cba.calendar_events"
  cursor = connector.cursor()
  cursor.execute(qry)
  records = cursor.fetchall()
  cursor.close()
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
  rb=build_rb(file)
  readerSheet = rb.sheet_by_index(0)
  skiprow=0
  while "No" != readerSheet.cell(skiprow,0).value:
    skiprow+=1
  data=pd.read_excel(file,skiprows = skiprow , na_values = "Missing")
  
#code can be improved by flaxible by change the column and row incase marc change formula
  data1=(data.loc[data['No'].isin([i for i in range(1000000)])]).copy()
  return data1
def get_skiprows(file):
  rb=build_rb(file)
  readerSheet = rb.sheet_by_index(0)
  skiprow=0
  while "No" != readerSheet.cell(skiprow,0).value:
    skiprow+=1
  return skiprow
#file=r"C:\Users\junshou.leow\Desktop\Client\AIRASIA\01 JAN\02-01-2019\AIRASIA9054-2018-12 (W).xls"
#file=r"C:\Users\junshou.leow\Downloads\AIRASIA115131-2019-11 (M).xls"
#df=read_BdxL(file)
#print(df)
def read_BdxL_template(file):
  """
  This function read the excel for all listing available

  Arguement is location of excel file

  End by return a dataframe of the listing
  """
  
#code can be improved by flaxible by change the column and row incase marc change formula
  data1=read_BdxL(file)
  return data1
def count_BdxL(file):

  """
  This function read the excel for all listing available

  Arguement is location of excel file

  End by return dataframe index number
  """
  
#code can be improved by flaxible by change the column and row incase marc change formula

  data1=read_BdxL(file)
  #print(data1)
  #row=0
  #for i in data1['No'].tolist():
  #  if (type(i)!=type("")):
  #    break
  #  row+=1
  return len(data1.index)
def count_BdxL_Colomn(file):
  """
  This function read the excel for all listing available

  Arguement is location of excel file

  End by return dataframe index number
  """
#code can be improved by flaxible by change the column and row incase marc change formula
  
  data1=read_BdxL(file).T

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
    readerSheet = rd.sheet_by_index(0)
    skiprow=0
    while "No" != readerSheet.cell(skiprow,0).value:
      skiprow+=1
    print("adjustting colomn")

    starting_point= skiprow-1
    ending_point=skiprow+1+count_BdxL(path)
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
def find_total_index(file):
  #read_BdxL(file,bdx_type)
  rb=build_rb(file)
  read_sheet=rb.sheet_by_index(0)
  row=get_skiprows(file)+count_BdxL(file)+3
  #column=read_BdxL(file,bdx_type).columns.get_loc("Nett Total (RM)")
  column=0
  string=""
  while True:
    string=read_sheet.cell(row,column).value
    print(string)
    try:
      if string.isdecimal():
        print(string)
      
        break
    except:
      break
    column+=1
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
  font1.colour_index=Style.colour_map['black']
  style1.font = font1
  pattern1 = Pattern()
  pattern1.pattern = Pattern.SOLID_PATTERN
  pattern1.pattern_fore_colour = Style.colour_map['light_blue']
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
  
  #write_sheet.write_merge(get_skiprows(bdx_type)+2+row_ammount,get_skiprows(bdx_type)+2+row_ammount,0,7,read_sheet.cell(get_skiprows(bdx_type)+2+row_ammount,copy_coloum_index).value,style2)
  skiprow=get_skiprows(file)
  for i in range(skiprow,row_ammount+skiprow+2):
    if i==skiprow:

      write_sheet.write(i,coloum_index,read_sheet.cell(i,copy_coloum_index).value,style1)
      if "Total Payable" in read_sheet.cell(i,copy_coloum_index).value:
        net_colomn=coloum_index
      continue

    if i>row_ammount+skiprow:
      if i==row_ammount+skiprow+1:
        try:
          temp=int(read_sheet.cell(i,copy_coloum_index).value)#for generating error
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
  
  current_Index=0 #index that walk to next coloum to keep
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
  #if bdx_type=="manual":
  #  w_sheet.write_merge(2,2,last_index-3,last_index-1,r_sheet.cell(3,14).value+"- MANUAL",style3)
  #elif bdx_type=="web":
  #  w_sheet.write_merge(2,2,last_index-3,last_index-1,r_sheet.cell(3,14).value+"- WEB",style3)
  #w_sheet.write(3,14,"")
  return net_colomn
def move_header(rb,wb,bdx_type,new_client):
  #delete_list = [x+1 for x in delete_list]
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

  First_No=0
  while "No" != read_sheet.cell(First_No,0).value:
    First_No+=1
  for i in range(1,First_No):
    if read_sheet.cell(i,0).value=="Year of Policy:":
      write_sheet.write_merge(i,i,0,1,read_sheet.cell(i,0).value,style4)
      write_sheet.write_merge(i,i,2,3,read_sheet.cell(i,3).value,style4)
      continue
    write_sheet.write_merge(i,i,0,1,read_sheet.cell(i,0).value,style3)
    write_sheet.write_merge(i,i,2,3,read_sheet.cell(i,3).value,style3)


def move_bottom(rb,wb,net_colomn,file,bdx_type):
  list_label=list(read_BdxL(file).T.index.values)
  Case_index=int(list_label.index("Effective Date"))
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
  print(read_sheet.cell(bottom+17,Case_index+1).value)
  write_sheet.write_merge(bottom+9,bottom+9,net_colomn2-2,net_colomn2-1,read_sheet.cell(bottom+8,Case_index).value,style4)
  write_sheet.write_merge(bottom+15,bottom+15,net_colomn2-2,net_colomn2-1,read_sheet.cell(bottom+15,Case_index).value,style4)
  write_sheet.write(bottom+9,net_colomn2,'',style5)
  write_sheet.write(bottom+15,net_colomn2,'',style5)
  print(read_sheet.cell(bottom+8,Case_index).value,"VALUE")



  write_sheet.write(bottom+3,net_colomn2,read_sheet.cell(bottom+3,Case_index+2).value,style2)
  write_sheet.write(bottom+11,net_colomn2-1,read_sheet.cell(bottom+10,Case_index+1).value,style4)
  write_sheet.write(bottom+12,net_colomn2-1,read_sheet.cell(bottom+11,Case_index+1).value,style4)
  write_sheet.write(bottom+11,net_colomn2,read_sheet.cell(bottom+10,Case_index+2).value,style6)
  write_sheet.write(bottom+12,net_colomn2,read_sheet.cell(bottom+11,Case_index+2).value,style6)
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
def get_all_File(file_path,file_type):
  if file_type=="xls":
    list_of_files = glob.glob(file_path+"\*.xls") # * means all if need specific format then *.csv
  elif file_type=="pdf":
    list_of_files = glob.glob(file_path+"\*.pdf") # * means all if need specific format then *.csv
  elif file_type=="zip":
    list_of_files = glob.glob(file_path+"\*.zip") # * means all if need specific format then *.csv
  return list_of_files
column_mapping={'No':'No',
                'Claim Id':'Claim Id',
                'Employee Id':'Employee Id',
                'Member':'Member',
                'Parent Contact':'Parent Contact',
                'MC':'MC Duration',
                'Date Visit':'Date Visit',
                'Clinic Name':'Clinic Name',
                'Diagnosis':'Diagnosis',
                'ASO Amount (RM)':'ASO Amount (RM)',
                'Member Amount (RM)':'Member Amount (RM)',
                'GST Total (RM)':'ASO GST Total (RM)',
                'Total Payable (RM)':'Total Payable To Clinic (RM)',
                'Available Balance (RM)':'Available Balance (RM)',
                'Clinical Service':'Clinical Service'
  }
def index_to_delete(template,target,bdx_type):
  old_df=read_BdxL_template(template)
  new_df=read_BdxL(target)
  
  index_counter=0
  index_list=[]
  old_label=old_df.columns.values
  new_label=new_df.columns.values
  print(old_label)
  print(len(new_label))
  print(len(old_label))
  print(new_label)
  for index in range(len(new_label)):
    index_list.append(index)
    #print(old_label[index_counter])
    if index_counter<len(old_label) and index_counter!=len(old_label):
      if new_label[index].strip() == old_label[index_counter].strip() or new_label[index].strip()==column_mapping[old_label[index_counter].strip()]:
        index_list.pop()
        index_counter+=1
  return index_list
def hasNumbers(inputString):
    if all(char.isdigit() for char in inputString):
      return False
    else:
      return True
def hasNumbers(inputString):
    if all(char.isdigit() for char in inputString):
      return False
    else:
      return True
def fix_excel(file_name,file_path,bdx_type):
  is_insurance=False
  if "insurance" in file_path:
    is_insurance=True
  file_name=file_name
  file_path=file_path
  download_path=file_path.replace(file_name+".xlsx",'')

  print(download_path)

  table = pd.read_excel(file_path)
  Total_row=len(table.index)
  table.to_excel(download_path+"raw.xls",index=False)
  rb=open_workbook(download_path+"raw.xls",formatting_info=True,ragged_rows=True)
  wb=copy(rb)
  writerSheet = wb.get_sheet(0)
  No_index=0
  readerSheet = rb.sheet_by_index(0)
  df = pd.read_excel(download_path+"raw.xls")
  df.to_excel(download_path+"raw.xls",index=False)
  rb=open_workbook(download_path+"raw.xls",formatting_info=True,ragged_rows=True)
  wb=copy(rb)
  writerSheet = wb.get_sheet(0)
  No_index=0
  readerSheet = rb.sheet_by_index(0)
  if is_insurance:
    while "Kindly note that the payment term is" not in str(readerSheet.cell(No_index,0).value):
      No_index+=1
  else:
    while "Kindly note that the payment term is" not in str(readerSheet.cell(No_index,1).value):
      No_index+=1

  
  
  First_No=0
  while "No" != readerSheet.cell(First_No,0).value:
    First_No+=1


  data=pd.read_excel(download_path+"raw.xls",skiprows =First_No  , na_values = "Missing")
  data1=(data.loc[data['No'].isin([i for i in range(1000)])]).copy()
  row_amount=len(data1['No'].tolist())


  last_column=0


  try:
    for i in range(100):
      if readerSheet.cell(0,i).value !="":
        last_column=i
      writerSheet.write(0,i," ")
  except:
    pass

  
  if is_insurance:
    start_column=0
  else:
    start_column=2
  remove_index=0
  print(First_No,First_No+row_amount+1,"xx")
  total_payable=""
  try:
    for x in range(start_column,last_column):
      while readerSheet.cell(First_No,x+remove_index).value=="":
        remove_index+=1
      for y in range(First_No,First_No+row_amount+2):
        if x==2 and is_insurance==False:
          writerSheet.write(y,x-1,readerSheet.cell(y,x-2).value)
        writerSheet.write(y,x,readerSheet.cell(y,x+remove_index).value)
        if "Total Payable" in readerSheet.cell(First_No,x+remove_index).value:
          total_payable=readerSheet.cell(First_No+row_amount+1,x+remove_index).value
        
  except:
    pass
  #wb.save(download_path+"raw.xls")
  #os.remove(download_path+"raws.xls")
  if is_insurance:
    pass
  else:
    print("ASO")
    try:
      for x in range(2+last_column-remove_index,last_column):
        for y in range(First_No,First_No+row_amount+2):
          writerSheet.write(y,x,"")
    except:
      pass
  print(total_payable)



  os.remove(download_path+"raw.xls")
  wb.save(download_path+"raw.xls")
  rb=open_workbook(download_path+"raw.xls",formatting_info=True,ragged_rows=True)
  wb=copy(rb)
  writerSheet = wb.get_sheet(0)
  readerSheet = rb.sheet_by_index(0)
  if is_insurance:
    pass
  else:
    try:
       for x in range(0,last_column):
         for y in range(0,First_No+row_amount+10):
            writerSheet.write(y,x,readerSheet.cell(y,x+1).value)
            writerSheet.write(y,x+1,"")
    except:
      pass

  writerSheet.write(First_No+row_amount+3,19,total_payable)
  writerSheet.write(First_No+row_amount+12-2,18,"Name:")
  writerSheet.write(First_No+row_amount+13-2,18,"Date:")
  writerSheet.write(First_No+row_amount+19-2,18,"Name:")
  writerSheet.write(First_No+row_amount+20-2,18,"Date:")

  writerSheet.write_merge(First_No+row_amount+3,First_No+row_amount+3,17,18,"Amount Due to AA =")
  writerSheet.write_merge(First_No+row_amount+10-2,First_No+row_amount+10-2,17,18,"Prepare By:")
  writerSheet.write_merge(First_No+row_amount+17-2,First_No+row_amount+17-2,17,18,"Acknowledged By:")
  writerSheet.write_merge(3,3,14,15,"MEDICLINIC")

  wb.save(download_path+"raw.xls")


  rb=open_workbook(download_path+"raw.xls",formatting_info=True,ragged_rows=True)
  wb=copy(rb)
  readerSheet = rb.sheet_by_index(0)
  writerSheet = wb.get_sheet(0)
  row_Bordereaux_No=0
  column_Bordereaux_No=1
  print("?")
  while "Borderaux No" not in readerSheet.cell(row_Bordereaux_No,0).value:
    row_Bordereaux_No+=1
  while readerSheet.cell(row_Bordereaux_No,column_Bordereaux_No).value=="":
    column_Bordereaux_No+=1
  print("?")
  today = date.today()
  d1 = today.strftime("-%Y/%m")
  bord_no=(readerSheet.cell(row_Bordereaux_No,column_Bordereaux_No).value).replace(str(d1),"1")
  print(readerSheet.cell(row_Bordereaux_No,column_Bordereaux_No).value)
  print(bord_no)
  year=today.strftime("%Y")
  month=today.strftime("%m")
  print(year)
  print(month)
  #bord_name = ''.join([i for i in bord_no if not i.isdigit()])
  if bdx_type.upper()=="MANUAL":
    bdx_name_convention='(M)'
  elif bdx_type.upper()=="WEB":
    bdx_name_convention="(W)"



  if is_insurance:
    for i in range(First_No):
      try:
        print(readerSheet.cell(i,2).value)
        writerSheet.write(i,2,"")
        writerSheet.write(i,3,readerSheet.cell(i,2).value)
      except:
        pass




  file_name=bord_no+"-"+year+"-"+month+" "+bdx_name_convention
  print(file_name)
  wb.save(download_path+file_name+".xls")
  print(file_name)
  os.remove(download_path+"raw.xls")
  print(file_name)

  return file_name+".xls",download_path+file_name+".xls"










def bdx_automation(file_name,BDFile,template_path,bdx_type,medi_base,path=None):
  print(file_name,BDFile,template_path,bdx_type)
  try:
    file_name,BDFile=fix_excel(file_name,BDFile,bdx_type)
    #print(file_name,BDFile)
    audit_log("Excel successfully converted", "Completed...", medi_base)
  except:
    print(traceback.format_exc())
    audit_log("Failed when converting excel", "Completed...", medi_base)
    audit_log("NO data in bdx", "Completed...", medi_base)
    return False,False,False
  
  if path==None:
    path=BDFile
  target=BDFile
  rb=open_workbook(BDFile,formatting_info=True)
  readerSheet = rb.sheet_by_index(0)

  for Corporate_row in range(30):
    if "Corporate" in readerSheet.cell(Corporate_row,0).value:
      break

  bdx_type=""
  template=""
  file_name_search=file_name.split("-")[0]
  file_name_search = ''.join([i for i in file_name_search if not i.isdigit()])
  if os.path.exists(template_path+r"\file_dir.txt") != True:
    print("gg")
    all_path=getListOfFiles(template_path)
    file_object  = open(template_path+r"\file_dir.txt", "w")
    file_object.write(str(all_path))
    file_object.close()

  file_dir  = open(template_path+r"\file_dir.txt", "r")
  string=file_dir.read()
  list=string[1:-1].split(", ")
  print(file_name_search,"file_name_search")
  print(BDFile,"BDFILE")
  if "WEB"  in file_name  or "(W)" in file_name:
    bdx_type="web"
    bdx_type2="(W)"
  elif "MANUAL"  in file_name or "(M)" in file_name:
    bdx_type="manual"
    bdx_type2="(M)"

  for i in list:
    if file_name_search in i:
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
  #print(client)
  #if client[-1] and client[-2] is "X":
  #  new_client=mask_mapping(template,client)
  #print(new_client)
  #template=template+"\\"+new_client+".xls"
  print("template:",template)
  print("bdx_type:",bdx_type)
  print("dc_template_path:",dc_template_path)


  delete_list=index_to_delete(template,target,bdx_type)
  print(len(delete_list))
  rb=build_rb(target)
  wb=build_wb(target)
  net_colomn=delete_bdx_Coloum(delete_list,rb,wb,target,bdx_type)
  move_header(rb,wb,bdx_type,client)
  name=move_bottom(rb,wb,net_colomn,target,bdx_type)
  audit_log("bdx Excel is updated", "Completed...", medi_base)
  savefile(wb,path)
  autoAdjustColumns(path,500,BDFile)

  return dc_template_path,name,path


#bdx_automation("Borderaux Listing ASO (3)",r"\\dtisvr2\mediclinic\New\Medi CRM_4396\New\Borderaux Listing ASO (3).xlsx",r"\\dtisvr2\mediclinic\Client","MANUAL")
#file_path=r"C:\Users\junshou.leow\Downloads\Borderaux Listing ASO.xlsx"
#file_name="Borderaux Listing ASO.xlsx"
#bdx_type="manual"
#template_path=r"C:\Users\junshou.leow\Desktop\Client"
#dc_template_path,name=bdx_automation(file_name,file_path,template_path,bdx_type,r"C:\Users\junshou.leow\Desktop\ss.xls")
