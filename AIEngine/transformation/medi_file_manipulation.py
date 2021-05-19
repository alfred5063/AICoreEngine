import win32com.client as win32
import zipfile
import glob
import os
from automagica import *
import pythoncom
from connector.connector import MSSqlConnector
from directory.createfolder import createfolder
from shutil import copyfile
from utils.audit_trail import audit_log
from datetime import datetime

def get_Lastest_File(file_path,file_type):
  if file_type=="xls":
    list_of_files = glob.glob(file_path+"\*.xls") # * means all if need specific format then *.csv
  elif file_type=="xlsm":
    list_of_files = glob.glob(file_path+"\*.xlsm") # * means all if need specific format then *.csv
  elif file_type=="xlsx":
    list_of_files = glob.glob(file_path+"\*.xlsx") # * means all if need specific format then *.csv
  elif file_type=="zip":
    list_of_files = glob.glob(file_path+"\*.zip") # * means all if need specific format then *.csv
  latest_file = sorted(list_of_files, key=os.path.getctime)[-1]
  return latest_file

def get_all_File(file_path,file_type):
  if file_type=="xls":
    list_of_files = glob.glob(file_path+"\*.xls") # * means all if need specific format then *.csv
  elif file_type=="pdf":
    list_of_files = glob.glob(file_path+"\*.pdf") # * means all if need specific format then *.csv
  elif file_type=="zip":
    list_of_files = glob.glob(file_path+"\*.zip") # * means all if need specific format then *.csv
  return list_of_files

def get_Second_Lastest_File(file_path,file_type):
  if file_type=="xls":
    list_of_files = glob.glob(file_path+"\*.xls") # * means all if need specific format then *.csv
  elif file_type=="xlsm":
    list_of_files = glob.glob(file_path+"\*.xlsm") # * means all if need specific format then *.csv
  elif file_type=="zip":
    list_of_files = glob.glob(file_path+"\*.zip") # * means all if need specific format then *.csv
  second_latest_file = sorted(list_of_files, key=os.path.getctime)[-2]
  return second_latest_file

def move_to_archive(file,path=None,destination=None):
  #if os.path.exists(file.replace("New","Archived")):
  #    os.remove(file.replace("New","Archived"))
  #os.rename(file, file.replace("New","Archived"))
  if destination !=None:
    try:
      if os.path.exists(file.replace(path,destination)):
        os.remove(file.replace(path,destination))
      # Move a file by renaming it's path
      os.rename(file, file.replace(path,destination))
    except:
      pass


def all_excel_move_to_archive(path):
  try:
    all_file_in_excel=get_all_File(path,"xls")
    for file in all_file_in_excel:
      if os.path.exists(file.replace("Result","Archived")):
          os.remove(file.replace("Result","Archived"))
      os.rename(file, file.replace("Result","Archived"))
  except:
    pass

def get_excel_path(file):
  path=file.replace("New","Archieved")
  path=path.replace("Archieved","New",1)
  return path

def extract_zip(path,download_achieve):
  print('zip path: {0}'.format(path))
  path_to_zip_file=get_Lastest_File(path,"zip")
  with zipfile.ZipFile(path_to_zip_file, 'r') as zip_ref:
     zip_ref.extractall(path)
  latest_excel=get_Lastest_File(path,"xls")
  move_to_archive(path_to_zip_file,path,download_achieve)
  move_to_archive(latest_excel,path,download_achieve)
  file_path=path_to_zip_file.replace("zip","xls")
  file_path=file_path.replace(path,download_achieve)
  file_name=path_to_zip_file.replace(path,"")
  file_name=file_name.replace(".zip","")
  file_name=file_name.replace("\\","")
  return file_path,file_name

def get_file_from_download(path,download_achieve):
  path_to_zip_file=get_Lastest_File(path,"xlsx")
  move_to_archive(path_to_zip_file,path,download_achieve)
  file_path=path_to_zip_file
  file_name=path_to_zip_file.replace(path,"")
  file_name=file_name.replace(".xlsx","")
  file_name=file_name.replace("\\","")
  return file_path,file_name

def get_file_from_download_zip(path,download_achieve):
  path_to_zip_file=get_Lastest_File(path,"zip")
  move_to_archive(path_to_zip_file,path,download_achieve)
  file_path=path_to_zip_file
  file_name=path_to_zip_file.replace(path,"")
  file_name=file_name.replace(".zip","")
  file_name=file_name.replace("\\","")
  return file_path,file_name

def fix_excel(raw_path,ori_path):
  pythoncom.CoInitialize()
  xl = win32.DispatchEx("Excel.Application")
  #xl = win32.gencache.EnsureDispatch('Excel.Application')
  xl.DisplayAlerts = False
  xl.Visible = 1
  xl.Visible = True 
  wb = xl.Workbooks.Open(raw_path)
  xl.WindowState = win32.constants.xlMaximized
  new_path=ori_path
  print(new_path)
  Wait(4)
  wb.SaveAs(new_path, FileFormat = 56)    #FileFormat = 51 is for .xlsx extension
  wb.Close()
  xl.Application.Quit()
  xl.Quit()
  del xl

#path=r"\\dtisvr2\RPA_CBA\Folder\MEDICLINIC\download"
def file_manipulation(path,download_new,download_achieve,medi_base,bdx_type):
  print('file_manipulation start')
  base_download_achieve = download_achieve
  now = datetime.now().strftime('%d%m%Y%H%M%S')
  timestamp = str(now)
  #create specific timestamp for a achieve folder
  download_achieve = download_achieve+'\\'+timestamp
  if(os.path.isdir(download_achieve)==False):
      createfolder(download_achieve, medi_base, parents=True, exist_ok=True)

  file_path,file_name=extract_zip(path,download_achieve)
  print(file_path,file_name,"extract")
  userid=file_name.split("_")[-1]
  #move_to_archive(file_path,download_new,download_achieve)
  new_file_name=file_name.replace("_"+userid.upper()," "+bdx_type.upper()+".xls")
  new_file_path=get_excel_path(file_path).replace("_"+userid.upper()," "+bdx_type.upper()).replace(timestamp+"\\", '')
  print(new_file_name,new_file_path,"new")
  if os.path.exists(new_file_path):
    os.remove(new_file_path)
  print(file_path,new_file_path,"rename")
  os.rename(file_path,new_file_path)
  #copyfile(get_excel_path(new_file_path),download_new+"\\"+new_file_name)
  audit_log("File is saved as name {}".format(new_file_name), "Completed...", medi_base)
  print('file_manipulation ended: file saved as name {0}'.format(new_file_name))
  return new_file_name,download_new+"\\"+new_file_name, download_achieve

def file_manipulation_crm(path,download_achieve,medi_base=None):
  #userid=medi_base.username
  file_path,file_name=get_file_from_download(path,download_achieve)
  #move_to_archive(file_path,path,download_achieve)
  #new_file_name=file_name.replace("_"+userid.upper()," "+bdx_type.upper())
  #new_file_path=get_excel_path(file_path).replace("_"+userid.upper()," "+bdx_type.upper())
  #print(file_path,file_name)
  #if os.path.exists(file_path):
  #os.remove(file_path)
  #os.rename(get_excel_path(file_path),file_path)
  copyfile(file_path.replace(path,download_achieve),file_path)
  audit_log("File is saved as name {}".format(file_name), "Completed...", medi_base)
  return file_name,file_path

def get_Task_Name(taseid):
  connector = MSSqlConnector()
  cursor = connector.cursor()
  qry="SELECT title FROM dbo.tasks where id={}".format(str(taseid))
  cursor = connector.cursor()
  cursor.execute(qry)
  taskname = cursor.fetchall()
  return taskname[0][0]



def create_path(print_path,medi_base):
  if os.path.exists(print_path) != True:
    createfolder(print_path,medi_base)
  if medi_base.email==None or medi_base.email=="":
    task_name_path=print_path+"\\{}".format(get_Task_Name(medi_base.taskid)+"_"+medi_base.taskid)
  else:
    email=medi_base.email.split("@")[0]
    task_name_path=print_path+"\\{}".format(get_Task_Name(medi_base.taskid)+"_"+email)
  if os.path.exists(task_name_path) != True:
    createfolder(task_name_path,medi_base)

  print_archieve=task_name_path+"\\Archieved"
  print_new=task_name_path+"\\New"

  if os.path.exists(print_archieve) != True:
    createfolder(print_archieve,medi_base)
  if os.path.exists(print_new) != True:
    createfolder(print_new,medi_base)
  return print_new,print_archieve

def move_file(path,des):
   if des !=None:
    try:
      if os.path.exists(des):
        os.remove(des)
      # Move a file by renaming it's path
      os.rename(path, des)
    except:
      pass
