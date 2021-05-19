#!/usr/bin/python
# FINAL SCRIPT updated as of 18th nov 2019
# Workflow - CBA/ADMISSION
# Version 1
from win32com import client
import os
import win32api
def excel_to_pdf(file,sheet_name,path):
  pythoncom.CoInitialize()
  os.system("taskkill /f /im  excel.exe")
  if os.path.exists(path):
    os.remove(path)  
  xlApp = client.DispatchEx("Excel.Application")
  xlApp.DisplayAlerts = False
  books = xlApp.Workbooks.Open(file)
  ws = books.Worksheets[sheet_name]
  ws.Visible = 1
  ws.Visible = True 
  ws.ExportAsFixedFormat(0, path )
  books.Close()
  xlApp.Quit()
  xlApp.Application.Quit()
  del xlApp


#import win32com.client as win32
#fname = r"C:\Users\junshou.leow\Source\Repos\RoboticProcessAutomation\RPA.DTI\TESTING\junshou\download\singBordTemplate.xls"
#excel = win32.DispatchEx('Excel.Application')
#excel.DisplayAlerts = False
#wb = excel.Workbooks.Open(fname)
#wb.SaveAs(r"C:\Users\junshou.leow\Source\Repos\RoboticProcessAutomation\RPA.DTI\TESTING\junshou\download\sohai.xls", FileFormat = 56)    #FileFormat = 51 is for .xlsx extension
#wb.Close()                               #FileFormat = 56 is for .xls extension
#excel.Application.Quit()
