from automagica import *
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException,TimeoutException,NoSuchWindowException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
import os
from selenium.webdriver.common.keys import Keys
from datetime import date,datetime,timedelta
import datetime as dt
from dateutil.relativedelta import relativedelta
from connector.connector import MSSqlConnector
import winreg
import asyncio
import threading
from multiprocessing import Process
from selenium.webdriver.common.action_chains import ActionChains
from configparser import ConfigParser
import time
import random,glob
from utils.audit_trail import audit_log
import win32gui
def callback(hwnd, extra):
    rect = win32gui.GetWindowRect(hwnd)
    x = rect[0]
    y = rect[1]
    w = rect[2] - x
    h = rect[3] - y
    print("Window %s:" % win32gui.GetWindowText(hwnd))
    print("\tLocation: (%d, %d)" % (x, y))
    print("\t    Size: (%d, %d)" % (w, h))

def set_interval(browser, sec):
    def func_wrapper():
        set_interval(close_Medi(browser), sec)
        close_Medi(browser)
    t = threading.Timer(sec, func_wrapper)
    t.start()
    return t

def find_key(key,value,data):
  reg_connection = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
  try:
      key = winreg.OpenKey(reg_connection, key,0,winreg.KEY_SET_VALUE)
      winreg.SetValueEx(key,value, 0, winreg.REG_SZ, str(data))
  except FileNotFoundError:
      import platform
      bitness = platform.architecture()[0]
      if bitness == '32bit':
          other_view_flag = winreg.KEY_WOW64_64KEY
      elif bitness == '64bit':
          other_view_flag = winreg.KEY_WOW64_32KEY
      try:
          key = winreg.OpenKey(reg_connection, key, access=winreg.KEY_SET_VALUE | other_view_flag)
          winreg.SetValueEx(key,value, 0, winreg.REG_SZ, str(data))
          print("another os")
      except FileNotFoundError:

        print(key,"error")
  

def setup_env(path):
  #download_prompt_key=r"Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppContainer\Storage\microsoft.microsoftedge_8wekyb3d8bbwe\MicrosoftEdge\Download\\"
  download_path_key=r"Software\Microsoft\Internet Explorer\Main\\"
  
  download_prompt_key2=r"Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppContainer\Storage\microsoft.microsoftedge_8wekyb3d8bbwe\MicrosoftEdge\Main\EnableSavePrompt"
  #download_path_key=r"Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppContainer\Storage\microsoft.microsoftedge_8wekyb3d8bbwe\MicrosoftEdge\Main\\"
  offer_save_password=r"Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppContainer\Storage\microsoft.microsoftedge_8wekyb3d8bbwe\MicrosoftEdge\Main\\"
  ie_key=r"Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BFCACHE"

  #find_key(download_prompt_key,'EnableSavePrompt',0)
  #find_key(download_prompt_key2,'(Default)',0)
  find_key(download_path_key,"Default Download Directory",path)
  print("Default Download Directory")
  #find_key(offer_save_password,"FormSuggest Passwords","no")
  find_key(ie_key,"iexplore.exe",0)
  print("iexplore.exe")
#file_path = os.path.dirname(os.path.realpath(__file__))
#IE_path=file_path+"\msedgedriver.exe"
#browser=webdriver.Edge(executable_path=IE_path)
def login_Medi(medi_base):
    # Initiate browser
    file_path = os.path.dirname(os.path.realpath(__file__))
    Edge_path=file_path+"\msedgedriver.exe"
    IE_path=file_path+"\IEDriverServer32bit.exe"
    download_path=file_path+"\download"#download to parent another floder

    caps = DesiredCapabilities.INTERNETEXPLORER.copy()
    caps["ignoreProtectedModeSettings"]=True
    caps["EnableNativeEvents"]=True
    caps["nativeEvents"]=True
    caps["unexpectedAlertBehaviour"]="accept"
    caps["disable-popup-blocking"]=True
    caps["requireWindowFocus"]=True
    caps["enablePersistentHover"]=True
    caps["ignoreZoomSetting"]=True
    caps["FORCE_CREATE_PROCESS"]=True
    caps["IE_ENSURE_CLEAN_SESSION"]=True
    caps["IE_SWITCHES"]="-private"
    caps["INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS"]=True
    browser=webdriver.Ie(executable_path=IE_path,capabilities=caps)
    #browser.implicitly_wait(5)
    #early configure for protected mode for all site , and zoom view to 100%
    #
    
    try:
      browser.get("https://dummy.asia-assistance.com/mosai/")
    except NoSuchWindowException:
      Wait(2)
      print("EXCEPTION")
      browser.get("https://dummy.asia-assistance.com/mosai/")


    Wait(3)
    browser.maximize_window()
    Wait(3)
    browser.find_element_by_xpath(("//html//body")).send_keys(Keys.F11)
    Wait(3)
    browser.set_window_size(720, 480)
    Wait(3)
    browser.set_window_position(0,0)




    #new_window = browser.window_handles[1]   
    #browser.switch_to_window(new_window)
    #browser.get("https://stackoverflow.com")
    wait = WebDriverWait(browser, 100)
    #wait.until(EC.presence_of_element_located((By.XPATH, "//a[@class='login-link s-btn btn-topbar-clear py8 js-gps-track']")))
    print("s")
    #wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'listing-item__price')))
    #browser.execute_script("arguments[0].click();", service_notes[0])
    #browser.find_element_by_xpath("//a[@class='login-link s-btn btn-topbar-clear py8 js-gps-track']").click()
    print(IE_path)

    #
    try:
      wait.until(EC.presence_of_element_located((By.XPATH, "//input[@name='userid']")))
      browser.find_element_by_xpath('/html[1]/body[1]/div[2]/form[1]/table[1]/tbody[1]/tr[1]/td[1]/table[1]/tbody[1]/tr[2]/td[2]/input[1]').clear()
      browser.find_element_by_xpath('/html[1]/body[1]/div[2]/form[1]/table[1]/tbody[1]/tr[1]/td[1]/table[1]/tbody[1]/tr[3]/td[2]/input[1]').clear()
    except:
      print("")
    if medi_base.password=="" and medi_base.username=="":
      #simon.lee
      #M@rcpass2
      parser = ConfigParser()
      path = os.path.dirname(os.path.abspath(os.path.dirname(__file__)))
      filename=path+r"\config.ini"
      parser.read(filename)
      if parser.has_section("medi"):
          item = parser.items("medi")
      Password=browser.find_element_by_xpath('/html[1]/body[1]/div[2]/form[1]/table[1]/tbody[1]/tr[1]/td[1]/table[1]/tbody[1]/tr[3]/td[2]/input[1]')
      ActionChains(browser).move_to_element(Password).click().send_keys(item[0][0])
      ID=browser.find_element_by_xpath('/html[1]/body[1]/div[2]/form[1]/table[1]/tbody[1]/tr[1]/td[1]/table[1]/tbody[1]/tr[2]/td[2]/input[1]')
      ActionChains(browser).move_to_element(ID).click().send_keys(item[1][0])
    else:

      browser.find_element_by_xpath("//input[@name='passwd']").send_keys(medi_base.password)
      ID=browser.find_element_by_xpath("//input[@name='userid']")
      ID.send_keys(medi_base.username)
    browser.find_element_by_xpath("//input[@name='submit']").send_keys(Keys.ENTER)
    # Select first hit
    # browser.find_elements_by_class_name('r')[0].click()
    return browser
def wait_to_load(browser,element):
  #wait for element to be show,suggested use in loading page for browser
  wait=WebDriverWait(browser,20)
  wait.until(EC.visibility_of_element_located((By.XPATH, element)))

def short_wait_to_load(browser,element):
  #wait for element to be show,suggested for shorting and manipulating element
  wait=WebDriverWait(browser,3)
  wait.until(EC.visibility_of_element_located((By.XPATH, element)))

def wait_to_click(browser,element):
  #wait for element to be show,suggested use in JAVAscript loading page for browser
  wait=WebDriverWait(browser,20)
  wait.until(EC.element_to_be_clickable((By.XPATH, element)))

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

def fill_date(browser,input):
  Date=input+timedelta(days=1)
  dateList=[]
  for i in get_EventDate():
    temp=list(i)
    dateList.append(temp[0])
  while Date in dateList:
    Date=Date+timedelta(days=1)
  formatted_Date = Date.strftime("%d/%m/%Y")
  browser.find_element_by_xpath("//input[@name='BordDate']").send_keys(formatted_Date)
  
def go_cover_report(browser):
  short_wait_to_load(browser,'/html[1]/body[1]/div[4]/div[3]/table[1]/tbody[1]/tr[1]/td[2]')
  #browser.find_element_by_xpath('/html[1]/body[1]/div[4]/div[3]/table[1]/tbody[1]/tr[1]/td[2]').send_keys(Keys.ENTER)
  #browser.find_element_by_xpath('/html[1]/body[1]/div[10]/div[5]/table[1]/tbody[1]/tr[1]/td[2]').send_keys(Keys.ENTER)
  #browser.find_element_by_xpath('/html[1]/body[1]/div[20]/div[1]/table[1]/tbody[1]/tr[1]/td[1]').send_keys(Keys.ENTER)
  WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, '/html[1]/body[1]/div[4]/div[3]/table[1]/tbody[1]/tr[1]/td[2]'))).click()
  browser.find_element_by_css_selector("body:nth-child(2) div:nth-child(24) > div:nth-child(2)").click()
  #browser.find_element_by_css_selector("body:nth-child(2) div:nth-child(30) > div:nth-child(4)").click()
  ele=browser.find_element_by_css_selector("body:nth-child(2) div:nth-child(30) > div:nth-child(5)").location
  #actions = ActionChains(browser)
  #actions.move_to_element(ele).perform()
  #actions.double_click.perform()
  #actions.context_click()
  ClickOnPosition(ele["x"], ele['y'])
  other=browser.find_element_by_xpath('/html[1]/body[1]/div[20]/div[1]/table[1]/tbody[1]/tr[1]/td[1]').location
  ClickOnPosition(other["x"], other['y'])
  WebDriverWait(browser, 20).until(EC.visibility_of_element_located((By.ID, 'MainiFrame')))
  coor=browser.find_element_by_id("MainiFrame").location
  iframe = browser.find_element_by_id("MainiFrame")

  browser.switch_to.frame(iframe)
  return coor
  #browser.switch_to.frame(browser.find_element_by_tag_name("iframe"))
  

def go_default(browser):
  browser.switch_to.default_content()

def find_client(browser,Client):
  """
  select client/insurer in medi in bdx filling page
  """
  element=browser.find_element_by_xpath("//select[@name='ClientOpt']")
  dropdown = Select(element)
  options = dropdown.options
  i=0
  for index in options:
    #need change to non linear search method
    i=i+1
    #print(index.text)
    if Client in index.get_attribute("value"):
      
      browser.find_element_by_xpath('/html[1]/body[1]/form[1]/table[1]/tbody[1]/tr[1]/td[3]/select[1]/option[{}]'.format(i)).click()
      #ALERT might have 4 choice but same company
      break

def find_insurance(browser,Insurance):
  """
  select client/insurer in medi in bdx filling page
  """
  Wait(2)
  wait_to_click(browser,"/html[1]/body[1]/form[1]/table[1]/tbody[1]/tr[2]/td[3]/select[1]/option[4]")
  element=browser.find_element_by_xpath("//select[@name='Insurance']")
  dropdown = Select(element)
  options = dropdown.options

  i=0
  for index in options:
    #need change to non linear search method
    i=i+1
    if Insurance in index.get_attribute("value"):
      browser.find_element_by_xpath('/html[1]/body[1]/form[1]/table[1]/tbody[1]/tr[2]/td[3]/select[1]/option[{}]'.format(i)).click()
      #ALERT might have 4 choice but same company
      break


def check_manual_web(browser,bdx_type,add_y):
  wait_to_click(browser,'/html[1]/body[1]/form[1]/table[1]/tbody[1]/tr[7]/td[1]/input[1]')
  wait_to_click(browser,'/html[1]/body[1]/form[1]/table[1]/tbody[1]/tr[6]/td[1]/input[1]')
  if "web" in bdx_type.lower():
    web_element=browser.find_element_by_xpath('/html[1]/body[1]/form[1]/table[1]/tbody[1]/tr[7]/td[1]/input[1]')
  elif "manual" in bdx_type.lower():
    web_element=browser.find_element_by_xpath('/html[1]/body[1]/form[1]/table[1]/tbody[1]/tr[6]/td[1]/input[1]')
  element=web_element.location
  size=web_element.size
  element["y"]+=add_y+size['height']/2
  element["x"]+=size['width']/2
  ClickOnPosition(element["x"], element['y'])

def fill_web(browser,input,add_y):
  #last=input- relativedelta(years=2)
  WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html[1]/body[1]/form[1]/table[1]/tbody[1]/tr[7]/td[3]/div[1]/input[1]"))).click()
  last_formatted=input[0]
  today_formatted=input[1]
  #last_formatted= last.strftime("%d/%m/%Y")
  #today_formatted=input.strftime("%d/%m/%Y")

  browser.find_element_by_xpath("/html[1]/body[1]/form[1]/table[1]/tbody[1]/tr[7]/td[3]/div[1]/input[1]").send_keys(last_formatted)
  browser.find_element_by_xpath("/html[1]/body[1]/form[1]/table[1]/tbody[1]/tr[7]/td[3]/div[1]/input[3]").send_keys(today_formatted)
  browser.find_element_by_xpath("/html[1]/body[1]/form[1]/table[1]/tbody[1]/tr[7]/td[3]/div[1]/select[1]/option[3]").click()
  #browser.find_element_by_xpath(  "/html[1]/body[1]/form[1]/input[2]").click()
  web_element=browser.find_element_by_xpath(  "/html[1]/body[1]/form[1]/input[2]")
  element=web_element.location
  size=web_element.size
  element["y"]+=size['height']/2+add_y
  element["x"]+=size['width']/2
  ClickOnPosition(element["x"], element['y'])


def click_print(browser,add_y):
  Wait(2)
  wait_to_click(browser,"//input[@name='PrtRpt']")
  browser.find_element_by_xpath("//input[@name='PrtRpt']").click()
  browser.find_element_by_xpath("//input[@name='PrtRpt']").send_keys(Keys.ENTER)
  web_element=browser.find_element_by_xpath(  "//input[@name='PrtRpt']")
  element=web_element.location
  size=web_element.size
  element["y"]+=size['height']/2+add_y
  element["x"]+=size['width']/2
  ClickOnPosition(element["x"], element['y'])

def fill_manual(browser,input,add_y):
  WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html[1]/body[1]/form[1]/table[1]/tbody[1]/tr[6]/td[3]/div[1]/input[1]"))).click()
  element="/html[1]/body[1]/form[1]/table[1]/tbody[1]/tr[6]/td[3]/div[1]/input[1]"
  browser.find_element_by_xpath(element).send_keys(input)
  #browser.find_element_by_xpath("/html[1]/body[1]/form[1]/input[2]").click()
  web_element=browser.find_element_by_xpath("/html[1]/body[1]/form[1]/input[2]")
  element=web_element.location
  size=web_element.size
  element["y"]+=size['height']/2+add_y
  element["x"]+=size['width']/2
  ClickOnPosition(element["x"], element['y'])

def key_in_serial(case_id_list,browser,add_y):
  wait_to_click(browser,"/html[1]/body[1]/form[1]/table[1]/tbody[1]/tr[1]/td[2]/input[1]")
  for caseid in case_id_list:
    wait_to_click(browser,"/html[1]/body[1]/form[1]/table[1]/tbody[1]/tr[1]/td[2]/input[2]")
    browser.find_element_by_xpath("/html[1]/body[1]/form[1]/table[1]/tbody[1]/tr[1]/td[2]/input[1]").send_keys(caseid)
    browser.find_element_by_xpath("/html[1]/body[1]/form[1]/table[1]/tbody[1]/tr[1]/td[2]/input[2]").click()
    browser.find_element_by_xpath("/html[1]/body[1]/form[1]/table[1]/tbody[1]/tr[1]/td[2]/input[2]").send_keys(Keys.ENTER)
    #web_element=browser.find_element_by_xpath(  "/html[1]/body[1]/form[1]/table[1]/tbody[1]/tr[1]/td[2]/input[2]")
    #element=web_element.location
    #size=web_element.size
    #element["y"]+=size['height']/2+add_y
    #element["x"]+=size['width']/2
    #ClickOnPosition(element["x"], element['y'])

    #DownloadZip.asp?File=C:\inetpub\wwwroot\mosai\Downloads\SPINMAL-0006-201910_HWEESEE.TEH.zip&Name=SPINMAL-0006-201910_HWEESEE.TEH.zip&Excel=SPINMAL-0006-201910_HWEESEE.TEH.xls


def get_bord_no(browser):
  element="//input[@name='BordNo']"
  wait_to_click(browser,element)
  wait_to_load(browser,element)
  bord_no=browser.find_element_by_xpath(element).get_attribute("value")
  return bord_no

def get_file_link(replacement,userid):
  text="https://dummy.asia-assistance.com/mosai/DownloadZip.asp?File=C:\inetpub\wwwroot\mosai\Downloads\SPIRITAERO-0000-000000_user.id.zip&Name=SPIRITAERO-0000-000000_user.id.zip&Excel=SPIRITAERO-0000-000000_user.id.xls"
  link=text.replace("SPIRITAERO-0000-000000",replacement)
  link=link.replace("user.id",userid.upper())
  return link
def close_Medi(browser):
  try:
    browser.close()
    browser.quit()
  except Exception as error:
    print(error)

def get_ele_link(replacement,userid):
  text="//a[contains(text(),'SPINMAL-0007-201910_HWEESEE.TEH.zip')]"
  text.replace("SPINMAL-0007-201910",replacement)
  text.replace("HWEESEE.TEH",userid)
  return text

def download_file(browser,link,ele):

  browser.switch_to.default_content()
  Wait(2)
  iframe = browser.find_element_by_id("MainiFrame")
  browser.switch_to.frame(iframe)
  print(ele)
  try:
    browser.find_element_by_xpath("/html[1]/body[1]/div[1]/p[1]/span[1]/a[1]").click()
    print("xpath")
  except:
    try:
      browser.find_element_by_xpath(ele).click()
      print("ele")
    except:
      browser.get(link)
      print("link")

async def download_and_close(browser,link,ele):
    print("start")
    browser.get(link)
    print("sleep")
    asyncio.sleep(5)
    close_Medi(browser)
    print("1")
def do_actions(browser,link):
  browser.get(link)
def file_check(file,list_f):
  if file in list_f:
    return False
  else:
    return True
def get_client_code(client):
  connector = MSSqlConnector()
  cursor = connector.cursor()
  qry="SELECT clientcode FROM mcl.bencom_client_list where ClientName=%s"
  cursor = connector.cursor()
  cursor.execute(qry,client)
  code = cursor.fetchall()
  return code[0][0]
def get_all_File(file_path,file_type):
  if file_type=="xls":
    list_of_files = glob.glob(file_path+"\*.xls") # * means all if need specific format then *.csv
  elif file_type=="pdf":
    list_of_files = glob.glob(file_path+"\*.pdf") # * means all if need specific format then *.csv
  elif file_type=="zip":
    list_of_files = glob.glob(file_path+"\*.zip") # * means all if need specific format then *.csv
  return list_of_files
#print(get_all_File(r"\\dtisvr2\mediclinic\New\Admission_354\New","zip"))
#download_path=r"\\dtisvr2\mediclinic\New\Admission_354\New"
#filename="SPIRITAERO-0022-201911"+"_"+"HWEESEE.TEH"
#extension=".zip"
#print(download_path+filename+extension)
def Medi_Mos(download_path,medi_base,bdx_type,client,insurance,input):
  setup_env(download_path)
  audit_log("Login to MOS", "Completed...", medi_base)
  try:
    try:
      browser=login_Medi(medi_base)
      coor=go_cover_report(browser)
      code=get_client_code(client)
      audit_log("Login sucessfully", "Completed...", medi_base)
    except:
      close_Medi(browser)
      browser=login_Medi(medi_base)
      coor=go_cover_report(browser)
      code=get_client_code(client)
      audit_log("Login sucessfully", "Completed...", medi_base)
    find_client(browser,code)
    find_insurance(browser,insurance)
    fill_date(browser,get_Date())
    if bdx_type=="web":
      check_manual_web(browser,"web",coor["y"])
      fill_web(browser,input,coor["y"])#get_Date()
      bord_no=get_bord_no(browser)

      #click_print(browser)
    elif bdx_type=="manual":
      check_manual_web(browser,"manual",coor["y"])
      fill_manual(browser,input[0],coor["y"])
      bord_no=get_bord_no(browser)
      key_in_serial(input,browser,coor["y"])
    userid=medi_base.username
    link=get_file_link(bord_no,userid)
    ele=get_ele_link(bord_no,userid)
    print(link)
    try:
      click_print(browser,coor["y"])
      audit_log("Ready to generate bdx file", "Completed...", medi_base)
    except:
      print("E")
    Wait(10)
    filename=download_path+"\\"+bord_no+"_"+userid.upper()+".zip"
    file_list =get_all_File(download_path,"zip")
    file_Checking=False

    if filename in file_list:
      file_Checking = True
    #WebDriverWait(browser, 5).until(EC.alert_is_present())
    #alert = browser.switch_to_alert()
    ##Retrieve the message on the Alert window
    #message=alert.text
    #print ("Alert shows following message: "+ message )
    ##use the accept() method to accept the alert
    #alert.accept()
    ## Or Dismiss the Alert using
    #alert.dismiss()
    #action_process = Process(target=do_actions(browser,link))
    wait = WebDriverWait(browser, 5)
    ClickOnPosition(771, 687)
    print(filename,"expected")

    if file_check(filename,file_list):
      try:
        browser.set_page_load_timeout(10)
        browser.get(link)
        #action_process = Process(target=do_actions(browser,link))
        ClickOnPosition(771, 687)
        WebDriverWait(browser, 20).until(EC.visibility_of_element_located((By.ID, 'MainiFrame')))
        print("ss")
        Wait(2)
        iframe = browser.find_element_by_id("MainiFrame")
        browser.switch_to.frame(iframe)
        print(len(get_all_File(download_path,"zip")))
        ClickOnPosition(771, 687)
        ClickOnPosition(771, 687)
        ClickOnPosition(771, 687)
        ClickOnPosition(771, 687)
      except:
        pass
    if file_check(filename,file_list):
      try:
        browser.find_element_by_xpath("//a[@class='File']").click()
        print("ss")
        Wait(2)
        ClickOnPosition(771, 687)
      except:
        pass
    if file_check(filename,file_list):
      try:
        browser.set_page_load_timeout(5)
        iframe = browser.find_element_by_id("MainiFrame")
        browser.switch_to.frame(iframe)
        browser.find_element_by_xpath("//a[@class='File']").click()
        print("ss")
        Wait(2)
        ClickOnPosition(771, 687)
        print(len(get_all_File(download_path,"zip")))
      except:
        pass
      if file_check(filename,file_list):
        try:
          browser.set_page_load_timeout(5)
          iframe = browser.find_element_by_id("MainiFrame")
          browser.switch_to.frame(iframe)
          browser.find_element_by_xpath("//a[@class='File']").send_keys(Keys.ENTER)
          print("ss")
          Wait(2)
          ClickOnPosition(771, 687)
          print(len(get_all_File(download_path,"zip")))
        except:
          pass
      if file_check(filename,file_list):
        try:
          link_location=browser.find_element_by_xpath("//a[@class='File']").location
          link_size=browser.find_element_by_xpath("//a[@class='File']").size
          ClickOnPosition(link_location["x"]+link_size["height"]/2,link_location["y"]+link_size['width']/2+coor["y"])
          Wait(2)
          ClickOnPosition(771, 687)
          Print("final")
          print(len(get_all_File(download_path,"zip")))
        except:
          pass

        if file_check(filename,file_list):
          audit_log("Downloaded BDX file into server", "Completed...", medi_base)
        else:
          audit_log("Failed-Download BDX file ", "Completed...", medi_base)
        print(len(get_all_File(download_path,"zip")))    
        close_Medi(browser)
        return bdx_type
  except:
    close_Medi(browser)
