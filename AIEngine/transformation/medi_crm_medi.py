from __future__ import unicode_literals
from xlwt import Workbook
import io
import pyautogui
import os
from datetime import date,datetime,timedelta
import datetime as dt
from dateutil.relativedelta import relativedelta
from automagica import *
from selenium import webdriver
import enum
import json
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException,TimeoutException,WebDriverException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
from connector.connector import MSSqlConnector
import traceback
from utils.logging import logging
from selenium.webdriver.common.action_chains import ActionChains
from utils.audit_trail import audit_log
from pathlib import Path

current_path = os.path.dirname(os.path.abspath(__file__))
dti_path = str(Path(current_path).parent)

def wait_to_load(browser,element):
  #wait for element to be show,suggested use in loading page for browser
  wait=WebDriverWait(browser,20)
  wait.until(EC.visibility_of_element_located((By.XPATH, element)))
def wait_to_load_short(browser,element):
  #wait for element to be show,suggested use in loading page for browser
  wait=WebDriverWait(browser,20)
  wait.until(EC.visibility_of_element_located((By.XPATH, element)))
def wait_for_overlag(browser):
  #wait for the webpage to load the overlay, overlay is the delay that make mouse cannot click
  wait_overlag=WebDriverWait(browser,200)
  wait_overlag.until(EC.invisibility_of_element_located((By.ID,"navBarOverlay")))
  wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,"//div[@id='reportViewer_AsyncWait_Wait']")))
  
  wait_overlag.until(EC.invisibility_of_element_located((By.CLASS_NAME,"navBarOverlay")))
  wait_overlag.until(EC.invisibility_of_element_located((By.ID,"/html[1]/body[1]/div[4]/div[3]")))
  Wait(1)
def login_crm(user,password,download_path, profile=None):
    file_path = os.path.dirname(os.path.realpath(__file__))
    print(download_path)
    #Chrome_path=file_path+"\chromedriver.exe"
    Chrome_path=dti_path+r'/selenium_interface/chromedriver.exe'
    chrome_profile = webdriver.ChromeOptions()
 
    #chrome_profile.add_argument('--no-sandbox')
    profile = {"download.default_directory": download_path,
               "download.prompt_for_download": False,
               "download.directory_upgrade": True,
               "plugins.plugins_disabled": ["Chrome PDF Viewer"]}
    chrome_profile.add_experimental_option("prefs", profile)
    chrome_profile.add_argument("--disable-extensions")

    browser = webdriver.Chrome(executable_path=Chrome_path,chrome_options=chrome_profile)
    browser.get("https://{}:{}@crm.asia-assistance.com:5443".format(user,password))
    
    browser.maximize_window()
    return browser

def cancel_login_near(browser,medi_base):
  try:
    Wait(2)
    wait_to_load(browser,"//iframe[@id='InlineDialog_Iframe']")
    Wait(2)
    iframe = browser.find_element_by_xpath("//iframe[@id='InlineDialog_Iframe']")
    browser.switch_to.frame(iframe)
    wait_to_load(browser,"//a[@id='buttonClose']")
    browser.find_element_by_xpath("//a[@id='buttonClose']").click()
    browser.switch_to.default_content()
  except:
    logging("MEDI-process mediclinic error in CRM",traceback.format_exc(),medi_base)
    pass
def go_bdx_page(browser):
  wait_for_overlag(browser)
  wait_to_load(browser,"/html[1]/body[1]/div[4]/div[2]/div[1]/span[4]/a[1]")
  browser.find_element_by_xpath("/html[1]/body[1]/div[4]/div[2]/div[1]/span[4]/a[1]").click()
  wait_to_load(browser,"//a[@id='borderaux']")
  browser.find_element_by_xpath("//a[@id='borderaux']").click()
def click_new_bdx(browser):
  try:
    #browser.find_element_by_id("navBarOverlay").click()
    wait_for_overlag(browser)
    #wait_to_load(browser,"/html[1]/body[1]/div[4]/div[3]")
    browser.find_element_by_xpath("/html[1]/body[1]/div[4]/div[3]").click()
    #wait_to_load(browser,'//*[@id="aste_borderaux|NoRelationship|HomePageGrid|Mscrm.HomepageGrid.aste_borderaux.NewRecord"]/span/a')
    browser.find_element_by_xpath('//*[@id="aste_borderaux|NoRelationship|HomePageGrid|Mscrm.HomepageGrid.aste_borderaux.NewRecord"]/span/a').click()
  except WebDriverException:
    Wait(2)
    print("another route")
    wait_to_load(browser,'//*[@id="aste_borderaux|NoRelationship|HomePageGrid|Mscrm.HomepageGrid.aste_borderaux.NewRecord"]/span/a')
    browser.find_element_by_xpath('//*[@id="aste_borderaux|NoRelationship|HomePageGrid|Mscrm.HomepageGrid.aste_borderaux.NewRecord"]/span/a').click()
def select_aso(browser):
  wait_for_overlag(browser)
  wait_to_load(browser,"//iframe[@id='contentIFrame1']")
  iframe = browser.find_element_by_xpath("//iframe[@id='contentIFrame1']")
  browser.switch_to.frame(iframe)
  element2="//option[contains(text(),'ASO')]"
  wait_to_load(browser,element2)
  browser.find_element_by_xpath(element2).click()
  browser.switch_to.default_content()
def select_insurance(browser):

  wait_for_overlag(browser)
  wait_to_load(browser,"//iframe[@id='contentIFrame1']")
  iframe = browser.find_element_by_xpath("//iframe[@id='contentIFrame1']")
  browser.switch_to.frame(iframe)
  element2="//option[contains(text(),'Insurance')]"
  wait_to_load(browser,element2)
  browser.find_element_by_xpath(element2).click()
  browser.switch_to.default_content()
def search_client(browser):
  wait_to_load(browser,"//iframe[@id='contentIFrame1']")
  iframe = browser.find_element_by_xpath("//iframe[@id='contentIFrame1']")
  browser.switch_to.frame(iframe)
  wait_to_load(browser,'//*[@id="aste_client"]/div[1]')
  browser.find_element_by_xpath('//*[@id="aste_client"]/div[1]').click()
  wait_to_load(browser,"//img[@id='aste_client_i']")
  browser.find_element_by_xpath("//img[@id='aste_client_i']").click()
  browser.switch_to.default_content()
def search_client_more(browser):
  iframe = browser.find_element_by_xpath("//iframe[@id='contentIFrame1']")
  browser.switch_to.frame(iframe)
  try:
    element=browser.find_element_by_xpath("//*[@class='ms-crm-IL-MenuItem-Lookupmore ms-crm-IL-MenuItem-Lookupmore-Rest']").click()
  except:
    browser.find_element_by_xpath("/html[1]/body[1]/div[10]/div[1]/ul[1]/li[11]").click()
  browser.switch_to.default_content()
def get_client(browser,client_name):
  iframe = browser.find_element_by_xpath("//iframe[@id='InlineDialog_Iframe']")
  browser.switch_to.frame(iframe)
  try:
    wait_to_load(browser,"//label[@id='crmGrid_findHintText']")
    browser.find_element_by_xpath("//label[@id='crmGrid_findHintText']").click()
  except:
    print("cannot find label[@id='crmGrid_findHintText'")
    pass

  browser.find_element_by_xpath("//input[@id='crmGrid_findCriteria']").send_keys(client_name)
  browser.find_element_by_xpath("//img[@id='crmGrid_findCriteriaImg']").click()
  try:
    wait_to_load(browser,"//tr[@class='ms-crm-List-SelectedRow']")
  except:
    wait_to_load(browser,"//tr[@class='ms-crm-List-SelectedRow']")
  browser.find_element_by_xpath("//button[@id='butBegin']").click()
  browser.switch_to.default_content()

def search_insurance(browser):
  iframe = browser.find_element_by_xpath("//iframe[@id='contentIFrame1']")
  browser.switch_to.frame(iframe)
  wait_to_load(browser,"//div[@id='aste_insurancecompany']")
  browser.find_element_by_xpath("//div[@id='aste_insurancecompany']").click()
  browser.find_element_by_xpath("//img[@id='aste_insurancecompany_i']").click()
  browser.switch_to.default_content()

def search_insurance_more(browser):

  iframe = browser.find_element_by_xpath("//iframe[@id='contentIFrame1']")
  browser.switch_to.frame(iframe)
  try:
    wait_to_load(browser,"//a[@class='ms-crm-IL-MenuItem-Anchor-Lookupmore ms-crm-IL-MenuItem-Anchor-Lookupmore-Rest']")
    element=browser.find_element_by_xpath("//a[@class='ms-crm-IL-MenuItem-Anchor-Lookupmore ms-crm-IL-MenuItem-Anchor-Lookupmore-Rest']").click()
  except:
    print("ms-crm-IL-MenuItem-Lookupmore ms-crm-IL-MenuItem-Lookupmore-Rest", "NOT FOUND")
    wait_to_load(browser,"/html[1]/body[1]/div[10]/div[1]/ul[1]/li[5]")
    browser.find_element_by_xpath("/html[1]/body[1]/div[10]/div[1]/ul[1]/li[5]").click()
  browser.switch_to.default_content()

def get_insurance(browser,insurance_name):
  iframe = browser.find_element_by_xpath("//iframe[@id='InlineDialog_Iframe']")
  browser.switch_to.frame(iframe)
  try:
    wait_to_load(browser,"//label[@id='crmGrid_findHintText']")
    browser.find_element_by_xpath("//label[@id='crmGrid_findHintText']").click()
  except:
    print("cannot find label[@id='crmGrid_findHintText")
    pass

  browser.find_element_by_xpath("//input[@id='crmGrid_findCriteria']").send_keys(insurance_name)
  browser.find_element_by_xpath("//img[@id='crmGrid_findCriteriaImg']").click()
  try:
    wait_to_load(browser,"//tr[@class='ms-crm-List-SelectedRow']")
  except:
    wait_to_load(browser,"//tr[@class='ms-crm-List-SelectedRow']")
  browser.find_element_by_xpath("//button[@id='butBegin']").click()
  browser.switch_to.default_content()


def key_in_insurance(browser,insurance_name):
  search_insurance(browser)
  search_insurance_more(browser)
  get_insurance(browser,insurance_name)
def key_in_client(browser,client_name):
  search_client(browser)
  search_client_more(browser)
  get_client(browser,client_name)

def click_clinic_portal(browser):
  #initial is NO
  wait_for_overlag(browser)
  wait_to_load(browser,"//iframe[@id='contentIFrame1']")
  iframe = browser.find_element_by_xpath("//iframe[@id='contentIFrame1']")
  browser.switch_to.frame(iframe)
  browser.find_element_by_xpath('//*[@id="aste_portal"]').click()
  browser.switch_to.default_content()
def key_in_start_date(browser,start_date):
  iframe = browser.find_element_by_xpath("//iframe[@id='contentIFrame1']")
  browser.switch_to.frame(iframe)
  wait_to_load(browser,"//div[@id='aste_startdate']")
  browser.find_element_by_xpath("//div[@id='aste_startdate']").click()
  browser.find_element_by_xpath("//table[@id='aste_startdate_i']//input[@id='DateInput']").send_keys(start_date)
  browser.switch_to.default_content()

def key_in_end_date(browser,end_date):
  iframe = browser.find_element_by_xpath("//iframe[@id='contentIFrame1']")
  browser.switch_to.frame(iframe)
  wait_to_load(browser,"//div[@id='aste_enddate']")
  browser.find_element_by_xpath("//div[@id='aste_enddate']").click()
  browser.find_element_by_xpath("//table[@id='aste_enddate_i']//input[@id='DateInput']").send_keys(end_date)
  browser.switch_to.default_content()

def key_in_policy_year(browser,policy_year):
  iframe = browser.find_element_by_xpath("//iframe[@id='contentIFrame1']")
  browser.switch_to.frame(iframe)
  wait_to_load(browser,"//div[@id='aste_yearsofpolicy']")
  browser.find_element_by_xpath("//div[@id='aste_yearsofpolicy']").click()
  wait_to_load(browser,"//input[@id='aste_yearsofpolicy_i']")
  browser.find_element_by_xpath("//input[@id='aste_yearsofpolicy_i']").send_keys(policy_year)
  browser.switch_to.default_content()

def key_in_dis_no(browser,dis_no):
  iframe = browser.find_element_by_xpath("//iframe[@id='contentIFrame1']")
  browser.switch_to.frame(iframe)
  wait_to_load(browser,"//div[@id='aste_disbursementno']")
  browser.find_element_by_xpath("//div[@id='aste_disbursementno']").click()
  wait_to_load(browser,"//input[@id='aste_disbursementno_i']")
  browser.find_element_by_xpath("//input[@id='aste_disbursementno_i']").send_keys(dis_no)
  browser.switch_to.default_content()

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

def fill_date(input):
  Date=input+timedelta(days=1)
  dateList=[]
  for i in get_EventDate():
    temp=list(i)
    dateList.append(temp[0])
  while Date in dateList:
    Date=Date+timedelta(days=1)
  formatted_Date = Date.strftime("%d/%m/%Y")
  return formatted_Date

def key_in_bdx_date(browser,bdx_date):
  iframe = browser.find_element_by_xpath("//iframe[@id='contentIFrame1']")
  browser.switch_to.frame(iframe)
  wait_to_load(browser,"//div[@id='aste_bordereauxdate']")
  browser.find_element_by_xpath("//div[@id='aste_bordereauxdate']").click()
  browser.find_element_by_xpath("//table[@id='aste_bordereauxdate_i']//input[@id='DateInput']").send_keys(bdx_date)
  browser.switch_to.default_content()

def select_reimbursement(browser):
  iframe = browser.find_element_by_xpath("//iframe[@id='contentIFrame1']")
  browser.switch_to.frame(iframe)
  wait_to_load(browser,"//div[@id='aste_claimtype']")
  browser.find_element_by_xpath("//div[@id='aste_claimtype']").click()
  browser.find_element_by_xpath("//option[contains(text(),'Reimbursement')]").click()
  browser.switch_to.default_content()

def select_cashless(browser):
  iframe = browser.find_element_by_xpath("//iframe[@id='contentIFrame1']")
  browser.switch_to.frame(iframe)
  wait_to_load(browser,"//div[@id='aste_claimtype']")
  browser.find_element_by_xpath("//div[@id='aste_claimtype']").click()
  browser.find_element_by_xpath("//option[contains(text(),'Cashless')]").click()
  browser.switch_to.default_content()

def click_save(browser,WEB_MANUAL):
  try:
    browser.find_element_by_xpath("//li[@id='aste_borderaux|NoRelationship|Form|Mscrm.Form.aste_borderaux.Save']").click()
    print("SAVED")
    if aso_insurance=="WEB":
      select_reimbursement(browser)
      
    else:
      select_cashless(browser)
  except:
    print("required code to edit")
    Wait(2)

def key_in_clinical_svr(browser,text):
  iframe = browser.find_element_by_xpath("//iframe[@id='contentIFrame1']")
  browser.switch_to.frame(iframe)
  wait_to_load(browser,"//div[@id='aste_clinicalservice']")
  browser.find_element_by_xpath("//div[@id='aste_clinicalservice']").click()
  wait_to_load(browser,"//input[@id='aste_clinicalservice_i']")
  browser.find_element_by_xpath("//input[@id='aste_clinicalservice_i']").send_keys(text)
  browser.switch_to.default_content()

def click_done_choose_cs(browser):
  #initial is NO
  wait_for_overlag(browser)
  wait_to_load(browser,"//iframe[@id='contentIFrame1']")
  iframe = browser.find_element_by_xpath("//iframe[@id='contentIFrame1']")
  browser.switch_to.frame(iframe)
  browser.find_element_by_xpath("//div[@id='aste_donechoosecs']").click()
  browser.switch_to.default_content()


def wait_to_click(browser,element):
  #wait for element to be show,suggested use in JAVAscript loading page for browser
  wait=WebDriverWait(browser,20)
  wait.until(EC.element_to_be_clickable((By.XPATH, element)))


def key_in_case_id(browser,id_list):
  wait_for_overlag(browser)
  wait_to_load(browser,"//iframe[@id='contentIFrame1']")
  iframe = browser.find_element_by_xpath("//iframe[@id='contentIFrame1']")
  browser.switch_to.frame(iframe)
  wait_to_load(browser,"//iframe[@id='WebResource_borderaux']")
  iframe = browser.find_element_by_xpath("//iframe[@id='WebResource_borderaux']")
  browser.switch_to.frame(iframe)
  for id in id_list:
    wait_to_click(browser,"//input[@id='searchTxt']")
    browser.find_element_by_xpath("//input[@id='searchTxt']").clear()
    browser.find_element_by_xpath("//input[@id='searchTxt']").send_keys(id)
    browser.find_element_by_xpath("//button[@id='searchBtn']").click()

  browser.switch_to.default_content()

  browser.switch_to.default_content()

class function_type(enum.Enum): 
    web = "web"
    manual = "manual"

class insurance_type(enum.Enum): 
    aso = "aso"
    insurance = "insurance"

def get_bord_number(browser):
  wait_for_overlag(browser)
  wait_to_load(browser,"//iframe[@id='contentIFrame1']")
  iframe = browser.find_element_by_xpath("//iframe[@id='contentIFrame1']")
  browser.switch_to.frame(iframe)
  bord_no=browser.find_element_by_xpath("//h1[@class='ms-crm-TextAutoEllipsis']").get_attribute("title")
  browser.switch_to.default_content()
  return bord_no

def click_run_workflow(browser):
  try:
    element="/html[1]/body[1]/div[5]/div[2]/div[1]/ul[1]/li[7]/span[1]/a[1]/span[1]"
    wait_to_load(browser,element)
    browser.find_element_by_xpath(element).click()
    wait_to_load(browser,element)
  except:
    print("another route")
    element="//li[@id='aste_borderaux|NoRelationship|Form|Mscrm.Form.aste_borderaux.RunWorkflow']//a[@class='ms-crm-Menu-Label']"
    wait_to_load(browser,element)
    browser.find_element_by_xpath("//span[contains(text(),'Run Workflow')]").click()
    #browser.find_element_by_xpath(element).click()
    wait_to_load(browser,element)

def click_calculate_bordereaux(browser):
  wait_to_load(browser,"//iframe[@id='InlineDialog_Iframe']")
  iframe = browser.find_element_by_xpath("//iframe[@id='InlineDialog_Iframe']")
  browser.switch_to.frame(iframe)
  try:
    wait_to_load(browser,"//label[@id='crmGrid_findHintText']")
    browser.find_element_by_xpath("//label[@id='crmGrid_findHintText']").click()
  except:
    print("cannot find label[@id='crmGrid_findHintText'")
    pass

  browser.find_element_by_xpath("//input[@id='crmGrid_findCriteria']").send_keys("calculate bordereaux")
  browser.find_element_by_xpath("//img[@id='crmGrid_findCriteriaImg']").click()
  try:
    wait_to_load(browser,"//tr[@class='ms-crm-List-SelectedRow']")
  except:
    wait_to_load(browser,"//tr[@class='ms-crm-List-SelectedRow']")
  browser.find_element_by_xpath("//button[@id='butBegin']").click()
  browser.switch_to.default_content()



def click_confirmation_window_ok(browser):
  error=0
  main_window_handle = None
  while not main_window_handle:
      main_window_handle = browser.current_window_handle
  signin_window_handle = None
  while not signin_window_handle:
      error+=1
      for handle in browser.window_handles:
          error+=1
          if handle != main_window_handle:
              signin_window_handle = handle
              break
          if error == 100:
              break
  browser.switch_to.window(signin_window_handle)
  wait_to_load(browser,"//button[@id='butBegin']")
  browser.find_element_by_xpath("//button[@id='butBegin']").click()
  browser.switch_to.window(main_window_handle)
  
def generate_bdx(browser,input_insurance_type):
  element_setting="//img[@class='ms-crm-moreCommand-image']"
  wait_to_load(browser,element_setting)
  browser.find_element_by_xpath(element_setting).click()
  element_run_report="//span[contains(text(),'Run Report')]"
  wait_to_load(browser,element_run_report)
  browser.find_element_by_xpath(element_run_report).click()
  if input_insurance_type.upper()=="ASO":
    element_report="//span[contains(text(),'Borderaux Listing ASO')]"
  elif input_insurance_type.upper()=="INSURANCE":
    element_report="//span[contains(text(),'Borderaux Listing insurance')]"
  wait_to_load(browser,element_report)
  browser.find_element_by_xpath(element_report).click()

def search_bdx_new_window(browser,bord_num,medi_base):
  error=0
  main_window_handle = None
  while not main_window_handle:
      main_window_handle = browser.current_window_handle
  signin_window_handle = None
  while not signin_window_handle:
      error+=1
      for handle in browser.window_handles:
          error+=1
          if handle != main_window_handle:
              signin_window_handle = handle
              break
          if error == 100:
              break
  browser.switch_to.window(signin_window_handle)
  wait_to_load(browser,"//iframe[@id='resultFrame']")
  iframe = browser.find_element_by_xpath("//iframe[@id='resultFrame']")
  browser.switch_to.frame(iframe)
  element_input="/html[1]/body[1]/div[1]/form[1]/span[1]/div[1]/table[1]/tbody[1]/tr[2]/td[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/table[1]/tbody[1]/tr[1]/td[2]/div[1]/input[1]"
  wait_to_load(browser,element_input)
  browser.find_element_by_xpath(element_input).send_keys(bord_num)

  element_view="//input[@id='reportViewer_ctl04_ctl00']"
  wait_to_load(browser,element_view)
  browser.find_element_by_xpath(element_view).click()

  wait_for_overlag(browser)
  element_save="//img[@id='reportViewer_ctl05_ctl04_ctl00_ButtonImg']"
  wait_to_load(browser,element_save)
  browser.find_element_by_xpath(element_save).click()
  try:

    element_excel="//a[contains(text(),'Excel')]"
    wait_to_load(browser,element_excel)
  
    browser.find_element_by_xpath(element_excel).click()
    Wait(2)
  except:
    audit_log("Cannot download bordreaux,Access error", "Completed...", medi_base)
  browser.close()
  browser.switch_to.window(main_window_handle)

def close_Crm(browser):
  try:
    browser.close()
  except Exception as error:
    print(error)
  try:
    browser.quit()
  except Exception as error:
    print(error)
def check_exists_by_xpath(browser,xpath):
    try:
        browser.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True

def refresh_until_web(browser,finish_time):
  frame_element=''
  try:
    browser.find_element_by_xpath("//iframe[@id='contentIFrame0']")
    frame_element="//iframe[@id='contentIFrame0']"
  except:
    frame_element="//iframe[@id='contentIFrame1']"
  print(frame_element)
  first_claim_element="/html[1]/body[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[4]/div[2]/div[1]/div[3]/table[1]/tbody[1]/tr[2]/td[1]/div[1]/span[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]"
  table="//div[@id='ASO_Claim_divDataArea']//table[@id='gridBodyTable']"
  wait_to_load(browser,frame_element)
  iframe = browser.find_element_by_xpath(frame_element)
  browser.switch_to.frame(iframe)
  while check_exists_by_xpath(browser,first_claim_element)==False and finish_time>datetime.now():
    wait_for_overlag(browser)
    print("WAIT")
    browser.switch_to.default_content()
    wait_to_load(browser,frame_element)
    iframe = browser.find_element_by_xpath(frame_element)
    browser.switch_to.frame(iframe)
    element_table=browser.find_element_by_xpath(table)
    ActionChains(browser).context_click(element_table).perform()
    ActionChains(browser).send_keys(Keys.ARROW_DOWN).perform()
    ActionChains(browser).send_keys(Keys.ARROW_DOWN).perform()
    ActionChains(browser).send_keys(Keys.ARROW_DOWN).perform()
    ActionChains(browser).send_keys(Keys.ARROW_DOWN).perform()
    ActionChains(browser).send_keys(Keys.ARROW_DOWN).perform()
    ActionChains(browser).send_keys(Keys.ARROW_DOWN).perform()
    ActionChains(browser).send_keys(Keys.ENTER).perform()
    browser.switch_to.default_content()
#print((function_type("web").value))
#print(insurance_type(insurance_type.lower()=="aso"))
def Medi_Crm(download_new,medi_base,bdx_type,Client,Insurance,input,option):
  print('Medi_CRM progress')
  user=medi_base.username
  password=medi_base.password
  client_name=Client
  insurance_name=Insurance
  id_list=input
  web_type=bdx_type
  input_insurance_type=option
  policy_year="2019"
  disbursement_no="MCL-"
  
  try:
    
    browser=login_crm(user,password,download_new)
    cancel_login_near(browser,medi_base)
    audit_log("Login successfuly", "Completed...", medi_base)
    go_bdx_page(browser)
    click_new_bdx(browser)
    #different betweeen aso and insurance
    if input_insurance_type.upper()=="ASO":
      select_aso(browser)
      key_in_client(browser,client_name)
    else :
      select_insurance(browser)
      key_in_client(browser,client_name)
      key_in_insurance(browser,insurance_name)
    audit_log("client and insurance in CRM is found", "Completed...", medi_base)
    clinical_svr="GP/RMO,DENTAL,MAT,CHRONIC"
    #different betweeen manual and web
    if web_type=="MANUAL":
      select_reimbursement(browser)
      audit_log("Reimbursement for manual", "Completed...", medi_base)
    elif web_type=="WEB":
      select_cashless(browser)
      click_clinic_portal(browser)
      key_in_start_date(browser,id_list[0])
      key_in_end_date(browser,id_list[1])
      audit_log("Cashless for web", "Completed...", medi_base)

    key_in_bdx_date(browser,fill_date(get_Date()))
    click_save(browser,web_type)
    bord_number=get_bord_number(browser)
    print(bord_number)
    key_in_policy_year(browser,policy_year)
    key_in_clinical_svr(browser,clinical_svr)
    key_in_dis_no(browser,disbursement_no)
    click_done_choose_cs(browser)
    audit_log("case id is ready to key in in CRM", "Completed...", medi_base)
    if web_type=="MANUAL":
      key_in_case_id(browser,id_list)
      audit_log("Excel file for Manual is downloaded", "Completed...", medi_base)
    elif web_type=="WEB":
      click_save(browser,web_type)
      click_save(browser,web_type)
      click_save(browser,web_type)
      click_save(browser,web_type)
      audit_log("Waiting for WEB Excel generated", "Completed...", medi_base)
      finish_time = datetime.now() + dt.timedelta(minutes=6)
      refresh_until_web(browser,finish_time)
      audit_log("Excel file for WEB is downloaded", "Completed...", medi_base)

    try:
      current=browser.find_element_by_xpath("//div[@id='crmContentPanel']").get_attribute("currentcontentid")
      print(current,"current2")
      click_run_workflow(browser)
    except:
      browser.switch_to.default_content()
      click_run_workflow(browser)
    audit_log("running workflow in CRM to generate Bordeaux", "Completed...", medi_base)
    click_calculate_bordereaux(browser)
    click_confirmation_window_ok(browser)
    generate_bdx(browser,input_insurance_type)
    search_bdx_new_window(browser,bord_number,medi_base)
    #close_Crm(browser)
    audit_log("Excel file downloaded sucessfully", "Completed...", medi_base)
    audit_log("google chrome for CRM is closed", "Completed...", medi_base)
    print("COMPLETED")
  except:
    logging("MEDI-process mediclinic error in CRM",traceback.format_exc(),medi_base)
    #close_Crm(browser)
    audit_log("google chrome for CRM is closed", "Completed...", medi_base)



