from Automagica import *
import json


def gen_preview(jsondf):
  try:
    #Execute RPA on all Case ID in dataframe
    browser=login_Marc()

    #Iterate all Case ID, navigate to search page in March and search case ID
    for i in jsondf.index:
      navigate_page(browser,'inpatient','cases','inpatient_cases_search')
      search_Marc(browser, jsondf['ID'][i])
      jsondf.loc[i,"Patient"]=browser.find_element_by_xpath('//*[@id="mmTab:j_idt143"]').get_attribute('value')
      jsondf.loc[i,"Hospital"]=browser.find_element_by_xpath('//*[@id="caseTab:hospitalname"]').get_attribute('value')
      jsondf.loc[i,"Pol. Type"]=browser.find_element_by_xpath('//*[@id="mmTab:j_idt198"]').get_attribute('value')
      jsondf.loc[i,"Client"]=browser.find_element_by_xpath('//*[@id="mmTab:j_idt222"]').get_attribute('value')
      jsondf.loc[i,"Bill. To"]=browser.find_element_by_xpath('//*[@id="mmTab:j_idt224"]').get_attribute('value')
      jsondf.loc[i,"Attn. To"]=""#retrieve from postgres db
      jsondf.loc[i,"Address"]=""#retrieve from postgres db

    close_Marc(browser)

  except Exception as error:
    print(error)

# --------------------------
# This script will have the following outputs:
# - .json
# Associated with the following scripts:
# - read_json.py
# - Automagica.py
