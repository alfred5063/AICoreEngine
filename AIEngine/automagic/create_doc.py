from automagica import *

def create_doc(browser, attn_to, doc_template, attn_Type):

    #Edit document of client
    browser.find_element_by_xpath('//*[@id="gotoUpdatePageBtn"]').click()
    Wait(0.5)
    browser.find_element_by_xpath('//*[@id="caseTab"]/ul/li[6]/a').click()
    Wait(0.5)

    #Select option in dropdown menu
    Select(browser.find_element_by_xpath('//*[@id="caseTab:docTemplateId"]')).select_by_visible_text(doc_template)
    Wait(1)
    Select(browser.find_element_by_xpath('//*[@id="caseTab:attnType"]')).select_by_visible_text(attn_Type)
    Wait(1)
    browser.find_element_by_xpath('//*[@id="caseTab:attnTo"]').send_keys(attn_to)

    #Select "Send as"
    browser.find_element_by_xpath('//*[@id="caseTab:sendAs:0"]').click()
    Wait(0.5)

    #Create Document
    browser.find_element_by_xpath('//*[@id="caseTab:tabDocCreateDocButton"]').click()
    

#Edit the document in word-viewer
def edit_doc(browser, bill_to, HNS, CLN, amount, date, time):

    #Set date and time
    Wait(20)
    Tab()
    Type(time, interval_seconds=0.1)
    PressHotkey("shift","tab")
    Type(date, interval_seconds=0.1)

    #Close date picker menu
    ClickOnPosition(850, 221)

    #Edit 'bill to' field
    ClickOnPosition(285, 569)
    Type(bill_to, interval_seconds=0.1)

    #Edit 3rd and 4th column of the document and set it centred
    ClickOnPosition(798, 668)
    Type(HNS, interval_seconds=0.1)
    ClickOnPosition(573, 268)#click on "centre paragraph"
    ClickOnPosition(1182, 667)
    Type(CLN, interval_seconds=0.1)
    ClickOnPosition(573, 268)

    #Scroll down the document
    MoveToPosition(1639, 610)
    DragToPosition(1657, 992)

    #Edit the total amount field
    ClickOnPosition(1608, 615)
    Backspace()
    Backspace()
    Backspace()
    Backspace()
    Type(str(amount), interval_seconds=0.2)
    #browser.find_element_by_xpath('//*[@id="caseTab:editDocSaveButton"]').click()
    #browser.close()
