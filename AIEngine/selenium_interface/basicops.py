#!/usr/bin/env python
#
# Created by Xixuan on 29/08/2018
#
# The file contains
#   1. MARC RPA engine
#

import os
import time
import numpy as np
import pandas as pd
import datetime

from numbers import Number

from selenium.common import exceptions as se

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as ec
import ctypes

from selenium_interface.logger import Logger
from selenium_interface.settings import *
from utils.logging import logging
from utils.audit_trail import audit_log
# TODO: Collapse data pickers after inputting
# TODO: Add appending content to existing content in a field

class BaseRPA(object):
    def __init__(self, driver=None, logger=None):
        #audit_log('__init__', 'initial selenium program.', base)
        """
        Create an interface to input data to MARC MY system.

        :param WebDriver driver:        Chrome driver instance
        :param Logger logger:           Logger instance
        """

        if logger:
            if isinstance(logger, Logger):
                self.logger = logger
                self.logger.info(
                    'BaseRPA.__init__',
                    'Used the logger %s.' %
                    self.logger
                )
            else:
                self.logger.error(
                    'BaseRPA.__init__',
                    '"logger" accepts an instance of %s, but got %s' %
                    (type(Logger), type(logger))
                )
                raise Exception(
                    '"logger" accepts an instance of %s, but got %s' %
                    (type(Logger), type(logger))
                )
        else:
            self.logger = Logger()
            self.logger.info(
                'BaseRPA.__init__',
                'Loaded the logger %s.' %
                self.logger
            )
        if driver:
            if isinstance(driver, WebDriver):
                self.driver = driver
                self.logger.info(
                    'BaseRPA.__init__',
                    'Used the driver %s.' %
                    self.driver
                )
            else:
                self.logger.error(
                    'BaseRPA.__init__',
                    '"driver" accepts an instance of %s, but got %s' %
                    (type(WebDriver), type(driver))
                )
                raise Exception(
                    '"driver" accepts an instance of %s, but got %s' %
                    (type(WebDriver), type(driver))
                )
        else:
            try:
                self.driver = webdriver.Chrome(
                    os.environ['chromedriver']
                )
            except (se.WebDriverException, KeyError):
                #logging('__init__', KeyError)
                try:
                    if BROWSER_CHOICE == "chrome":
                        driverPath = DEFAULT_PATH_TO_DRIVER
                        option = webdriver.ChromeOptions()
                        option.add_argument("--disable-extensions")
                        #option.setExperimentalOption("useAutomationExtension", false)
                        self.driver = webdriver.Chrome(
                            os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))), DEFAULT_PATH_TO_DRIVER),
                            chrome_options = option
                        )
                    elif BROWSER_CHOICE == "firefox":
                        driverPath = PATH_TO_FIREFOX_DRIVER
                        self.driver = webdriver.Firefox(
                            executable_path = os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))), PATH_TO_FIREFOX_DRIVER)
                        )

                except Exception as e:
                    self.logger.error(
                        'BaseRPA.__init__',
                        'Cannot find the driver at %s.' %
                        driverPath
                    )
                    raise e
                else:
                    self.logger.info(
                        'BaseRPA.__init__',
                        'Loaded the driver at %s.' %
                        driverPath
                    )
                    self.logger.info(
                        'BaseRPA.__init__',
                        'Loaded the driver %s.' %
                        self.driver
                    )
            else:
                self.logger.info(
                    'BaseRPA.__init__',
                    'Loaded the driver at %s.' %
                    os.environ['chromedriver']
                )
                self.logger.info(
                    'BaseRPA.__init__',
                    'Loaded the driver %s.' %
                    self.driver
                )

        self.marc_address = None
        self.uid = None
        self.pw = None

        self.__login_status = None

        self.wait = WebDriverWait(self.driver, TIMEOUT)

    # login/logout methods

    def login(self, base, uid=UID, pw=PW, marc_address=MARC_ADDRESS):
        """
        Log in the MARC MY system using provided uid and pw.

        :param str uid:                 User name
        :param str pw:                  Password
        :param str marc_address:        HTTP address of MARC MY
        :rtype:                         bool
        """

        uid = uid if uid else UID
        pw = pw if pw else PW
        n_login_try = 0
        if self.__login_status:
            self.logger.warning(
                'BaseRPA.login',
                'Has logged in previously, but will log in again.'
            )
           
        while n_login_try < MAX_N_RETRY:
            if n_login_try > 0:
                self.logger.warning(
                    'BaseRPA.login',
                    'Retrying the whole login process with uid %s and pw %s.' %
                    (
                        self.uid,
                        self.pw[:1] + '*' * len(self.pw[1:])
                    )
                )
            self.__login(uid, pw, marc_address)
            n_login_try += 1
            if self.__login_status:
                break
        self.logger.info(
            'BaseRPA.login',
            'Login status with uid %s and pw %s: %s' %
            (
                self.uid,
                self.pw[:1] + '*' * len(self.pw[1:]),
                self.__login_status
            )
        )
        if not self.__login_status:
            self.logger.info(
                'BaseRPA.login',
                'Failed to log in with uid %s and pw %s.' %
                (
                    self.uid,
                    self.pw[:1] + '*' * len(self.pw[1:])
                )
            )
            logging('BaseRPA.login','Failed to log in with uid {0} and pw {1}.'.format(self.uid,self.pw[:1] + '*' * len(self.pw[1:])), base)
            raise Exception(
                'Failed to log in with uid %s and pw %s.' %
                (
                    self.uid,
                    self.pw[:1] + '*' * len(self.pw[1:])
                )
            )
        return self.__login_status

    def __login(self, uid, pw, marc_address):
        """
        Worker function for login()

        :param str uid:                 User name
        :param str pw:                  Password
        :param str marc_address:        HTTP address of MARC MY
        """
        self.uid = uid
        self.pw = pw
        self.marc_address = marc_address

        self.logger.info(
            'BaseRPA.__login',
            'Trying to load the login page at %s' %
            marc_address
        )
        self.driver.get(self.marc_address)

        n_login_page_load_try = 0
        login_page_loaded = False
        uid_elem = None
        pw_elem = None
        while (not login_page_loaded) and (n_login_page_load_try < MAX_N_RETRY):
            try:
                uid_elem = self.wait.until(
                    lambda x: x.find_element_by_xpath(
                        LOGIN_FIELD_NAME_TO_XPATH_MAPPER['login']
                    )
                )
                pw_elem = self.wait.until(
                    lambda x: x.find_element_by_xpath(
                        LOGIN_FIELD_NAME_TO_XPATH_MAPPER['password']
                    )
                )
            except se.TimeoutException:
                self.logger.warning(
                    'BaseRPA.__login',
                    'No uid and password fields on the page %s.' %
                    self.marc_address
                )
                self.logger.warning(
                    'BaseRPA.__login',
                    'The login page of MARC might not be loaded. Trying again.'
                )
                n_login_page_load_try += 1
            else:
                login_page_loaded = True
        if not login_page_loaded:
            self.logger.error(
                'BaseRPA.__login',
                'The login page of MARC might not be loaded.'
            )
            self.__login_status = False
            return
        uid_elem.send_keys(self.uid)
        pw_elem.send_keys(self.pw)
        pw_elem.send_keys(Keys.RETURN)

        try:
            error = self.wait.until(
                lambda x: x.find_element_by_xpath(
                    "//*[contains(text(), 'Login/Password incorrect.')]"
                )
            )
        except se.TimeoutException:
            current_url = self.wait.until(lambda x: x.current_url)
            if current_url == MARC_DASHBOARD_ADDRESS:
                self.logger.info(
                    'BaseRPA.__login',
                    'Login successful with uid %s and pw %s.' %
                    (
                        self.uid,
                        self.pw[:1] + '*' * len(self.pw[1:])
                    )
                )
                self.__login_status = True
            else:
                self.logger.error(
                    'BaseRPA.__login',
                    'Login failed with uid %s and pw %s: '
                    'time is out and cannot go to the dashboard page.' %
                    (
                        self.uid,
                        self.pw[:1] + '*' * len(self.pw[1:])
                    )
                )
                self.__login_status = False
        else:
            self.logger.error(
                'BaseRPA.__login',
                'Login failed with uid %s and pw %s: %s' %
                (
                    self.uid,
                    self.pw[:1] + '*' * len(self.pw[1:]),
                    error.text)
            )
            self.__login_status = False

    def logout(self):
        """
        Log out the MARC MY system.

        :rtype:     bool
        """
        if self.__login_status:
            elem = self.driver.find_element_by_xpath(
                "//*[contains(text(), 'Logout')]"
            )
            elem.click()
            try:
                url = self.wait.until(
                    lambda x: x.current_url
                )
            except se.TimeoutException:
                self.logger.warning(
                    'BaseRPA.logout',
                    'Logout failed with uid %s and pw %s: '
                    'time is out when logging out.' %
                    (
                        self.uid,
                        self.pw[:1] + '*' * len(self.pw[1:])
                    )
                )
                self.driver.quit()
                return False
            else:
                if url == MARC_ADDRESS:
                    self.logger.info(
                        'BaseRPA.logout',
                        'Logout with uid %s and pw %s: successful.' %
                        (
                            self.uid,
                            self.pw[:1] + '*' * len(self.pw[1:])
                        )
                    )
                    self.driver.quit()
                    return True
                else:
                    self.logger.error(
                        'BaseRPA.logout',
                        'Logout failed and current URL: %s' % url
                    )
                    self.driver.quit()
                    return False
        else:
            self.logger.warning(
                'BaseRPA.logout',
                'Has not logged in.'
            )
            self.driver.quit()
            return True

    # locate page/element methods

    def go_to(self, webpage_address):
        """
        Direct browser to specified address

        :param str webpage_address:         Destination website address
        """
        self.driver.get(webpage_address)

    def find_element_on_current_page_by_id_or_name(self, id_or_name):
        """
        Find out the HTML element identified by its id or name.

        :param str id_or_name:      Id or name of the HTML element
        :rtype:                    WebElement
        """
        try:
            elem_list_by_id = self.driver.find_elements_by_id(id_or_name)
            if len(elem_list_by_id) > 1:
                self.logger.warning(
                    'BaseRPA.find_element_on_current_page_by_id_or_name',
                    'There are %d elements identified by html id %s. '
                    'Choose the first one.' %
                    id_or_name
                )
            elem_by_id = elem_list_by_id[0]
        except IndexError:
            elem_by_id = None
        try:
            elem_list_by_name = self.driver.find_elements_by_name(id_or_name)
            if len(elem_list_by_name) > 1:
                self.logger.warning(
                    'BaseRPA.find_element_on_current_page_by_id_or_name',
                    'There are %d elements identified by html name %s. '
                    'Choose the first one.' %
                    id_or_name
                )
            elem_by_name = elem_list_by_name[0]
        except IndexError:
            elem_by_name = None
        if (elem_by_id is not None) and (elem_by_name is not None):
            if elem_by_id == elem_by_name:
                return elem_by_id
            else:
                self.logger.error(
                    'BaseRPA.find_element_on_current_page_by_id_or_name',
                    'elem_by_id is not the same as elem_by_name for target %s.' %
                    id_or_name
                )
                raise se.NoSuchElementException(
                    'elem_by_id is not the same as elem_by_name for target %s.' %
                    id_or_name
                )
        elif (elem_by_id is None) and (elem_by_name is None):
            self.logger.error(
                'BaseRPA.find_element_on_current_page_by_id_or_name',
                'No element identified by id_or_name %s.' %
                id_or_name
            )
            raise se.NoSuchElementException(
                'No element identified by id_or_name %s.' %
                id_or_name
            )
        else:
            if elem_by_id is not None:
                return elem_by_id
            else:
                return elem_by_name

    def find_element_on_current_page_by_xpath(self, xpath):
        """
        Find out the HTML element identified by an xpath
        .
        :param str xpath:      XPath to the HTML element
        :rtype:                WebElement
        """
        try:
            elem_list_by_xpath = self.driver.find_elements_by_xpath(xpath)
            if len(elem_list_by_xpath) > 1:
                self.logger.warning(
                    'BaseRPA.find_element_on_current_page_by_xpath',
                    'Found %d elements identified by xpath %s. '
                    'Choose the first one.' %
                    xpath
                )
            elem_by_xpath = elem_list_by_xpath[0]
        except IndexError:
            self.logger.error(
                'BaseRPA.find_element_on_current_page_by_xpath',
                'No element identified by xpath %s.' %
                xpath
            )
            raise se.NoSuchElementException(
                'No element identified by xpath %s.' %
                xpath
            )
        except Exception as error:
            self.logger.error(
                'BaseRPA.find_element_on_current_page_by_xpath',
                'Cannot find the element identified by xpath %s: %s' %
                (
                    xpath,
                    error
                )
            )
            raise se.NoSuchElementException(
                'Cannot find the element identified by xpath %s: %s' %
                (
                    xpath,
                    error
                )
            )
        else:
            return elem_by_xpath

    # input methods

    def click_on_onclick_on_current_page_by_onclick_content(self, onclick_content, base):
        """
        Click on the onclick button identified by id or name

        :param str onclick_content:     Content of the onclick button
        :rtype:                         bool
        """
        try:
            elem = self.driver.find_element_by_xpath(
                "//a[contains(@onclick, '%s')]" % onclick_content
            )
            elem.click()
        except se.NoSuchElementException:
            self.logger.error(
                'BaseRPA.click_on_onclick_on_current_page_by_id_or_name',
                'Clicking on the onclick html element by id_or_name %s failed: '
                'no such html element.' %
                onclick_content
            )
            return False
        except Exception as error:
            logging('click_on_onclick_on_current_page_by_onclick_content', 'Clicking on the onclick html element by id_or_name %s failed: %s.' %
                (onclick_content, error), base)
            self.logger.error(
                'BaseRPA.click_on_onclick_on_current_page_by_id_or_name',
                'Clicking on the onclick html element by id_or_name %s failed: %s.' %
                (onclick_content, error)
            )
            return False
        else:
            self.logger.info(
                'BaseRPA.click_on_onclick_on_current_page_by_id_or_name',
                'Clicking on the onclick html element by id_or_name %s is successful.' %
                onclick_content
            )
            return True

    def click_on_button_on_current_page_by_id_or_name(self, id_or_name, base):
        """
        Click on the button identified by id or name

        :param str id_or_name:      Id or name of the text field to fill
        :rtype:                     bool
        """
        try:
            elem = self.find_element_on_current_page_by_id_or_name(id_or_name)
            elem.click()
        except se.NoSuchElementException:
            self.logger.error(
                'BaseRPA.click_on_button_on_current_page_by_id_or_name',
                'Clicking on the button of html element by id_or_name %s failed: '
                'no such html element.' %
                id_or_name
            )
            return False
        except Exception as error:
            logging('click_on_button_on_current_page_by_id_or_name', 'Clicking on the button of html element by id_or_name %s failed: %s.' %
                (id_or_name, error), base)
            self.logger.error(
                'BaseRPA.click_on_button_on_current_page_by_id_or_name',
                'Clicking on the button of html element by id_or_name %s failed: %s.' %
                (id_or_name, error)
            )
            return False
        else:
            self.logger.info(
                'BaseRPA.click_on_button_on_current_page_by_id_or_name',
                'Clicking on the button of html element by id_or_name %s is successful.' %
                id_or_name
            )
            return True

    def choose_option_on_current_page_by_id_or_name(self, id_or_name, option_name, base):
        """
        Choose a given option for a select field identified by id or name.

        :param str id_or_name:      Name of the select field
        :param str option_name:     Name of the option to choose
        :rtype:                     bool
        """
        try:
            elem = Select(
                self.find_element_on_current_page_by_id_or_name(id_or_name)
            )
            options = {
                item.text.lower(): item.text
                for item in elem.options
                if 'select one' not in item.text
            }
            if option_name.lower() in options.keys():
                elem.select_by_visible_text(
                    options[option_name.lower()]
                )
            else:
                raise Exception(
                    'The option %s provided for the select field %s '
                    'does not exist in its option list: %s.' %
                    (option_name, id_or_name, options)
                )
        except se.NoSuchElementException:
            self.logger.error(
                'BaseRPA.choose_option_on_current_page_by_id_or_name',
                'Selecting option "%s" to html element by id_or_name %s failed: '
                'no such html element.' %
                (option_name, id_or_name)
            )
            return False
        except Exception as error:
            logging('choose_option_on_current_page_by_id_or_name', 'Selecting option "%s" to html element by id_or_name %s failed: %s.' %
                (option_name, id_or_name, error), base)
            self.logger.error(
                'BaseRPA.choose_option_on_current_page_by_id_or_name',
                'Selecting option "%s" to html element by id_or_name %s failed: %s.' %
                (option_name, id_or_name, error)
            )
            return False
        else:
            self.logger.info(
                'BaseRPA.choose_option_on_current_page_by_id_or_name',
                'Selecting option "%s" to html element by id_or_name %s is successful.' %
                (option_name, id_or_name)
            )
            return True

    def click_on_checkbox_on_current_page_by_id_or_name(self, id_or_name, content, base):
        """
        Enable the checkbox field identified by id or name

        :param str id_or_name:      Id or name of the text field to fill
        :param bool content:         Content to input
        :rtype:                     bool
        """
        try:
            elem = self.find_element_on_current_page_by_id_or_name(id_or_name)
            if content:
                if not elem.is_selected():
                    elem.click()
            else:
                if elem.is_selected():
                    elem.click()
        except se.NoSuchElementException:
            self.logger.error(
                'BaseRPA.click_on_checkbox_on_current_page_by_id_or_name',
                'Switching checkbox to "%s" of html element by id_or_name %s failed: '
                'no such html element.' %
                (content, id_or_name)
            )
            return False
        except Exception as error:
            logging('click_on_checkbox_on_current_page_by_id_or_name', 'Switching checkbox to "%s" of html element by id_or_name %s failed: %s.' %
                (content, id_or_name, error), base)
            self.logger.error(
                'BaseRPA.click_on_checkbox_on_current_page_by_id_or_name',
                'Switching checkbox to "%s" of html element by id_or_name %s failed: %s.' %
                (content, id_or_name, error)
            )
            return False
        else:
            self.logger.info(
                'BaseRPA.click_on_checkbox_on_current_page_by_id_or_name',
                'Switching checkbox to "%s" of html element by id_or_name %s is successful.' %
                (content, id_or_name)
            )
            return True

    def input_content_on_current_page_by_id_or_name(self, id_or_name, content, base):
        """
        Input content to the text field identified by id or name

        :param str id_or_name:      Id or name of the text field to fill
        :param str content:         Content to input
        :rtype:                     bool
        """
        try:
            elem = self.find_element_on_current_page_by_id_or_name(id_or_name)
            elem.clear()
            if isinstance(content, Number) and not np.isnan(content):
                elem.send_keys('%d' % content)
            else:
                elem.send_keys(content)
        except se.NoSuchElementException:
            self.logger.error(
                'BaseRPA.input_content_on_current_page_by_id_or_name',
                'Inputting content "%s" to html element by id_or_name %s failed: '
                'no such html element.' %
                (content, id_or_name)
            )
            return False
        except Exception as error:
            logging('input_content_on_current_page_by_id_or_name', 'Inputting content "%s" to html element by id_or_name %s failed: %s.' %
                (content, id_or_name, error), base)
            self.logger.error(
                'BaseRPA.input_content_on_current_page_by_id_or_name',
                'Inputting content "%s" to html element by id_or_name %s failed: %s.' %
                (content, id_or_name, error)
            )
            return False
        else:
            self.logger.info(
                'BaseRPA.input_content_on_current_page_by_id_or_name',
                'Inputting content "%s" to html element by id_or_name %s is successful.' %
                (content, id_or_name)
            )
            return True

    def input_date_on_current_page_by_id_or_name(self, id_or_name, date, base):
        """
        Input date to the date select field identified by id or name.

        :param str id_or_name:      Id or name of the date select field to fill
        :param str date:            Date string following the format 'dd/mm/yyyy'
        :rtype:                     bool
        """
        if isinstance(date, pd.Timestamp):
            date = date.strftime('%d/%m/%Y')
        elif isinstance(date, datetime.datetime):
            date = date.strftime('%d/%m/%Y')
        try:
            elem = self.find_element_on_current_page_by_id_or_name(id_or_name)
           
            elem.clear()
            elem.send_keys(date)
            elem.send_keys(Keys.ESCAPE)
        except se.NoSuchElementException:
            self.logger.error(
                'BaseRPA.input_date_on_current_page_by_id_or_name',
                'Inputting date "%s" to html element by id_or_name %s failed: '
                'no such html element.' %
                (date, id_or_name)
            )
            return False
        except Exception as error:
            logging('input_date_on_current_page_by_id_or_name', 'Inputting date "%s" to html element by id_or_name %s failed: %s.' %
                (date, id_or_name, error), base)
            self.logger.error(
                'BaseRPA.input_date_on_current_page_by_id_or_name',
                'Inputting date "%s" to html element by id_or_name %s failed: %s.' %
                (date, id_or_name, error)
            )
            return False
        else:
            self.logger.info(
                'BaseRPA.input_date_on_current_page_by_id_or_name',
                'Inputting date "%s" to html element by id_or_name %s is successful.' %
                (date, id_or_name)
            )
            return True

    def click_on_onclick_on_current_page_by_xpath(self, xpath, base):
        """
        Click on the onclick button identified by xpath.

        :param str xpath:     Xpath of the onclick button
        :rtype:               bool
        """
        try:
            elem = self.driver.find_element_by_xpath(xpath)
            elem.click()
        except se.NoSuchElementException:
            self.logger.error(
                'BaseRPA.click_on_onclick_on_current_page_by_xpath',
                'Clicking on the onclick html element by xpath %s failed: '
                'no such html element.' %
                xpath
            )
            return False
        except Exception as error:
            logging('click_on_onclick_on_current_page_by_xpath', 'Clicking on the onclick html element by xpath %s failed: %s.' %
                (
                    xpath,
                    error
                ), base)
            self.logger.error(
                'BaseRPA.click_on_onclick_on_current_page_by_xpath',
                'Clicking on the onclick html element by xpath %s failed: %s.' %
                (
                    xpath,
                    error
                )
            )
            return False
        else:
            self.logger.info(
                'BaseRPA.click_on_onclick_on_current_page_by_xpath',
                'Clicking on the onclick html element by xpath %s is successful.' %
                xpath
            )
            return True

    def click_on_button_on_current_page_by_xpath(self, xpath, base):
        """
        Click on the button identified by xpath.

        :param str xpath:      Xpath of the button to click
        :rtype:                bool
        """
        try:
            elem = self.find_element_on_current_page_by_xpath(xpath)
            elem.click()
        except se.NoSuchElementException:
            self.logger.error(
                'BaseRPA.click_on_button_on_current_page_by_xpath',
                'Clicking on the button of html element by xpath %s failed: '
                'no such html element.' %
                xpath
            )
            return False
        except Exception as error:
            logging('click_on_button_on_current_page_by_xpath', 'Clicking on the button of html element by xpath %s failed: %s.' %
                (
                    xpath,
                    error
                ), base)
            self.logger.error(
                'BaseRPA.click_on_button_on_current_page_by_xpath',
                'Clicking on the button of html element by xpath %s failed: %s.' %
                (
                    xpath,
                    error
                )
            )
            return False
        else:
            self.logger.info(
                'BaseRPA.click_on_button_on_current_page_by_xpath',
                'Clicking on the button of html element by xpath %s is successful.' %
                xpath
            )
            return True

    def choose_option_on_current_page_by_xpath(self, xpath, option_name, base):
        """
        Choose a given option for a select field identified by xpath.

        :param str xpath:           Xpath of the select field
        :param str option_name:     Name of the option to choose
        :rtype:                     bool
        """
        try:
            elem = Select(
                self.find_element_on_current_page_by_xpath(xpath)
            )
            options = {
                item.text.lower(): item.text
                for item in elem.options
                if 'select one' not in item.text
            }
            if option_name.lower() in options.keys():
                elem.select_by_visible_text(
                    options[option_name.lower()]
                )
            else:
                raise se.NoSuchElementException(
                    'The option %s provided for the select field %s '
                    'does not exist in its option list: %s.' %
                    (
                        option_name,
                        xpath,
                        ', '.join(options.values())
                    )
                )
        except se.NoSuchElementException:
            self.logger.error(
                'BaseRPA.choose_option_on_current_page_by_xpath',
                'Selecting option "%s" to html element by xpath %s failed: '
                'no such html element.' %
                (
                    option_name,
                    xpath
                )
            )
            return False
        except Exception as error:
            logging('choose_option_on_current_page_by_xpath', 'Selecting option "%s" to html element by xpath %s failed: %s.' %
                (
                    option_name,
                    xpath,
                    error
                ), base)
            self.logger.error(
                'BaseRPA.choose_option_on_current_page_by_xpath',
                'Selecting option "%s" to html element by xpath %s failed: %s.' %
                (
                    option_name,
                    xpath,
                    error
                )
            )
            return False
        else:
            self.logger.info(
                'BaseRPA.choose_option_on_current_page_by_xpath',
                'Selecting option "%s" to html element by xpath %s is successful.' %
                (
                    option_name,
                    xpath
                )
            )
            return True

    def click_on_checkbox_on_current_page_by_xpath(self, xpath, content, base):
        """
        Enable the checkbox field identified by xpath.

        :param str xpath:           Xpath of the text field to fill
        :param bool content:         Content to input
        :rtype:                     bool
        """
        try:
            elem = self.find_element_on_current_page_by_xpath(xpath)
            if content:
                if not elem.is_selected():
                    elem.click()
            else:
                if elem.is_selected():
                    elem.click()
        except se.NoSuchElementException:
            self.logger.error(
                'BaseRPA.click_on_checkbox_on_current_page_by_xpath',
                'Switching checkbox to "%s" of html element by xpath %s failed: '
                'no such html element.' %
                (
                    content,
                    xpath
                )
            )
            return False
        except Exception as error:
            logging('click_on_checkbox_on_current_page_by_xpath', 'Switching checkbox to "%s" of html element by xpath %s failed: %s.' %
                (
                    content,
                    xpath,
                    error
                ), base)
            self.logger.error(
                'BaseRPA.click_on_checkbox_on_current_page_by_xpath',
                'Switching checkbox to "%s" of html element by xpath %s failed: %s.' %
                (
                    content,
                    xpath,
                    error
                )
            )
            return False
        else:
            self.logger.info(
                'BaseRPA.click_on_checkbox_on_current_page_by_xpath',
                'Switching checkbox to "%s" of html element by xpath %s is successful.' %
                (
                    content,
                    xpath
                )
            )
            return True

    def input_content_on_current_page_by_xpath(self, xpath, content, base):
        """
        Input content to the text field identified by xpath.

        :param str xpath:           Xpath of the text field to fill
        :param str content:         Content to input
        :rtype:                     bool
        """
        try:
            elem = self.find_element_on_current_page_by_xpath(xpath)
            elem.clear()
            if isinstance(content, Number) and not np.isnan(content):
                elem.send_keys('%d' % content)
            elif type(content) == float and np.isnan(content):
                elem.send_keys("")
            else:
                elem.send_keys(content)
        except se.NoSuchElementException:
            self.logger.error(
                'BaseRPA.input_content_on_current_page_by_xpath',
                'Inputting content "%s" to html element by xpath %s failed: '
                'no such html element.' %
                (
                    content,
                    xpath
                )
            )
            return False
        except Exception as error:
            logging('input_content_on_current_page_by_xpath', 'Inputting content "%s" to html element by xpath %s failed: %s.' %(content,xpath,error), base)
            self.logger.error(
                'BaseRPA.input_content_on_current_page_by_xpath',
                'Inputting content "%s" to html element by xpath %s failed: %s.' %
                (
                    content,
                    xpath,
                    error
                )
            )
            return False
        else:
            self.logger.info(
                'BaseRPA.input_content_on_current_page_by_xpath',
                'Inputting content "%s" to html element by xpath %s is successful.' %
                (
                    content,
                    xpath
                )
            )
            return True

    def input_date_on_current_page_by_xpath(self, xpath, date, base):
        """
        Input date to the date select field identified by xpath.

        :param str xpath:       Xpath of the date select field to fill
        :param str date:        Date string following the format 'dd/mm/yyyy'
        :rtype:                 bool
        """
        if isinstance(date, pd.Timestamp):
            date = date.strftime('%d/%m/%Y')
        elif isinstance(date, datetime.datetime):
            date = date.strftime('%d/%m/%Y')
        try:
            elem = self.find_element_on_current_page_by_xpath(xpath)
            elem.clear()
            elem.send_keys(date)
            elem.send_keys(Keys.ESCAPE)
        except se.NoSuchElementException:
            self.logger.error(
                'BaseRPA.input_date_on_current_page_by_xpath',
                'Inputting date "%s" to html element by xpath %s failed: '
                'no such html element.' %
                (
                    date,
                    xpath
                )
            )
            return False
        except Exception as error:
            logging('input_date_on_current_page_by_xpath', 'Inputting date "%s" to html element by xpath %s failed: %s.' %
                (
                    date,
                    xpath,
                    error
                ), base)
            self.logger.error(
                'BaseRPA.input_date_on_current_page_by_xpath',
                'Inputting date "%s" to html element by xpath %s failed: %s.' %
                (
                    date,
                    xpath,
                    error
                )
            )
            return False
        else:
            self.logger.info(
                'BaseRPA.input_date_on_current_page_by_xpath',
                'Inputting date "%s" to html element by xpath %s is successful.' %
                (
                    date,
                    xpath
                )
            )
            return True

    # all-in-one input method

    def input_on_current_page_by_field_name(self, loc_mapper, type_mapper, field_name, base, content=None, method='xpath'):
        """
        Input anything to a field on the current page.

        :param dict loc_mapper:         Maps field name to field id_or_name or xpath
        :param dict type_mapper:        Maps field name to field type
        :param str field_name:          Full name of the field
        :param str content:             Content to be filled
        :param str method:              Method to input the content: 'xpath' or 'id_or_name'
        :rtype:                         bool
        """
        if method == 'id_or_name':
            try:
                field_id_or_name = loc_mapper[field_name.lower()]
            except KeyError:
                self.logger.error(
                    'BaseRPA.input_on_current_page_by_field_name',
                    'No such field_name %s contained in the id or name mapper.' %
                    field_name
                )
                return False
            else:
                field_type = type_mapper[field_name.lower()]
                if field_type == 'input':
                    return self.input_content_on_current_page_by_id_or_name(field_id_or_name, content, base)
                elif field_type == 'select':
                    return self.choose_option_on_current_page_by_id_or_name(field_id_or_name, content, base)
                elif field_type == 'date_input':
                    return self.input_date_on_current_page_by_id_or_name(field_id_or_name, content, base)
                elif field_type == 'checkbox':
                    return self.click_on_checkbox_on_current_page_by_id_or_name(field_id_or_name, content, base)
                elif field_type == 'button':
                    return self.click_on_button_on_current_page_by_id_or_name(field_id_or_name, base)
                elif field_type == 'onclick':
                    return self.click_on_onclick_on_current_page_by_onclick_content(field_id_or_name, base)
                else:
                    self.logger.error(
                        'BaseRPA.input_on_current_page_by_field_name',
                        'Got wrong type %s from the type mapper for the field %s.' %
                        (
                            field_type,
                            field_name
                        )
                    )
                    raise ValueError(
                        'Got wrong type %s from the type mapper for the field %s.' %
                        (
                            field_type,
                            field_name
                        )
                    )
        elif method == 'xpath':
            try:
                field_xpath = loc_mapper[field_name.lower()]
            except KeyError:
                self.logger.error(
                    'BaseRPA.input_on_current_page_by_field_name',
                    'No such field_name %s contained in the xpath mapper.' %
                    field_name
                )
                return False
            else:
                field_type = type_mapper[field_name.lower()]
                if field_type == 'input':
                    return self.input_content_on_current_page_by_xpath(field_xpath, content, base)
                elif field_type == 'select':
                    return self.choose_option_on_current_page_by_xpath(field_xpath, content, base)
                elif field_type == 'date_input':
                    return self.input_date_on_current_page_by_xpath(field_xpath, content, base)
                elif field_type == 'checkbox':
                    return self.click_on_checkbox_on_current_page_by_xpath(field_xpath, content, base)
                elif field_type == 'button':
                    return self.click_on_button_on_current_page_by_xpath(field_xpath, base)
                elif field_type == 'onclick':
                    return self.click_on_onclick_on_current_page_by_xpath(field_xpath, base)
                else:
                    self.logger.error(
                        'BaseRPA.input_on_current_page_by_field_name',
                        'Got wrong type %s from the type mapper for the field %s.' %
                        (
                            field_type,
                            field_name
                        )
                    )
                    raise ValueError(
                        'Got wrong type %s from the type mapper for the field %s.' %
                        (
                            field_type,
                            field_name
                        )
                    )
        else:
            self.logger.error(
                'BaseRPA.input_on_current_page_by_field_name',
                'Got a wrong method %s. Should be one among %s.' %
                (
                    method,
                    ', '.join(
                        [
                            'xpath',
                            'id_or_name'
                        ]
                    )
                )
            )
            raise Exception(
                'Got a wrong method %s. Should be one among %s.' %
                (
                    method,
                    ', '.join(
                        [
                            'xpath',
                            'id_or_name'
                        ]
                    )
                )
            )

    # get methods

    def get_content_on_current_page_by_id_or_name(self, id_or_name):
        """
        Get the record of the field identified by id or name.

        :param str id_or_name:      Id or name of the field
        :rtype:                     str
        """
        elem = self.find_element_on_current_page_by_id_or_name(id_or_name)
        if elem.tag_name == 'select':
            return Select(elem).first_selected_option.text
        else:
            return elem.text

    def get_content_on_current_page_by_xpath(self, xpath):
        """
        Get the record of the field identified by xpath.

        :param str xpath:      Xpath of the field
        :rtype:                str
        """
        elem = self.find_element_on_current_page_by_xpath(xpath)
        if elem.tag_name == 'select':
            return Select(elem).first_selected_option.text
        else:
            return elem.text

    def get_content_on_current_page_by_field_name(self, loc_mapper, field_name, method='xpath'):
        """
        Get the content of an element identified by field name.

        :param dict loc_mapper:     Maps field name to field id_or_name or xpath
        :param str field_name:      Full name of the field
        :param str method:          Method to input the content: 'xpath' or 'id_or_name'
        :rtype:                     str
        """
        if method == 'id_or_name':
            try:
                field_id_or_name = loc_mapper[field_name.lower()]
            except KeyError:
                self.logger.error(
                    'BaseRPA.get_content_on_current_page_by_field_name',
                    'No such field_name %s contained in the id or name mapper.' %
                    field_name
                )
                return None
            else:
                return self.get_content_on_current_page_by_id_or_name(
                    field_id_or_name
                )
        elif method == 'xpath':
            try:
                field_xpath = loc_mapper[field_name.lower()]
            except KeyError:
                self.logger.error(
                    'BaseRPA.get_content_on_current_page_by_field_name',
                    'No such field_name %s contained in the xpath mapper.' %
                    field_name
                )
                return None
            else:
                return self.get_content_on_current_page_by_xpath(
                    field_xpath
                )
        else:
            self.logger.error(
                'BaseRPA.get_content_on_current_page_by_field_name',
                'Got a wrong method %s. Should be one among %s.' %
                (
                    method,
                    ', '.join(
                        [
                            'xpath',
                            'id_or_name'
                        ]
                    )
                )
            )
            raise Exception(
                'Got a wrong method %s. Should be one among %s.' %
                (
                    method,
                    ', '.join(
                        [
                            'xpath',
                            'id_or_name'
                        ]
                    )
                )
            )

    def get_options_of_select_on_current_page_by_id_or_name(self, id_or_name):
        """
        Get the options of a select element identified by field name or id.

        :param str id_or_name:                      id or full name of field
        :rtype:                                                 str
        """
        try:
            elem = Select(
                self.find_element_on_current_page_by_id_or_name(id_or_name)
            )
        except se.NoSuchElementException:
            self.logger.error(
                'BaseRPA.get_options_of_select_on_current_page_by_id_or_name',
                'Listing options for field_id %s failed: no such html element.' %
                id_or_name
            )
            return None
        else:
            options = {
                item.text
                for item in elem.options
                if 'select one' not in item.text
            }
            return options

    def get_options_of_select_on_current_page_by_xpath(self, xpath):
        """
        Get the options of a select element identified by field name or id.

        :param str xpath:       id or full name of field
        :rtype:                 str
        """
        try:
            elem = Select(
                self.find_element_on_current_page_by_xpath(xpath)
            )
        except se.NoSuchElementException:
            self.logger.error(
                'BaseRPA.get_options_of_select_on_current_page_by_xpath',
                'Listing options for xpath %s failed: no such html element.' %
                xpath
            )
            return None
        else:
            options = {
                item.text
                for item in elem.options
                if 'select one' not in item.text
            }
            return options

    def get_options_of_select_on_current_page_by_field_name(self, loc_mapper, field_name, method='xpath'):
        """
        Get the options of a select element identified by field name.

        :param dict loc_mapper:     Maps field name to field id_or_name or xpath
        :param str field_name:      Full name of the field
        :param str method:          Method to input the content: 'xpath' or 'id_or_name'
        :rtype:                     str
        """
        if method == 'id_or_name':
            try:
                field_id_or_name = loc_mapper[field_name.lower()]
            except KeyError:
                self.logger.error(
                    'BaseRPA.get_options_of_select_on_current_page_by_field_name',
                    'No such field_name %s contained in the id or name mapper.' %
                    field_name
                )
                return None
            else:
                return self.get_options_of_select_on_current_page_by_id_or_name(
                    field_id_or_name
                )
        elif method == 'xpath':
            try:
                field_xpath = loc_mapper[field_name.lower()]
            except KeyError:
                self.logger.error(
                    'BaseRPA.get_options_of_select_on_current_page_by_field_name',
                    'No such field_name %s contained in the xpath mapper.' %
                    field_name
                )
                return None
            else:
                return self.get_options_of_select_on_current_page_by_xpath(
                    field_xpath
                )
        else:
            self.logger.error(
                'BaseRPA.get_options_of_select_on_current_page_by_field_name',
                'Got a wrong method %s. Should be one among %s.' %
                (
                    method,
                    ', '.join(
                        [
                            'xpath',
                            'id_or_name'
                        ]
                    )
                )
            )
            raise Exception(
                'Got a wrong method %s. Should be one among %s.' %
                (
                    method,
                    ', '.join(
                        [
                            'xpath',
                            'id_or_name'
                        ]
                    )
                )
            )

    # wait methods
    @staticmethod
    def sleep(n_second):
        time.sleep(n_second)

    def wait_until_clickability_of_element_by_field_name(self, loc_mapper, field_name, method='xpath'):
        """
        Wait for form element to become active.

        :param dict loc_mapper:     Maps field name to field id or xpath
        :param str field_name:      Full name of the field
        :param str method:          Method to input the content: 'xpath' or 'id_or_name'
        :rtype:                     bool
        """
        if method == 'id_or_name':
            try:
                field_id = loc_mapper[field_name.lower()]
            except KeyError:
                self.logger.error(
                    'BaseRPA.wait_until_clickability_of_element_by_field_name',
                    'No such field_name %s contained in the id mapper.' %
                    field_name
                )
                return False
            else:
                try:
                    self.wait.until(
                        ec.element_to_be_clickable(
                            (By.ID, field_id)
                        )
                    )
                except se.TimeoutException:
                    self.logger.warning(
                        'BaseRPA.wait_until_clickability_of_element_by_field_name',
                        'Waiting for %s to be clickable takes too long.' %
                        field_name
                    )
                    return False
                else:
                    return True
        elif method == 'xpath':
            try:
                field_xpath = loc_mapper[field_name.lower()]
            except KeyError:
                self.logger.error(
                    'BaseRPA.wait_until_clickability_of_element_by_field_name',
                    'No such field_name %s contained in the xpath mapper.' %
                    field_name
                )
                return False
            else:
                try:
                    self.wait.until(
                        ec.element_to_be_clickable(
                            (By.XPATH, field_xpath)
                        )
                    )
                except se.TimeoutException:
                    self.logger.warning(
                        'BaseRPA.wait_until_clickability_of_element_by_field_name',
                        'Waiting for %s to be clickable takes too long.' %
                        field_name
                    )
                    return False
                else:
                    return True
        else:
            self.logger.error(
                'BaseRPA.wait_until_clickability_of_element_by_field_name',
                'Got a wrong method %s. Should be one among %s.' %
                (
                    method,
                    ', '.join(
                        [
                            'xpath',
                            'id_or_name'
                        ]
                    )
                )
            )
            raise Exception(
                'Got a wrong method %s. Should be one among %s.' %
                (
                    method,
                    ', '.join(
                        [
                            'xpath',
                            'id_or_name'
                        ]
                    )
                )
            )

    def wait_until_visibility_of_element_by_field_name(self, loc_mapper, field_name, method='xpath'):
        """
        Wait for form element to become visible.

        :param dict loc_mapper:     Maps field name to field id
        :param str field_name:      Full name of the field
        :param str method:          Method to input the content: 'xpath' or 'id_or_name'
        :rtype:                     bool
        """
        if method == 'id_or_name':
            try:
                field_id = loc_mapper[field_name.lower()]
            except KeyError:
                self.logger.error(
                    'BaseRPA.wait_until_visibility_of_element_by_field_name',
                    'No such field_name %s contained in the id mapper.' %
                    field_name
                )
                return False
            else:
                try:
                    self.wait.until(
                        ec.visibility_of_element_located(
                            (By.ID, field_id)
                        )
                    )
                except se.TimeoutException:
                    self.logger.warning(
                        'BaseRPA.wait_until_visibility_of_element_by_field_name',
                        'Waiting for %s to appear takes too long.' %
                        field_name
                    )
                    return False
                else:
                    return True
        elif method == 'xpath':
            try:
                field_xpath = loc_mapper[field_name.lower()]
            except KeyError:
                self.logger.error(
                    'BaseRPA.wait_until_visibility_of_element_by_field_name',
                    'No such field_name %s contained in the xpath mapper.' %
                    field_name
                )
                return False
            else:
                try:
                    self.wait.until(
                        ec.visibility_of_element_located(
                            (By.XPATH, field_xpath)
                        )
                    )
                except se.TimeoutException:
                    self.logger.warning(
                        'BaseRPA.wait_until_visibility_of_element_by_field_name',
                        'Waiting for %s to appear takes too long.' %
                        field_name
                    )
                    return False
                else:
                    return True
        else:
            self.logger.error(
                'BaseRPA.wait_until_visibility_of_element_by_field_name',
                'Got a wrong method %s. Should be one among %s.' %
                (
                    method,
                    ', '.join(
                        [
                            'xpath',
                            'id_or_name'
                        ]
                    )
                )
            )
            raise Exception(
                'Got a wrong method %s. Should be one among %s.' %
                (
                    method,
                    ', '.join(
                        [
                            'xpath',
                            'id_or_name'
                        ]
                    )
                )
            )

    def wait_until_invisibility_of_element_by_field_name(self, loc_mapper, field_name, method='xpath'):
        """
        Wait for form element to become invisible.

        :param dict loc_mapper:     Maps field name to field id
        :param str field_name:      Full name of the field
        :param str method:          Method to input the content: 'xpath' or 'id_or_name'
        :rtype:                     bool
        """
        if method == 'id_or_name':
            try:
                field_id = loc_mapper[field_name.lower()]
            except KeyError:
                self.logger.error(
                    'BaseRPA.wait_until_invisibility_of_element_by_field_name',
                    'No such field_name %s contained in the id mapper.' %
                    field_name
                )
                return False
            else:
                try:
                    self.wait.until(
                        ec.invisibility_of_element_located(
                            (By.ID, field_id)
                        )
                    )
                except se.TimeoutException:
                    self.logger.warning(
                        'BaseRPA.wait_until_invisibility_of_element_by_field_name',
                        'Waiting for %s to appear takes too long.' %
                        field_name
                    )
                    return False
                else:
                    return True
        elif method == 'xpath':
            try:
                field_xpath = loc_mapper[field_name.lower()]
            except KeyError:
                self.logger.error(
                    'BaseRPA.wait_until_invisibility_of_element_by_field_name',
                    'No such field_name %s contained in the xpath mapper.' %
                    field_name
                )
                return False
            else:
                try:
                    self.wait.until(
                        ec.invisibility_of_element_located(
                            (By.XPATH, field_xpath)
                        )
                    )
                except se.TimeoutException:
                    self.logger.warning(
                        'BaseRPA.wait_until_invisibility_of_element_by_field_name',
                        'Waiting for %s to appear takes too long.' %
                        field_name
                    )
                    return False
                else:
                    return True
        else:
            self.logger.error(
                'BaseRPA.wait_until_invisibility_of_element_by_field_name',
                'Got a wrong method %s. Should be one among %s.' %
                (
                    method,
                    ', '.join(
                        [
                            'xpath',
                            'id_or_name'
                        ]
                    )
                )
            )
            raise Exception(
                'Got a wrong method %s. Should be one among %s.' %
                (
                    method,
                    ', '.join(
                        [
                            'xpath',
                            'id_or_name'
                        ]
                    )
                )
            )

    def wait_until_revealed_of_element_by_xpath(self, xpath):
        """
        Wait for element to become visible. Takes xpath of the object and checks aria_hidden attribute

        :param str xpath:       Element to monitor
        :rtype:                 bool
        """
        try:
            elem = self.driver.find_element_by_xpath(xpath)
        except se.NoSuchElementException:
            self.logger.error(
                'BaseRPA.wait_until_revealed_of_element_by_xpath',
                'No such element %s contained in the xpath.' %
                xpath
            )

        def revealed(_):
            try:
                return elem.get_attribute("aria-hidden") == "false"
            except se.NoSuchAttributeException:
                self.logger.error(
                    'BaseRPA.wait_until_revealed_of_element_by_xpath',
                    'No such element aria-hidden contained in the element.'
                )

        self.wait.until(revealed)

    def wait_until_hidden_of_element_by_xpath(self, xpath):
        """
        Wait for element to become invisible. Takes xpath of the object and checks aria_hidden attribute

        :param str xpath:       Element to monitor
        :rtype:                 bool
        """
        try:
            elem = self.driver.find_element_by_xpath(xpath)
        except se.NoSuchElementException:
            self.logger.error(
                'BaseRPA.wait_until_hidden_of_element_by_xpath',
                'No such element %s contained in the xpath.' %
                xpath
            )

        def hidden(_):
            try:
                return elem.get_attribute("aria-hidden") == "true"
            except se.NoSuchAttributeException:
                self.logger.error(
                    'BaseRPA.wait_until_hidden_of_element_by_xpath',
                    'No such element aria-hidden contained in the element.'
                )

        self.wait.until(hidden)

    def wait_until_visibility_of_message(self, message_mapper):
        """
                        Wait for form element to become visible.

                        :param dict message_mapper:      Maps field name to field id
                        :rtype:                     str, bool
                  """
        # TODO: may be change to use text_to_be_present_in_element_value
        waiter = self.message_waiter_wrapper(message_mapper)
        try:
            message, result = self.wait.until(waiter)
        except se.TimeoutException:
            self.logger.warning(
                'BaseRPA.wait_until_visibility_of_message',
                'Waiting for one among the messages %s to appear takes too long.' %
                ', '.join(message_mapper.keys())
            )
            raise se.TimeoutException(
                'Waiting for one among the messages %s to appear takes too long.' %
                ', '.join(message_mapper.keys())
            )
        else:
            return message, result

    @staticmethod
    def message_waiter_wrapper(message_mapper, elem_type='*'):
        """
        Returns a function that waits for a type of field given by elem_type.

        :param dict message_mapper:         Maps field name to field id
        :param str elem_type:                     Field type
        :return:                                           func
        """
        def message_waiter(driver):
            xpath_template = (
                    '//%s[contains(text(), "%s")]' %
                    (elem_type, '%s')
            )
            for text in message_mapper:
                try:
                    elem = driver.find_element_by_xpath(
                        xpath_template % text
                    )
                except se.NoSuchElementException:
                    pass
                else:
                    return elem.text, message_mapper[text]
            raise se.NoSuchElementException(
                'Has not found one among the messages: %s.' %
                ', '.join(message_mapper.keys())
            )

        return message_waiter

    # validation
    def fields_in_row(self, fields, row, op):
        """
                Checks that all the input fields are in a row
                Use this at the beginning of operation functions to ensure that essential fields are present

                :param list fields:                           List of fields to check
                :param dict row:                            Row to check
                :param string op:                           Name of operation that field is for.
                :return:                                          bool
                """
        for field in fields:
            if field not in row:
                self.logger.error(
                    'BaseRPA.fields_in_row',
                    'Field %s is not  in operation %s' %
                    (
                        field,
                        op
                     )
                )
                return False
        return True

    def wait_for_user_confirmation(self, message="Press ok to continue"):
        if TOGGLE_UPDATE_CONFIRMATION:
            ctypes.windll.user32.MessageBoxW(0, message, "RPA paused", 0x1000)
