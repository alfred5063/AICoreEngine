#!/usr/bin/env python
#
# Created by Xixuan on 29/08/2018
#
# This is a major interface of everything.
#
# The file contains
#   1. MARC RPA engine
#

from selenium.webdriver.chrome.webdriver import WebDriver

from selenium_interface.memberSetup import MemberSetupRPA
from selenium_interface.policyManipulation import PolicyManipulationRPA
from selenium_interface.bulkUpload import BulkUploadRPA
from selenium_interface.logger import Logger
from utils.logging import logging
from utils.audit_trail import audit_log

class RPA(MemberSetupRPA, PolicyManipulationRPA, BulkUploadRPA):
    def __init__(self, driver=None, logger=None):
        """
        Create an interface to input data to MARC MY system.
        :param WebDriver driver:        Chrome driver instance
        :param Logger logger:           Logger instance
        """
        super(RPA, self).__init__(driver, logger)
