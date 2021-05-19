#!/usr/bin/env python
#
#
# The file contains
#   1. BulkUploadRPA
#

import pandas as pd

from selenium.common import exceptions as se

from selenium_interface.basicops import BaseRPA
from selenium_interface.settings import *
from pathlib import Path
import os
from utils.logging import logging
from utils.audit_trail import audit_log

class BulkUploadRPA(BaseRPA):
    def __init__(self, driver=None, logger=None):
        """
        Create an interface to input data to MARC MY system.
        :param WebDriver driver:        Chrome driver instance
        :param Logger logger:           Logger instance
        """
        super(BulkUploadRPA, self).__init__(driver, logger)

    def up_upload_file(self, base, file):
        """


	    :param file:
	    :return:
	    """
        success = False
        try:
            success = self.upload_action(file)
        except:
            self.logger.error(
                'BulkUploadRPA.up_upload_file',
                'Unexpected error in file upload'
            )
            logging('up_upload_file', 'Unexpected error in file upload.', base)


        try:
            if success:
                os.makedirs(os.path.join(RELATIVE_PATH_TO_FOLDER_OF_OUTPUT_XLSX_FOR_BULK_UPLOAD, "good"), exist_ok=True)
                try:
                    os.rename(file, os.path.join(RELATIVE_PATH_TO_FOLDER_OF_OUTPUT_XLSX_FOR_BULK_UPLOAD, "good", Path(file).name))
                except:
                    os.remove(os.path.join(RELATIVE_PATH_TO_FOLDER_OF_OUTPUT_XLSX_FOR_BULK_UPLOAD, "good", Path(file).name))
                    os.rename(file, os.path.join(RELATIVE_PATH_TO_FOLDER_OF_OUTPUT_XLSX_FOR_BULK_UPLOAD, "good", Path(file).name))
            else:
                os.makedirs(os.path.join(RELATIVE_PATH_TO_FOLDER_OF_OUTPUT_XLSX_FOR_BULK_UPLOAD, "bad"), exist_ok=True)
                try:
                    os.rename(file, os.path.join(RELATIVE_PATH_TO_FOLDER_OF_OUTPUT_XLSX_FOR_BULK_UPLOAD, "bad", Path(file).name))
                except:
                    os.remove(os.path.join(RELATIVE_PATH_TO_FOLDER_OF_OUTPUT_XLSX_FOR_BULK_UPLOAD, "bad", Path(file).name))
                    os.rename(file, os.path.join(RELATIVE_PATH_TO_FOLDER_OF_OUTPUT_XLSX_FOR_BULK_UPLOAD, "bad", Path(file).name))
        except:
            self.logger.error(
                'BulkUploadRPA.up_upload_file',
                'Could not overwrite existing file, possibly open already'
            )
            logging('up_upload_file', 'Could not overwrite existing file, possibly open already.', base)

        self.go_to(MARC_DASHBOARD_ADDRESS)


    def upload_action(self, base, file):
        flags = list()
        self.go_to(MARC_INPT_BULK_UPLOAD_ADDRESS)

        self.find_element_on_current_page_by_xpath(BULK_IMPORT_FIELD_NAME_TO_XPATH_MAPPER["choose"]).send_keys(file)

        self.wait_for_user_confirmation("Press ok to upload")
        flags.append(
            self.up_input_content_by_field_name(
                "upload"
            )
        )

        try:
            self.wait.until(
                lambda x: x.find_elements_by_xpath(BULK_IMPORT_FIELD_NAME_TO_XPATH_MAPPER["result"])
            )
        except se.TimeoutException:
            self.logger.error(
                'BulkUploadRPA.up_upload_file',
                "Bulk upload timed out"
            )
            logging('upload_action', 'Bulk upload timed out.', base)
            return False

        return all(flags)

    def up_input_content_by_field_name(self, field_name, content=None):
        """

	    :param str field_name:      Full name of the field
	    :param str content:         Content to input
	    :rtype:                     str
	    """
        return self.input_on_current_page_by_field_name(
		    BULK_IMPORT_FIELD_NAME_TO_XPATH_MAPPER,
		    BULK_IMPORT_FIELD_NAME_TO_TYPE_MAPPER,
		    field_name,
		    content
	    )

    def up_get_content_by_field_name(self, field_name):
        """
        Get content of record in Member Setup.
        :param str where:           Which page to use
        :param str field_name:      Full name of the field
        :rtype:                     str
        """
        return self.get_content_on_current_page_by_field_name(
	        BULK_IMPORT_FIELD_NAME_TO_XPATH_MAPPER,
            field_name
        )
