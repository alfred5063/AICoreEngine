#!/usr/bin/env python
#
# Created by Xixuan on 29/08/2018
#
# The file contains
#   1. a monitor of the folders of input data files
#

import os

from selenium_interface.tools import *
from selenium_interface.settings import *
from regex import findall


class Monitor(object):
    def __init__(self, paths_under_monitoring=None, refresh_time=MONITOR_REFRESH_TIME):
        """
        Create a monitor to keep checking the folders for input files.

        Note:   If provides paths_under_monitoring, one should provide
                RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_POLICY_MANIPULATION
                and
                RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_MEMBER_SETUP
                at the same time.

        :param dict paths_under_monitoring:     List of folders under monitoring
        :param int refresh_time:                Number of seconds between refreshing
        """
        if paths_under_monitoring is None:
            self.__full_path_to_folder_of_input_files_for_pm = os.path.abspath(
                RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_POLICY_MANIPULATION
            )
            self.__full_path_to_folder_of_input_files_for_ms = os.path.abspath(
                RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_MEMBER_SETUP
            )
        else:
            try:
                self.__full_path_to_folder_of_input_files_for_pm = os.path.abspath(
                    paths_under_monitoring[
                        'RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_POLICY_MANIPULATION'
                    ]
                )
                self.__full_path_to_folder_of_input_files_for_ms = os.path.abspath(
                    paths_under_monitoring[
                        'RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_MEMBER_SETUP'
                    ]
                )
                self.__full_path_to_folder_of_input_files_for_upload = os.path.abspath(
                    paths_under_monitoring[
                        'RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_BULK_UPLOAD'
                    ]
                )
            except KeyError as error:
                raise KeyError(
                    'Cannot find out current keys in paths_under_monitoring: %s' %
                    error
                )
            except Exception as error:
                raise error

        self.__refresh_time = refresh_time

        # for the folders that hold input files and output files
        if not os.path.exists(self.__full_path_to_folder_of_input_files_for_pm):
            os.mkdir(self.__full_path_to_folder_of_input_files_for_pm)
        if not os.path.exists(self.__full_path_to_folder_of_input_files_for_ms):
            os.mkdir(self.__full_path_to_folder_of_input_files_for_ms)
        if not os.path.exists(self.__full_path_to_folder_of_input_files_for_upload):
            os.mkdir(self.__full_path_to_folder_of_input_files_for_upload)

    def refresh(self):
        """
        Refresh and check if get new input files.
        """
        #time.sleep(self.__refresh_time)
        self.__file_list_for_pm = []
        self.__file_list_for_ms = []
        self.__file_list_for_upload = []
        pm_temp = os.listdir(
            self.__full_path_to_folder_of_input_files_for_pm
        )
        ms_temp = os.listdir(
            self.__full_path_to_folder_of_input_files_for_ms
        )
        upload_temp = os.listdir(
            self.__full_path_to_folder_of_input_files_for_upload
        )

        for filename in pm_temp:
            if len(findall("(?i)\\.xls[x]?$", filename)) > 0:
                self.__file_list_for_pm.append(filename)
        for filename in ms_temp:
            if len(findall("(?i)\\.xls[x]?$", filename)) > 0:
                self.__file_list_for_ms.append(filename)
        for filename in upload_temp:
            if len(findall("(?i)\\.xls[x]?$", filename)) > 0:
                self.__file_list_for_upload.append(os.path.join(self.__full_path_to_folder_of_input_files_for_upload, filename))

        current_time = ct()
        print(
            '[%s] Monitor has refreshed for \n%s%s \n%sand \n%s%s.' %
            (
                current_time,
                ''.join([' ' for _ in range(len(current_time) + 3)]),
                self.__full_path_to_folder_of_input_files_for_pm,
                ''.join([' ' for _ in range(len(current_time) + 3)]),
                ''.join([' ' for _ in range(len(current_time) + 3)]),
                self.__full_path_to_folder_of_input_files_for_ms
            )
        )
        if (len(self.__file_list_for_pm) + len(self.__file_list_for_ms) + len(self.__file_list_for_upload) ) > 0:
            print(
                '[%s] Found files %s.' %
                (
                    ct(),
                    ', '.join(
                        self.__file_list_for_pm
                        +
                        self.__file_list_for_ms
                        +
                        self.__file_list_for_upload
                    )
                )
            )

    def file_list_for_pm(self):
        """
        Provide the names of the new files for policy manipulation.
        :return:
        """
        return self.__file_list_for_pm

    def file_list_for_ms(self):
        """
        Provide the names of the new files for member setup.
        :return:
        """
        return self.__file_list_for_ms

    def file_list_for_upload(self):
        """
        Provide the names of the new files for bulk upload.
        :return:
        """
        return self.__file_list_for_upload
