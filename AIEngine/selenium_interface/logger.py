#!/usr/bin/env python
#
# Created by Xixuan on 29/08/2018
#
# The file contains
#   1. a logger for logging all messages
#

import os
import sys
import logging

from logging.handlers import TimedRotatingFileHandler

from selenium_interface.tools import *
from selenium_interface.settings import *


class Logger(object):
    def __init__(self, log_root_folder=None):
        """
        Create a logger.
        :param str log_root_folder: Relative path to the root folder of all logs
        """
        self.__logger = logging.getLogger('rpa')
        if len(self.__logger.handlers) == 0:
            self.__log_formatter = logging.Formatter(LOG_FORMATTER)

            self.__console_handler = logging.StreamHandler(sys.stdout)
            self.__console_handler.setFormatter(self.__log_formatter)
            if ONLY_DISPLAY_ERRORS_IN_CONSOLE:
                self.__console_handler.setLevel(logging.WARNING)

            self.__log_file_name = '%s%s.log' % (LOG_PREFIX, ct(plain=True))
            if log_root_folder:
                self.__log_root_folder = log_root_folder
            else:
                self.__log_root_folder = RELATIVE_PATH_TO_FOLDER_OF_LOGS
            if not os.path.exists(self.__log_root_folder):
                os.mkdir(os.path.abspath(self.__log_root_folder))
            self.__full_path_to_log_file = os.path.join(
                os.path.abspath(self.__log_root_folder),
                self.__log_file_name
            )

            self.__file_handler = TimedRotatingFileHandler(
                self.__full_path_to_log_file,
                when='midnight'
            )
            self.__file_handler.setFormatter(self.__log_formatter)

            self.__logger.addHandler(self.__console_handler)
            self.__logger.addHandler(self.__file_handler)

        else:
            for handler in self.__logger.handlers:
                if isinstance(handler, logging.handlers.TimedRotatingFileHandler):
                    self.__file_handler = handler
                elif isinstance(handler, logging.StreamHandler):
                    self.__console_handler = handler

            self.__log_formatter = self.__console_handler.formatter

            self.__full_path_to_log_file = self.__file_handler.baseFilename
            self.__log_root_folder, self.__log_file_name = os.path.split(
                self.__file_handler.baseFilename
            )

        self.__logger.setLevel(logging.INFO)
        self.__logger.propagate = False

        self.info(
            'Logger.__init__',
            'Logger %s is using the file handler %s and the console handler %s.' %
            (self, self.__file_handler, self.__console_handler)
        )

        self.__latest_error_message_list = []
        self.__n_latest_error_messages = 0

    def info(self, place, message):
        """
        Generate an info message.
        :param str place:   Where the message is issued
        :param str message: Message
        """
        self.__logger.info(
            '%s - %s' %
            (
                place,
                message.replace('\n', ' ')
            )
        )

    def warning(self, place, message):
        """
        Generate a warning message.
        :param str place:   Where the message is issued
        :param str message: Message
        """
        self.__logger.warning(
            '%s - %s' %
            (
                place,
                message.replace('\n', ' ')
            )
        )

    def error(self, place, message):
        """
        Generate an error message.
        :param str place:   Where the message is issued
        :param str message: Message
        """
        self.__logger.error(
            '%s - %s' %
            (
                place,
                message.replace('\n', ' ')
            )
        )
        self.__n_latest_error_messages += 1
        self.__latest_error_message_list.append(
            message.replace('\n', ' ')
        )

    def get_latest_error_messages(self):
        """
        Get the latest error messages since the last call of this method.
        :rtype:     str
        """
        if self.__n_latest_error_messages == 0:
            print(self.__n_latest_error_messages)
        result = ''
        for idx in range(self.__n_latest_error_messages):
            result += '%d. %s ' % (idx+1, self.__latest_error_message_list[idx])
        return result

    def clear_latest_error_messages(self):
        self.__n_latest_error_messages = 0
        self.__latest_error_message_list.clear()
