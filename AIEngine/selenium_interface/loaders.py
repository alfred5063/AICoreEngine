#!/usr/bin/env python
#
# Created by Xixuan on 29/08/2018
#
# The file contains
#   1. BaseFileLoader
#   2. PolicyManipulationXLSXFileLoader
#   3. MemberSetupXLSXFileLoader
#

import os
import numpy as np
import pandas as pd

from numbers import Number

from selenium_interface.tools import ct
from selenium_interface.logger import Logger
from selenium_interface.settings import *
from utils.logging import logging
from utils.audit_trail import audit_log
# TODO: add unfinished rows in addition to good rows and bad rows

class BaseFileLoader(object):
    def __init__(self, relative_input_folder, relative_output_folder, input_file_name, logger=None):
        """
        Create a XLSX file loader providing accesses to the data files.
        :param str input_file_name:     Name of the XLSX file
        :param Logger logger:           Logger instance
        """
        # for input files
        self.__input_file_name = input_file_name
        abs_path_to_folder_of_input_files = os.path.abspath(
            relative_input_folder
        )
        self.full_path_to_input_file = os.path.join(
            abs_path_to_folder_of_input_files,
            self.__input_file_name
        )
        # for output files
        self.__output_file_name_for_good_rows = (
                OUTPUT_FILE_PREFIX
                +
                'good_rows_%s_from_' % ct(plain=True)
                +
                self.__input_file_name
        )
        self.__output_file_name_for_bad_rows = (
                OUTPUT_FILE_PREFIX
                +
                'bad_rows_%s_from_' % ct(plain=True)
                +
                self.__input_file_name
        )
        self.__output_file_name_for_ignored_rows = (
                OUTPUT_FILE_PREFIX
                +
                'ignored_rows_%s_from_' % ct(plain=True)
                +
                self.__input_file_name
        )
        self.__output_file_name_for_original_table = (
                OUTPUT_FILE_PREFIX
                +
                'directly_returned_table_%s_from_' % ct(plain=True)
                +
                self.__input_file_name
        )
        abs_path_to_folder_of_output_files = os.path.abspath(
            relative_output_folder
        )
        self.full_path_to_output_file_of_good_rows = os.path.join(
            abs_path_to_folder_of_output_files,
            self.__output_file_name_for_good_rows
        )
        self.full_path_to_output_file_of_bad_rows = os.path.join(
            abs_path_to_folder_of_output_files,
            self.__output_file_name_for_bad_rows
        )
        self.full_path_to_output_file_of_ignored_rows = os.path.join(
            abs_path_to_folder_of_output_files,
            self.__output_file_name_for_ignored_rows
        )
        self.full_path_to_output_file_of_original_table = os.path.join(
            abs_path_to_folder_of_output_files,
            self.__output_file_name_for_original_table
        )
        if not os.path.exists(abs_path_to_folder_of_output_files):
            os.mkdir(abs_path_to_folder_of_output_files)
        # for the logger
        if logger:
            if isinstance(logger, Logger):
                self.logger = logger
                self.logger.info(
                    'BaseFileLoader.__init__',
                    'Used the logger %s.' %
                    self.logger
                )
            else:
                self.logger.error(
                    'BaseFileLoader.__init__',
                    'logger accepts an instance of %s, but got %s' %
                    (type(Logger), type(logger))
                )
                raise Exception(
                    'logger accepts an instance of %s, but got %s' %
                    (type(Logger), type(logger))
                )
        else:
            self.logger = Logger()
            self.logger.info(
                'BaseFileLoader.__init__',
                'Loaded the logger %s.' %
                self.logger
            )

        self.logger.info(
            'BaseFileLoader.__init__',
            'Loading XLSX file %s.' %
            self.full_path_to_input_file
        )

        self.good_flags = list()
        for idx in range(MAX_N_TRIES_TO_FIND_HEADER):
            self.table = pd.read_excel(
                self.full_path_to_input_file,
                header=idx
            )
            if KEY_WORD_TO_FIND_HEADER in [item.lower() for item in self.table.keys()]:
                self.good_flags.append(True)
                break
            if idx == MAX_N_TRIES_TO_FIND_HEADER - 1:
                self.logger.error(
                    'BaseFileLoader.__init__',
                    'Failed to load XLSX file %s: cannot identify the header row among first %d rows based on %s' %
                    (
                        self.full_path_to_input_file,
                        MAX_N_TRIES_TO_FIND_HEADER,
                        KEY_WORD_TO_FIND_HEADER
                    )
                )
                self.good_flags.append(False)

        self.logger.info(
            'BaseFileLoader.__init__',
            'Successfully loaded XLSX file %s.' %
            self.full_path_to_input_file
        )
        # some public variables
        self.flags = list()
        self.current_row_iterator = None
        self.temp_dict_list_good_rows = list()
        self.temp_dict_list_bad_rows = list()
        self.total_n_rows = len(self.table)
        self.layout = list()

    def __repr__(self):
        return (
                '<%s of %s at %s>' %
                (
                    self.__class__.__name__,
                    self.full_path_to_input_file,
                    hex(id(self))
                )
        )

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.good():
            if exc_val is not None:
                self.logger.error(
                    'BaseFileLoader.__exit__',
                    'Meet a problem %s when dealing with the data in the loader of %s: %s' %
                    (exc_type, self.full_path_to_input_file, exc_val)
                )
                self.finish(self.layout, directly_return_original_table=True)
            else:
                self.finish(self.layout, )
        else:
            self.logger.error(
                'BaseFileLoader.good',
                'Got problems when preprocessing the input file %s. '
                'Will return it to %s.' %
                (
                    self.full_path_to_input_file,
                    self.full_path_to_output_file_of_bad_rows
                )
            )
            self.finish(self.layout, directly_return_original_table=True)
        return False

    def __iter__(self):
        return self

    def __next__(self):
        if self.current_row_iterator is not None:
            idx, row = self.current_row_iterator.__next__()
            return idx, row[row.notna()]
        else:
            self.logger.warning(
                'BaseFileLoader.next',
                'Did not attached a row generator.'
            )
            raise StopIteration(
                'Did not attached a row generator.'
            )

    def good(self):
        """
        Returns whether the input date file is good or not.
        :rtype:     bool
        """
        return all(self.good_flags)

    def decorate_table(self, equivalence_mapper, cols_to_be_lowercased, cols_to_be_dropped, all_to_int):
        """
        Decorate the data by
            1. changing equivalent names
            2. lowercasing all column names
            3. lowercasing chosen columns
            4. dropping irrelevant columns
            5. converting all numbers to int

        :param dict equivalence_mapper:         Equivalent col names
        :param list cols_to_be_lowercased:      Columns to be lowercased
        :param list cols_to_be_dropped:         Columns to be dropped
        :param bool all_to_int:                 Whether or not to convert all numbers to int
        :rtype:                                 bool
        """
        try:
            # 1 and 2
            new_name_mapper = dict()
            for key in self.table.keys():
                flag = True
                for equiv_keys in equivalence_mapper:
                    if key.lower() in [item.lower() for item in equiv_keys]:
                        flag = False
                        new_name_mapper[key] = equivalence_mapper[equiv_keys].lower()
                if flag:
                    new_name_mapper[key] = key.lower()
            self.table.rename(columns=new_name_mapper, inplace=True)
            # 3
            for col_name in cols_to_be_lowercased:
                if col_name.lower() in self.table.keys():
                    self.table[col_name.lower()] = self.table[col_name.lower()].str.lower()
            # 4
            for col_name in cols_to_be_dropped:
                if col_name.lower() in self.table.keys():
                    self.table.drop(col_name.lower(), axis=1, inplace=True)
            # 5
            if all_to_int:
                # TODO: applymap might be pretty slow
                self.table = self.table.applymap(
                    lambda x: int(x) if isinstance(x, Number) and not np.isnan(x) else x
                )
        except Exception as error:
            logging('decorate_table', error)
            self.logger.error(
                'BaseFileLoader.decorate_table',
                'Got an error when decorating the table: %s' %
                error
            )
            return False
        else:
            return True

    def get_row_iterator(self, **kwargs):
        """
        Return self as an iterator for all of the rows.
        :rtype:     BaseFileLoader
        """
        self.current_row_iterator = self.table.iterrows()
        return self

    def append_row_to_output_table(self, row, good_or_bad, reason=None):
        """
        Append problematic rows to the output file.
        :param pd.Series or dict row:       Problematic row
        :param str good_or_bad:             'good' or 'bad'
        :param str reason:                  Reason why it is good or bad
        """
        if isinstance(row, pd.Series):
            if reason is not None:
                row['comments'] = reason
            if good_or_bad == 'good':
                self.temp_dict_list_good_rows.append(row.to_dict())
            elif good_or_bad == 'bad':
                self.temp_dict_list_bad_rows.append(row.to_dict())
            else:
                self.temp_dict_list_bad_rows.append(row.to_dict())
                self.logger.error(
                    'BaseFileLoader.append_row_to_output_table',
                    'Specified wrong good_or_bad %s for %s. Put it to bad.' %
                    (good_or_bad, row.__str__().replace('\n', ''))
                )
        elif isinstance(row, dict):
            if reason is not None:
                row['comments'] = reason
            if good_or_bad == 'good':
                self.temp_dict_list_good_rows.append(row)
            elif good_or_bad == 'bad':
                self.temp_dict_list_bad_rows.append(row)
            else:
                self.logger.error(
                    'BaseFileLoader.append_row_to_output_table',
                    'Specified wrong good_or_bad %s for %s. Put it to bad.' %
                    (good_or_bad, row.__str__().replace('\n', ''))
                )
                raise ValueError(
                    'BaseFileLoader.append_row_to_output_table',
                    'Specified wrong good_or_bad %s for %s. Put it to bad.' %
                    (good_or_bad, row.__str__().replace('\n', ''))
                )
        else:
            self.logger.error(
                'BaseFileLoader.append_row_to_output_table',
                'Loader only accepts dict or list for append to the output table, '
                'but got %s for %s.' %
                (type(row), row.__str__().replace('\n', ''))
            )
            raise ValueError(
                'BaseFileLoader.append_row_to_output_table',
                'Loader only accepts dict or list for append to the output table, '
                'but got %s for %s.' %
                (type(row), row.__str__().replace('\n', ''))
            )

    def finish(self, layout, directly_return_original_table=False):
        """
        Finish the data file by deleting it and return good and bad rows to another folder.
        :param list layout:                                            Column layouts
        :param bool directly_return_original_table:     Directly return the original data if necessary
        """
        flags = []
        if len(self.temp_dict_list_good_rows) > 0:
            try:
                with pd.ExcelWriter(self.full_path_to_output_file_of_good_rows) as writer:
                    output_table = pd.DataFrame(
                        data = self.temp_dict_list_good_rows,
                        columns = layout
                    )
                    output_table.to_excel(writer)
            except Exception as error:
                audit_log('finish', error)
                self.logger.error(
                    'BaseFileLoader.finish',
                    'Got an error when saving good rows to %s: %s' %
                    (
                        self.full_path_to_output_file_of_good_rows,
                        error
                    )
                )
                flags.append(False)
            else:
                self.logger.info(
                    'BaseFileLoader.finish',
                    'Saved good rows to %s.' %
                    (
                        self.full_path_to_output_file_of_good_rows
                    )
                )
                flags.append(True)
        if len(self.temp_dict_list_bad_rows) > 0:
            try:
                with pd.ExcelWriter(self.full_path_to_output_file_of_bad_rows) as writer:
                    output_table = pd.DataFrame(
                        data = self.temp_dict_list_bad_rows,
                        columns = layout
                    )
                    output_table.to_excel(writer)
            except Exception as error:
                logging('finish: if len for temp dict list is bad rows', error)
                self.logger.error(
                    'BaseFileLoader.finish',
                    'Got an error when saving bad rows to %s: %s' %
                    (
                        self.full_path_to_output_file_of_bad_rows,
                        error
                    )
                )
                flags.append(False)
            else:
                self.logger.info(
                    'BaseFileLoader.finish',
                    'Saved bad rows to %s.' %
                    (
                        self.full_path_to_output_file_of_bad_rows
                    )
                )
                flags.append(True)
        if directly_return_original_table:
            try:
                with pd.ExcelWriter(self.full_path_to_output_file_of_original_table) as writer:
                    self.table.to_excel(writer)
            except Exception as error:
                logging('finish: directly_return_original_table', error)
                self.logger.error(
                    'BaseFileLoader.finish',
                    'Got an error when saving the original data table to %s: %s' %
                    (
                        self.full_path_to_output_file_of_bad_rows,
                        error
                    )
                )
                flags.append(False)
            else:
                self.logger.info(
                    'BaseFileLoader.finish',
                    'Directly returned the original data table to %s' %
                    self.full_path_to_output_file_of_bad_rows
                )
                flags.append(True)
        try:
            os.remove(self.full_path_to_input_file)
        except Exception as error:
            logging('finish: Tried to remove the original input file', error)
            self.logger.error(
                'BaseFileLoader.finish',
                'Tried to remove the original input file %s. '
                'But got an error: %s' %
                (
                    self.full_path_to_input_file,
                    error
                )
            )
        else:
            self.logger.info(
                'BaseFileLoader.finish',
                'Removed the original input file %s. ' %
                (
                    self.full_path_to_input_file
                )
            )
        return all(flags)


class PolicyManipulationXLSXFileLoader(BaseFileLoader):
    def __init__(self, input_file_name, logger=None,
                 input_file_path=RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_POLICY_MANIPULATION):
        """
        Create a XLSX file loader providing accesses to the data files.
        :param str input_file_name:     Name of the XLSX file for policy manipulation
        :param Logger logger:           Logger instance
        """
        super(PolicyManipulationXLSXFileLoader, self).__init__(
            input_file_path,
            RELATIVE_PATH_TO_FOLDER_OF_OUTPUT_XLSX_FOR_POLICY_MANIPULATION,
            input_file_name,
            logger
        )

        self.action_flag_series_dict = dict()

        self.good_flags.append(
            self.decorate_table(
                INPT_POLICY_EQUIVALENCE_MAPPER,
                COLS_TO_BE_LOWERCASED_IN_INPUT_XLSX_FOR_PM,
                COLS_TO_BE_DROPPED_IN_INPUT_XLSX_FOR_PM,
                True
            )
        )
        self.good_flags.append(
            self.list_contained_actions()
        )
        self.layout = POLICY_HANDSHAKE_LAYOUT

    def list_contained_actions(self):
        """
        List the actions contained in the data file for policy manipulation.
        :rtype:     bool
        """
        try:
            # new policy
            new_policy_flag_series = self.table['policy creation type'].notna()
            if new_policy_flag_series.sum() > 0:
                self.action_flag_series_dict['new'] = new_policy_flag_series
            # attach plan
            attach_plan_flag_series = None
            for key in self.table.keys():
                if 'plan' in key:
                    if attach_plan_flag_series is None:
                        attach_plan_flag_series = self.table[key].notna()
                    else:
                        attach_plan_flag_series = (
                                attach_plan_flag_series
                                |
                                self.table[key].notna()
                        )
            if attach_plan_flag_series is not None:
                self.action_flag_series_dict['attach plan'] = attach_plan_flag_series
            # policy movements
            contained_movements = set(
                item.lower()
                for item in
                self.table['move type'][self.table['move type'].notna()]
            )
            for movement in contained_movements:
                self.action_flag_series_dict[movement] = (
                        self.table['move type'] == movement
                )
            # unknown rows
            if len(self.action_flag_series_dict) == 0:
                raise Exception(
                    'Cannot identify any actions in the file %s.' %
                    self.full_path_to_input_file
                )
            else:
                ignored_row_flag_series = None
                for key in self.action_flag_series_dict:
                    if ignored_row_flag_series is None:
                        ignored_row_flag_series = self.action_flag_series_dict[key].astype('int')
                    else:
                        ignored_row_flag_series = (
                                ignored_row_flag_series
                                +
                                self.action_flag_series_dict[key].astype('int')
                        )
                ignored_row_flag_series = (
                        ignored_row_flag_series != 1
                )
                accepted_row_flag_series = ~ignored_row_flag_series
                if ignored_row_flag_series.sum() > 0:
                    with pd.ExcelWriter(self.full_path_to_output_file_of_ignored_rows) as writer:
                        self.table[ignored_row_flag_series].to_excel(writer)
                    self.logger.error(
                        'PolicyManipulationXLSXFileLoader.list_contained_actions',
                        'Ignored %d rows that are not identified to have a single action.'
                        'Saved the ignored rows to %s.' %
                        (
                            ignored_row_flag_series.sum(),
                            self.full_path_to_output_file_of_ignored_rows
                        )
                    )
                for key in self.action_flag_series_dict:
                    self.action_flag_series_dict[key] = (
                            self.action_flag_series_dict[key]
                            &
                            accepted_row_flag_series
                    )
        except Exception as error:
            logging('list_contained_actions', error)
            self.logger.error(
                'PolicyManipulationXLSXFileLoader.list_contained_actions',
                'Got an error when extracting actions contained in the table: %s.' %
                error
            )
            return False
        else:
            return True

    def get_row_iterator(self, **kwargs):
        """
        Return self as an iterator for the rows of the action.
        Should provide a keyword argument 'action' indicating
        :rtype:     PolicyManipulationXLSXFileLoader
        """
        if 'action' not in kwargs:
            self.logger.error(
                'PolicyManipulationXLSXFileLoader.get_row_iterator',
                "Miss the argument 'action'."
            )
            raise ValueError(
                "Miss the argument 'action'."
            )
        else:
            action = kwargs['action']
        if action in self.action_flag_series_dict:
            self.current_row_iterator = (
                self.table[
                    self.action_flag_series_dict[action]
                ].iterrows()
            )
        else:
            self.current_row_iterator = None
            self.logger.error(
                'PolicyManipulationXLSXFileLoader.get_row_iterator',
                'Got an action %s that is not contained in %s. '
                'Do not attach a row iterator.' %
                (
                    action,
                    ', '.join(self.action_flag_series_dict.keys())
                )
            )
        self.logger.info(
            'PolicyManipulationXLSXFileLoader.get_row_iterator',
            'Obtained the row generator for %s.' %
            action
        )
        return self


class MemberSetupXLSXFileLoader(BaseFileLoader):
    def __init__(self, input_file_name, logger=None,
                 input_file_path = RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_MEMBER_SETUP):
        """
        Create a XLSX file loader providing accesses to the data files.
        :param str input_file_name:     Name of the XLSX file for member setup
        :param Logger logger:           Logger instance
        """
        super(MemberSetupXLSXFileLoader, self).__init__(
            input_file_path,
            RELATIVE_PATH_TO_FOLDER_OF_OUTPUT_XLSX_FOR_MEMBER_SETUP,
            input_file_name,
            logger
        )

        self.action_flag_series_dict = dict()

        self.good_flags.append(
            self.decorate_table(
                MEMBER_SETUP_EQUIVALENCE_MAPPER,
                COLS_TO_BE_LOWERCASED_IN_INPUT_XLSX_FOR_MS,
                COLS_TO_BE_DROPPED_IN_INPUT_XLSX_FOR_MS,
                True
            )
        )
        self.good_flags.append(
            self.list_contained_actions()
        )
        self.layout = MEMBER_HANDSHAKE_LAYOUT

    def list_contained_actions(self):
        """
        List the actions contained in the data file for member setup.
        :rtype:     bool
        """
        try:
            # known rows
            keep_series = self.table['action'] != '<empty>'
            all_action_list_series = (
                (self.table['status'] + ';' + self.table['action']).map(
                    lambda x: [x.split(';')[0] + '::' + item for item in x.split(';')[1:]]
                )
            )
            action_list_series = all_action_list_series[keep_series]
            unique_action_set = set()
            for action_list in action_list_series:
                unique_action_set.update(action_list)
            for unique_action in unique_action_set:
                self.action_flag_series_dict[unique_action] = (
                    all_action_list_series.map(lambda x: unique_action in x)
                )

            # unknown rows
            if len(self.action_flag_series_dict) == 0:
                raise Exception(
                    'Cannot identify any actions in the file %s.' %
                    self.full_path_to_input_file
                )
            else:
                ignored_row_flag_series = None
                for key in self.action_flag_series_dict:
                    if ignored_row_flag_series is None:
                        ignored_row_flag_series = self.action_flag_series_dict[key].astype('int')
                    else:
                        ignored_row_flag_series = (
                                ignored_row_flag_series
                                +
                                self.action_flag_series_dict[key].astype('int')
                        )
                ignored_row_flag_series = (
                        ignored_row_flag_series < 1
                )
                accepted_row_flag_series = ~ignored_row_flag_series
                if ignored_row_flag_series.sum() > 0:
                    with pd.ExcelWriter(self.full_path_to_output_file_of_ignored_rows) as writer:
                        self.table[ignored_row_flag_series].to_excel(writer)
                    self.logger.error(
                        'MemberSetupXLSXFileLoader.list_contained_actions',
                        'Ignored %d rows that are not identified to have a single action.'
                        'Saved the ignored rows to %s.' %
                        (
                            ignored_row_flag_series.sum(),
                            self.full_path_to_output_file_of_ignored_rows
                        )
                    )
                for key in self.action_flag_series_dict:
                    self.action_flag_series_dict[key] = (
                            self.action_flag_series_dict[key]
                            &
                            accepted_row_flag_series
                    )
        except Exception as error:
            logging('list_contained_actions', error)
            self.logger.error(
                'MemberSetupXLSXFileLoader.list_contained_actions',
                'Got an error when extracting actions contained in the table: %s.' %
                error
            )
            return False
        else:
            return True

    def get_row_iterator(self, **kwargs):
        """
        Return self as an iterator for all of the rows.
        :rtype:     MemberSetupXLSXFileLoader
        """
        if 'action' not in kwargs:
            self.logger.error(
                'MemberSetupXLSXFileLoader.get_row_iterator',
                "Miss the argument 'action'."
            )
            raise ValueError(
                "Miss the argument 'action'."
            )
        else:
            action = kwargs['action']
        if action in self.action_flag_series_dict:
            self.current_row_iterator = (
                self.table[
                    self.action_flag_series_dict[action]
                ].iterrows()
            )
        else:
            self.current_row_iterator = None
            self.logger.error(
                'MemberSetupXLSXFileLoader.get_row_iterator',
                'Got an action %s that is not contained in %s. '
                'Do not attach a row iterator.' %
                (
                    action,
                    ', '.join(self.action_flag_series_dict.keys())
                )
            )
        self.logger.info(
            'MemberSetupXLSXFileLoader.get_row_iterator',
            'Obtained the row generator for %s.' %
            action
        )
        return self
