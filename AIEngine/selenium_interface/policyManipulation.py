#!/usr/bin/env python
#
# Created by Xixuan on 29/08/2018
#
# The file contains
#   1. PolicyManipulationRPA
#

from selenium.common import exceptions as se

from selenium_interface.basicops import BaseRPA
from selenium_interface.settings import *


class PolicyManipulationRPA(BaseRPA):
    def __init__(self, driver=None, logger=None):
        """
        Create an interface to input data to MARC MY system.
        :param WebDriver driver:        Chrome driver instance
        :param Logger logger:           Logger instance
        """
        super(PolicyManipulationRPA, self).__init__(driver, logger)

    # policy creation

    def pc_create_policy(self, row):
        self.go_to(MARC_POLICY_ADD_ADDRESS)
        flags = list()
        for key in row.keys():
            self.pc_input_by_field_name(key, row[key])
        self.sleep(SLEEP_TIME)
        flags.append(
            self.pc_complete_inputting()
        )
        return all(flags)

    def pc_input_by_field_name(self, field_name, content=None):
        """
        Input content to record in Member Setup.
        :param str field_name:      Full name of the field
        :param str content:         Content to be filled
        :rtype:                     bool
        """
        return self.input_on_current_page_by_field_name(
            INPT_POLICY_CREATE_FIELD_NAME_TO_XPATH_MAPPER,
            INPT_POLICY_CREATE_FIELD_NAME_TO_TYPE_MAPPER,
            field_name,
            content,
        )

    def pc_complete_inputting(self):
        """
        Complete the updating of the member information.
        :rtype:     bool
        """
        try:
            elem = self.find_element_on_current_page_by_xpath(
                INPT_POLICY_CREATE_FIELD_NAME_TO_XPATH_MAPPER['save']
            )
            elem.click()
        except se.NoSuchElementException:
            self.logger.error(
                'RPA.pc_complete_inputting',
                'Completing policy creation failed: '
                'cannot find the save button.'
            )
            return False
        except se.WebDriverException as error:
            self.logger.error(
                'RPA.pc_complete_inputting',
                'Completing policy creation failed: %s.' %
                error
            )
            return False
        # the order of the message has an impact on the speed
        message_mapper = {
            'Validation Error: Value is required': False,
            'Policy already exist': False,
            'Record saved successfully': True
        }
        try:
            message, result = self.wait_until_visibility_of_message(message_mapper)
        except se.TimeoutException:
            self.logger.error(
                'RPA.pc_complete_inputting',
                'Completing policy creation failed: '
                'time is out after clicking on the save button.'
            )
            return False
        else:
            if result:
                self.logger.info(
                    'RPA.pc_complete_inputting',
                    'Completing policy creation is successful: %s.' %
                    message
                )
            else:
                self.logger.error(
                    'RPA.pc_complete_inputting',
                    'Completing policy creation failed: %s.' %
                    message
                )
            return result

    def pc_clear(self):
        """
        Clear the updated policy information.
        :rtype:     bool
        """
        elem = self.find_element_on_current_page_by_xpath(
            INPT_POLICY_CREATE_FIELD_NAME_TO_XPATH_MAPPER['clear']
        )
        elem.click()
        return True

    # policy manipulation

    def ps_input_by_field_name(self, field_name, content=None):
        """
        Input content in Policy Search.
        :param str field_name:      Full name of the field
        :param str content:         Content to be filled
        :rtype:                     bool
        """
        return self.input_on_current_page_by_field_name(
            INPT_POLICY_SEARCH_FIELD_NAME_TO_XPATH_MAPPER,
            INPT_POLICY_SEARCH_FIELD_NAME_TO_TYPE_MAPPER,
            field_name,
            content
        )

    def pm_search_policy(self, row):
        """
        Search for the policy indicated in the row.
        :param pd.Series row:       Row of data of a new policy
        :rtype:                     bool
        """
        self.driver.get(MARC_POLICY_SEARCH_ADDRESS)
        field_name_tuple = ('policy id', 'policy owner name', 'policy num', 'policy creation type', 'policy enabled')
        for field_name in field_name_tuple:
            if field_name in row:
                self.ps_input_by_field_name(field_name, row[field_name])
        self.click_on_button_on_current_page_by_xpath(
            INPT_POLICY_SEARCH_FIELD_NAME_TO_XPATH_MAPPER['search']
        )
        string_for_printing = (
            row.to_dict().__str__().replace('{', '').replace('}', '').replace("'", '')
        )
        try:
            elems = self.wait.until(
                lambda x: x.find_elements_by_xpath('//*[@title="Update Record"]')
            )
        except se.TimeoutException:
            self.logger.error(
                'RPA.pm_search_policy',
                'Searching for the policy %s failed: '
                'time is out and found no matched policies.' %
                string_for_printing
            )
            return False
        if len(elems) != 1:
            self.logger.error(
                'RPA.pm_search_policy',
                'Searching for the policy %s failed: '
                'found %d policies and cannot proceed.' %
                (
                    string_for_printing,
                    len(elems)
                )
            )
            return False
        else:
            elems[0].click()
            self.logger.info(
                'RPA.pm_search_policy',
                'Searching for the policy %s is successful.' %
                string_for_printing
            )
            return True

    def pm_input_by_field_name(self, field_name, content=None):
        """
        Input content to record in Member Setup.
        :param str field_name:  Full name of the field
        :param str content:     Content to be filled
        :rtype:                 bool
        """
        return self.input_on_current_page_by_field_name(
            INPT_POLICY_UPDATE_FIELD_NAME_TO_XPATH_MAPPER,
            INPT_POLICY_UPDATE_FIELD_NAME_TO_TYPE_MAPPER,
            field_name,
            content
        )

    def pm_complete_inputting(self):
        """
        Complete the updating of the member information.
        :rtype:                 bool
        """
        try:
            elem = self.find_element_on_current_page_by_xpath(
                INPT_POLICY_UPDATE_FIELD_NAME_TO_XPATH_MAPPER['update']
            )
            elem.click()
        except se.NoSuchElementException:
            self.logger.error(
                'RPA.pm_complete_inputting',
                'Completing policy setup failed: '
                'cannot find the update button.'
            )
            return False
        except se.WebDriverException as error:
            self.logger.error(
                'RPA.pm_complete_inputting',
                'Completing policy setup failed when '
                'trying to click on the update button: %s.' %
                error
            )
            return False
        # the order of the message has an impact on the speed
        message_mapper = {
            'Validation Error: Value is required': False,
            'Policy already exist': False,
            'Record updated successfully': True
        }
        try:
            message, result = self.wait_until_visibility_of_message(message_mapper)
        except se.TimeoutException:
            self.logger.error(
                'RPA.pm_complete_inputting',
                'Completing policy setup failed: '
                'time is out after clicking on the update button.'
            )
            return False
        else:
            if result:
                self.logger.info(
                    'RPA.pm_complete_inputting',
                    'Completing policy setup is successful: %s.' %
                    message
                )
            else:
                self.logger.error(
                    'RPA.pm_complete_inputting',
                    'Completing policy setup failed: %s.' %
                    message
                )
            return result

    def pm_close_lookup(self):
        """
        Close the lookup window that pops up.
        :rtype:     bool
        """
        close_button_list = [
            button for button in
            self.driver.find_elements_by_xpath(
                "//a["
                "@class='ui-dialog-titlebar-icon ui-dialog-titlebar-close ui-corner-all' "
                "and "
                "@role='button'"
                "]"
            )
            if button.is_displayed()
        ]
        try:
            close_button_list[0].click()
            return True
        except IndexError:
            self.logger.error(
                'RPA.pm_close_lookup',
                'Cannot close the lookup window: '
                'cannot find the close button.'
            )
            return False
        except se.WebDriverException as error:
            self.logger.error(
                'RPA.pm_close_lookup',
                'Cannot close the lookup window: %s.' %
                error
            )
            return False

    def pm_attach_plan(self, row):
        """
        Attach a plan to a policy contained in a row.
        :param pd.Series row:       Row of information for attaching plan
        :rtype:                     bool
        """
        flags = list()
        flags.append(
            self.pm_search_policy(row)
        )
        if not all(flags):
            return False
        for key in [item for item in row.keys() if 'plan' not in item]:
            self.pm_input_by_field_name(key, row[key])

        self.pm_input_by_field_name('attach plan lookup')
        self.wait_until_revealed_of_element_by_xpath(
            INPT_POLICY_UPDATE_FIELD_NAME_TO_XPATH_MAPPER['int plan code (id)']
            +
            '/../../../../../../../..'
        )

        self.pm_input_by_field_name('int plan code (id)', '')
        for key in [item for item in row.keys() if 'plan' in item]:
            self.pm_input_by_field_name(key, row[key])
        flags.append(
            self.pm_input_by_field_name('search')
        )

        self.wait_until_hidden_of_element_by_xpath(
            LOADING_INDICATOR
        )

        string_for_printing = (
            row.to_dict().__str__().replace('{', '').replace('}', '').replace("'", '')
        )
        try:
            # because the id has a special name and is not likely to change
            total_rows = self.wait.until(
                lambda x: x.find_elements_by_xpath(
                    "//a[contains(@id, 'inptPlanResultList')]"
                )
            )
        except se.TimeoutException:
            try:
                self.wait.until(
                    lambda x: x.find_elements_by_xpath(
                        "//*[contains(text(), 'No records found.')]"
                    )
                )
            except se.TimeoutException:
                self.logger.error(
                    'RPA.pm_attach_plan',
                    'Attaching plan for %s failed: '
                    'time is out when searching for plans.' %
                    string_for_printing
                )
                flag = False
                for _ in range(3):
                    flag = self.pm_close_lookup()
                    if flag:
                        break
                flags.append(
                    flag
                )
            else:
                self.logger.error(
                    'RPA.pm_attach_plan',
                    'Attaching plan for %s failed: '
                    'no matched plan found.' %
                    string_for_printing
                )
                flag = False
                flags.append(
                    flag
                )
                # try to close the lookup window
                for _ in range(3):
                    flag = self.pm_close_lookup()
                    if flag:
                        break
                flags.append(
                    flag
                )
        else:
            total_rows[0].click()
            self.wait_until_hidden_of_element_by_xpath(
                INPT_POLICY_UPDATE_FIELD_NAME_TO_XPATH_MAPPER['int plan code (id)']
                +
                '/../../../../../../../..'
            )
            flags.append(
                self.pm_input_by_field_name('add')
            )
            for row in total_rows[1:]:
                self.pm_input_by_field_name('attach plan lookup')
                self.wait_until_revealed_of_element_by_xpath(
                    INPT_POLICY_UPDATE_FIELD_NAME_TO_XPATH_MAPPER['int plan code (id)']
                    +
                    '/../../../../../../../..'
                )
                row.click()
                self.wait_until_hidden_of_element_by_xpath(
                    INPT_POLICY_UPDATE_FIELD_NAME_TO_XPATH_MAPPER['int plan code (id)']
                    +
                    '/../../../../../../../..'
                )
                flags.append(
                    self.pm_input_by_field_name('add')
                )
            self.wait_until_hidden_of_element_by_xpath(
                LOADING_INDICATOR
            )
            flags.append(
                self.pm_complete_inputting()
            )
        return all(flags)

    def pm_move_policy(self, row):
        """
        Move a policy indicated by a row.
        :param pd.Series row:       Row of information for moving a policy
        :rtype:                     bool
        """
        flags = list()
        flags.append(
            self.pm_search_policy(row)
        )
        if not all(flags):
            return False
        for key in [item for item in row.keys() if 'move' in item]:
            self.pm_input_by_field_name(key, row[key])
        self.pm_input_by_field_name('add movement')
        message_mapper = {
            'Action Start Date required': False,
            'Action End Date required': False,
            'Action Start and End Date required': False
        }
        try:
            message, result = self.wait_until_visibility_of_message(message_mapper)
        except se.TimeoutException:
            flags.append(
                self.pm_complete_inputting()
            )
            return all(flags)
        else:
            self.logger.error(
                'PolicyManipulationRPA.pm_move_policy',
                'Moving policy failed: %s' %
                message
            )
            return result

    def pm_manipulate_policy(self, loader):
        """
        Manipulate policies according the table contained in PolicyManipulationXLSXFileLoader.
        :param PolicyManipulationXLSXFileLoader loader:     Loader of the XLSX file
        :rtype:                                             boot
        """
        flags = list()
        if 'new' in loader.action_flag_series_dict:
            for idx, row in loader.get_row_iterator(action='new'):
                self.logger.clear_latest_error_messages()
                flag = self.pc_create_policy(row)
                flags.append(flag)
                if flag:
                    loader.append_row_to_output_table(row, 'good')
                    self.logger.info(
                        'PolicyManipulationRPA.pm_manipulate_policy',
                        'No. %d of row in %s is "new" and is processed successfully.' %
                        (idx + 1, loader)
                    )
                else:
                    loader.append_row_to_output_table(
                        row,
                        'bad',
                        self.logger.get_latest_error_messages()
                    )
                    self.logger.error(
                        'PolicyManipulationRPA.pm_manipulate_policy',
                        'No. %d of row in %s is "new" and is processed unsuccessfully.' %
                        (idx + 1, loader)
                    )
        if 'attach plan' in loader.action_flag_series_dict:
            for idx, row in loader.get_row_iterator(action='attach plan'):
                self.logger.clear_latest_error_messages()
                flag = self.pm_attach_plan(row)
                flags.append(flag)
                if flag:
                    loader.append_row_to_output_table(row, 'good')
                    self.logger.info(
                        'PolicyManipulationRPA.pm_manipulate_policy',
                        'No. %d of row in %s is "attach plan" and is processed successfully.' %
                        (idx + 1, loader)
                    )
                else:
                    loader.append_row_to_output_table(
                        row,
                        'bad',
                        self.logger.get_latest_error_messages()
                    )
                    self.logger.error(
                        'PolicyManipulationRPA.pm_manipulate_policy',
                        'No. %d of row in %s is "attach plan" and is processed unsuccessfully.' %
                        (idx + 1, loader)
                    )
        other_actions = (
            set(loader.action_flag_series_dict.keys()).difference(['new', 'attach plan'])
        )
        for action in other_actions:
            for idx, row in loader.get_row_iterator(action=action):
                self.logger.clear_latest_error_messages()
                flag = self.pm_move_policy(row)
                flags.append(flag)
                if flag:
                    loader.append_row_to_output_table(row, 'good')
                    self.logger.info(
                        'PolicyManipulationRPA.pm_manipulate_policy',
                        'No. %d of row in %s is "%s" and is processed successfully.' %
                        (idx + 1, loader, action)
                    )
                else:
                    loader.append_row_to_output_table(
                        row,
                        'bad',
                        self.logger.get_latest_error_messages()
                    )
                    self.logger.error(
                        'PolicyManipulationRPA.pm_manipulate_policy',
                        'No. %d of row in %s is "%s" and is processed unsuccessfully.' %
                        (idx + 1, loader, action)
                    )
        if sum(flags) > 0:
            self.logger.info(
                'PolicyManipulationRPA.pm_manipulate_policy',
                '%d among %d rows in %s are processed successfully. '
                'Send to %s.' %
                (
                    sum(flags),
                    loader.total_n_rows,
                    loader,
                    loader.full_path_to_output_file_of_good_rows
                )
            )
        if len(flags) - sum(flags) > 0:
            self.logger.error(
                'PolicyManipulationRPA.pm_manipulate_policy',
                '%d among %d rows in %s are processed unsuccessfully. '
                'Send to %s.' %
                (
                    len(flags) - sum(flags),
                    loader.total_n_rows,
                    loader,
                    loader.full_path_to_output_file_of_bad_rows
                )
            )
        self.go_to(MARC_DASHBOARD_ADDRESS)
        return all(flags)
