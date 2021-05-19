#!/usr/bin/env python
#
# Created by Xixuan on 29/08/2018
#
# The file contains
#   1.  MemberSetupRPA
#
import pandas as pd
from selenium.common import exceptions as se
from selenium_interface.basicops import BaseRPA
from selenium_interface.settings import *
from datetime import datetime
from time import sleep
from output.get_action_codes_from_db import convert_action_code_to_actions
from utils.logging import logging
from loading.excel.createexcel import create_excel
from regex import sub, findall
import os
import numpy as np
from utils.audit_trail import audit_log, update_audit
from analytics.marc_analytics import analytic_dm_by_row

class MemberSetupRPA(BaseRPA):
    def __init__(self, driver=None, logger=None):
        """
        Create an interface to input data to MARC MY system.
        :param WebDriver driver:        Chrome driver instance
        :param Logger logger:           Logger instance
        """
        super(MemberSetupRPA, self).__init__(driver, logger)

    def ms_search_plan_rpa(self, where, row, method_name, search_field_list, button, base):
        """
        Search for the inpt member policy indicated in the row.
        :param str where:                   'member search' or 'inpt member policy search'
        :param pd.Series row:               Row of data of a new policy
        :param str method_name:             'primary' or 'secondary'
        :param list search_field_list:      List of fields for searching the member
        :param str button:                  'View Record' or 'Edit Record'
        :rtype:                             bool
        """
        audit_log('Search Plan RPA', 'Finding plan record.', base)
        flags = list()
        self.go_to(MARC_INPT_PLAN_SEARCH_ADDRESS)
        for field in search_field_list:
            if field in row:
                flags.append(self.ms_input_content_by_field_name(where,
                        field,
                        row[field]))

        flags.append(self.ms_input_content_by_field_name(where,
                'search'))

        string_for_printing = (row.__str__().replace('{', '').replace('}', '').replace("'", ''))
        try:
            self.wait_until_hidden_of_element_by_xpath(LOADING_INDICATOR)
            elems = self.driver.find_elements_by_xpath("//*[@title='%s']" % button)
            #elems = self.wait.until(
            #    lambda x: x.find_elements_by_xpath("//*[@title='%s']" %
            #    button)
            #)
        except se.TimeoutException:
            logging('ms_search_plan_rpa', 'time is out and found no matched plan.', base)
            self.logger.error('RPA.ms_search_plan_rpa',
                'Searching for the plan %s using the %s way failed: '
                'time is out and found no matched plan.' % (string_for_printing,
                    method_name))
            return False
        else:
            if len(elems) != 1:
                self.logger.error('RPA.ms_search_plan_rpa',
                    'Searching for the plan %s using the %s way failed: '
                    'found %d matched plan and do not know which one to proceed.' % (string_for_printing,
                        method_name,
                        len(elems)))
                return False
            else:
                elems[0].click()
                self.logger.info('RPA.ms_search_plan_rpa',
                    'Searching for the plan %s using the %s way is successful.' % (string_for_printing,
                        method_name))
                return True
      

    def ms_search_member_rpa(self, where, row, method_name, search_field_list, button, base):
        """
        Search for the inpt member policy indicated in the row.
        :param str where:                   'member search' or 'inpt member policy search'
        :param pd.Series row:               Row of data of a new policy
        :param str method_name:             'primary' or 'secondary'
        :param list search_field_list:      List of fields for searching the member
        :param str button:                  'View Record' or 'Edit Record'
        :rtype:                             bool
        """
        audit_log('Search Member RPA', 'Finding member record.', base)
        flags = list()
        
        self.go_to(MARC_MEMBER_SEARCH_ADDRESS)
        for field in search_field_list:
            if field in row:
                flags.append(self.ms_input_content_by_field_name(where,
                        field,
                        base,
                        row[field]))

        flags.append(self.ms_input_content_by_field_name(where,
                'search', base))

        string_for_printing = (row.__str__().replace('{', '').replace('}', '').replace("'", ''))
        try:
            self.wait_until_hidden_of_element_by_xpath(LOADING_INDICATOR)
            elems = self.driver.find_elements_by_xpath("//*[@title='%s']" % button)
            #elems = self.wait.until(
            #    lambda x: x.find_elements_by_xpath("//*[@title='%s']" %
            #    button)
            #)
        except se.TimeoutException:
            logging('ms_search_member', 'time is out and found no matched members.', base)
            self.logger.error('RPA.ms_search_member',
                'Searching for the member %s using the %s way failed: '
                'time is out and found no matched members.' % (string_for_printing,
                    method_name))
            return False
        else:
            if len(elems) != 1:
                self.logger.error('RPA.ms_search_member',
                    'Searching for the member %s using the %s way failed: '
                    'found %d matched members and do not know which one to proceed.' % (string_for_printing,
                        method_name,
                        len(elems)))
                return False
            else:
                elems[0].click()
                self.logger.info('RPA.pm_search_policy',
                    'Searching for the member %s using the %s way is successful.' % (string_for_printing,
                        method_name))
                return True

    def ms_search_member(self, where, row, method_name, search_field_list, button, base):
        """
        Search for the inpt member policy indicated in the row.
        :param str where:                   'member search' or 'inpt member policy search'
        :param pd.Series row:               Row of data of a new policy
        :param str method_name:             'primary' or 'secondary'
        :param list search_field_list:      List of fields for searching the member
        :param str button:                  'View Record' or 'Edit Record'
        :rtype:                             bool
        """
        audit_log('Search Member', 'Finding member record.', base)
        flags = list()
        if where == 'member search':
            self.go_to(MARC_MEMBER_SEARCH_ADDRESS)
        elif where == 'inpt member policy search':
            self.go_to(MARC_INPT_MEMBER_POLICY_SEARCH_ADDRESS)
        else:
            raise ValueError('Got a wrong where %s. Shoule be chosen from %s.' % (where,
                    ', '.join(['member search',
                            'inpt member policy search'])))
       
        for field in search_field_list:
            
            if field == 'nric':
              #print('search field value: {0}, search field type: {1}, search field: {2}'.format(row[field], type(row[field]), field))
              if pd.isnull(row[field]):
                field = 'otheric'
            elif field == 'plan expiry date':
               try:
                  print('search field value: {0}, search field type: {1}, search field: {2}'.format(row[field], type(row[field]), field))
                  #row[field] = "{0}/{1}/{2}".format(str(day), str(month), str(year))
                  #print('plan expiry date: {0}'.format(row[field]))
                  temp = findall("^([0-9]{4})-([0-9]{2})-([0-9]{2})$", row["plan expiry date"])[0]
                  row["plan expiry date"] = "%s/%s/%s" % (temp[2], temp[1], temp[0])
               except Exception as error:
                  print('plan expiry date error: {0}'.format(error))

            if field in row:
                flags.append(self.ms_input_content_by_field_name(where,
                        field,
                        base,
                        row[field]))
                #self.wait_for_user_confirmation()
        flags.append(self.ms_input_content_by_field_name(where,
                'search',base))
        
        string_for_printing = (row.__str__().replace('{', '').replace('}', '').replace("'", ''))
        try:
            self.wait_until_hidden_of_element_by_xpath(LOADING_INDICATOR)
            elems = self.driver.find_elements_by_xpath("//*[@title='%s']" % button)
            #elems = self.wait.until(
            #    lambda x: x.find_elements_by_xpath("//*[@title='%s']" %
            #    button)
            #)
        except se.TimeoutException:
            logging('ms_search_member', 'time is out and found no matched members.', base)
            self.logger.error('RPA.ms_search_member',
                'Searching for the member %s using the %s way failed: '
                'time is out and found no matched members.' % (string_for_printing,
                    method_name))
            return False
        else:
            if len(elems) != 1:
                self.logger.error('RPA.ms_search_member',
                    'Searching for the member %s using the %s way failed: '
                    'found %d matched members and do not know which one to proceed.' % (string_for_printing,
                        method_name,
                        len(elems)))
                return False
            else:
                elems[0].click()
                self.logger.info('RPA.pm_search_policy',
                    'Searching for the member %s using the %s way is successful.' % (string_for_printing,
                        method_name))
                return True

    def ms_search_member_policy(self, row, base, excludes=(), record_action="View Record"):
        # Optional override column for searches.  If exists, is a string and is
        # non empty, will force search
        # to use contents of "custom search" column.  Format should be [field
        # name;field name;...],
        # column names seperated by semicolons.  Field names that can be used
        # are:
        # 'member full name','nric','otheric','gender','employee
        # id','dob','external ref id (aka client)',
        # 'policy num','internal plan code id','external plan code','policy eff
        # date','plan attach date',
        # 'plan expiry date'
        if "custom search" in row and not row["custom search"] == "" and type(row["custom search"]) == str:
            fields = [i.strip(" ") for i in row["custom search"].split(";")]
            print('search criteria: {0}'.format(fields))
            flag = self.ms_search_member('inpt member policy search',
                row,
                'primary',
                fields,
                record_action, base)
            return flag

        contains = False
        for exclude in excludes:
            if exclude in ["policy num"]:
                contains = True
        if not contains:
            flag = self.ms_search_member('inpt member policy search',
                row,
                'primary',
                ["policy num"],
                record_action, base)
            if flag:
                return flag

    def ms_complete_inputting(self, xpath, where, message_mapper, base):
        """
        Complete member update.
        :param str xpath:               Xpath of the completing button
        :param str where:               Where we want to complete inputting
        :param dict message_mapper:     Mapper of messages to results
        :rtype:                         bool
        """
        try:
            elem = self.find_element_on_current_page_by_xpath(xpath)
            elem.click()
        except se.NoSuchElementException:
            logging('ms_complete_inputting', 'cannot find the update button.', base)
            self.logger.error('RPA.ms_complete_inputting',
                'Completing %s failed: '
                'cannot find the update button.' % where)
            return False
        except se.WebDriverException as error:
            logging('ms_complete_inputting', 'trying to click on the update button. Error %s' % error, base)
            self.logger.error('RPA.ms_complete_inputting',
                'Completing %s failed when '
                'trying to click on the update button: %s.' % (where,
                    error))
            return False
        try:
            message, result = self.wait_until_visibility_of_message(message_mapper)
        except se.TimeoutException:
            self.logger.error('RPA.ms_complete_inputting',
                'Completing %s failed: '
                'time is out after clicking on the update button.' % where)
            return False
        else:
            if result:
                self.logger.info('RPA.ms_complete_inputting',
                    'Completing %s is successful: %s.' % (where,
                        message))
            else:
                self.logger.error('RPA.ms_complete_inputting',
                    'Completing %s failed: %s.' % (where,
                        message))
            return result

    def ms_chg_name(self, row):
        '''
        Operation function for action chg::name

        :param dict row:                          Data fields attached to the action
        :return:                                        bool
        '''
        flags = list()
        if not self.fields_in_row(["member full name"], row, "chg::name"):
            return False

        primary_search_flag = self.ms_search_member_policy(row, ["member full name"])
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False

        try:
            mm_id = self.get_content_on_current_page_by_xpath(INPT_MEMBER_POLICY_VIEW_FIELD_NAME_TO_XPATH_MAPPER["mm id"])
        except se.NoSuchElementException:
            self.logger.error('MemberSetupRPA.ms_chg_name',
                'Cannot find out the member using MM ID (member ID) at xpath %s.' % INPT_MEMBER_POLICY_VIEW_FIELD_NAME_TO_XPATH_MAPPER["mm id"])
            return False
        except Exception as error:
            self.logger.error('MemberSetupRPA.ms_chg_name',
                'Cannot find out the member using MM ID (member ID) at xpath %s: %s' % (INPT_MEMBER_POLICY_VIEW_FIELD_NAME_TO_XPATH_MAPPER["mm id"],
                    error))
            return False
        # self.go_to(MARC_MEMBER_SEARCH_ADDRESS)
        # self.ms_input_content_by_field_name(
        #     'member search',
        #     'member id',
        #     mm_id
        # )
        # self.driver.find_element_by_xpath("//span[contains(text(),
        # 'Search')]/..").click()
        mm_search_flag = self.ms_search_member('member search',
            pd.Series({'member id': mm_id}),
            'Mm Id',
            ['member id'],
            'Edit Record')
        if mm_search_flag:
            flags.append(mm_search_flag)
        else:
            return False
        try:
            full_name = self.ms_get_content_by_field_name('member update',
                'full name')
        except se.NoSuchElementException:
            self.logger.info('MemberSetupRPA.ms_chg_name',
                'Did not click view record')
            return False

        if full_name != row['member full name']:
            flags.append(self.ms_input_content_by_field_name('member update',
                    'full name',
                    row['member full name']))

            
            flags.append(self.ms_complete_inputting(MEMBER_UPDATE_FIELD_NAME_TO_XPATH_MAPPER['update'],
                    'member update',
                    {
                        'Validation Error: Value is required': False,
                        'Successfully saved record into database': True
                    }))
            return all(flags)
        else:
            self.logger.info('MemberSetupRPA.ms_chg_name',
                'Member full name %s in the data is the same as what is stored in MARC.' % full_name)
            return all(flags)

    def ms_chg_nric_ic_passport(self, row):
        '''
        Operation function for action chg::nric_ic_passport

        :param dict row:                          Data fields attached to the action
        :return:                                        bool
        '''
        flags = list()

        excludes = []
        if "nric" in row:
            excludes.append("nric")

        if 'otheric' in row:
            excludes.append("otheric")

        primary_search_flag = self.ms_search_member_policy(row, excludes)
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False
        try:
            mm_id = self.get_content_on_current_page_by_xpath(INPT_MEMBER_POLICY_VIEW_FIELD_NAME_TO_XPATH_MAPPER["mm id"])
        except se.NoSuchElementException:
            self.logger.error('MemberSetupRPA.ms_chg_nric_ic_passport',
                'Cannot find out the member using MM ID (member ID) at xpath %s.' % INPT_MEMBER_POLICY_VIEW_FIELD_NAME_TO_XPATH_MAPPER["mm id"])
            return False
        except Exception as error:
            self.logger.error('MemberSetupRPA.ms_chg_nric_ic_passport',
                'Cannot find out the member using MM ID (member ID) at xpath %s: %s' % (INPT_MEMBER_POLICY_VIEW_FIELD_NAME_TO_XPATH_MAPPER["mm id"],
                    error))
            return False
        # self.go_to(MARC_MEMBER_SEARCH_ADDRESS)
        # self.ms_input_content_by_field_name(
        #     'member search',
        #     'member id',
        #     mm_id
        # )
        # self.driver.find_element_by_xpath("//span[contains(text(),
        # 'Search')]/..").click()
        mm_search_flag = self.ms_search_member('member search',
            pd.Series({'member id': mm_id}),
            'Mm Id',
            ['member id'],
            'Edit Record')
        if mm_search_flag:
            flags.append(mm_search_flag)
        else:
            return False

        if 'nric' in row:
            national_id = self.ms_get_content_by_field_name('member update',
                'national id')
            if national_id != row['nric']:
                flags.append(self.ms_input_content_by_field_name('member update',
                        'national id',
                        row['nric']))
            else:
                self.logger.warning('MemberSetupRPA.ms_chg_name',
                    'Member national id %s in the data is the same as what is stored in MARC.' % national_id)
        if 'otheric' in row:
            other_ic = self.ms_get_content_by_field_name('member update',
                'other national id')
            if other_ic != row['otheric']:
                flags.append(self.ms_input_content_by_field_name('member update',
                        'other national id',
                        row['otheric']))
            else:
                self.logger.warning('MemberSetupRPA.ms_chg_name',
                    'Member other id %s in the data is the same as what is stored in MARC.' % other_ic)
        if not 'otheric' in row and not 'nric' in row:
            self.logger.error('MemberSetupRPA.ms_chg_name',
                'No IC in IC change operation')


        flags.append(self.ms_complete_inputting(MEMBER_UPDATE_FIELD_NAME_TO_XPATH_MAPPER['update'],
                'member update',
                {
                    'Validation Error: Value is required': False,
                    'Successfully saved record into database': True
                }))
        return all(flags)

    def ms_ter_cancel_delete(self, row):
        '''
        Operation function for action ter::cancel_delete

        :param dict row:                          Data fields attached to the action
        :rtype:                                        bool
        '''
        flags = list()
        if not self.fields_in_row(["cancellation date", "date received by aan"], row, "ter::cancel_delete"):
            return False

        primary_search_flag = self.ms_search_member_policy(row)
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False
        try:
            self.driver.find_element_by_xpath("//span[text()='Cancel']/..").click()
        except se.NoSuchElementException:
            self.logger.error('MemberSetupRPA.ms_ter_cancel_delete',
                'Cannot click on the cancel button.')
            return False
        except Exception as error:
            self.logger.error('MemberSetupRPA.ms_ter_cancel_delete',
                'Cannot click on the cancel button: %s' % error)
            return False
        else:
            self.wait_until_revealed_of_element_by_xpath("//input[@title='Cancellation Date' and @id='commCancelDate_input']/../../..")
            flags.append(self.input_date_on_current_page_by_xpath("//input[@title='Cancellation Date' and @id='commCancelDate_input']",
                    row['cancellation date']))
            flags.append(self.input_date_on_current_page_by_xpath("//input[@title='Cancellation Date' and @id='commDeclareDate_input']",
                    row['date received by aan']))
            flags.append(self.click_on_button_on_current_page_by_xpath("//span[text()='Save']/.."))
        self.wait_until_hidden_of_element_by_xpath("//div[@class='three-quarters']/../..")
        return all(flags)

    def ms_renewal_edit_policy(self, row):
        '''
        Operation function for action renewal::edit_policy

        :param dict row:                          Data fields attached to the action
        :rtype:                                        bool
        '''
        flags = list()
        if not self.fields_in_row(["policy end date", "policy num"], row, "renewal::edit_policy"):
            return False

        primary_search_flag = self.ms_search_member_policy(row, record_action="Edit Record")
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False
        flags.append(self.click_on_onclick_on_current_page_by_xpath("//a[contains(@onclick, 'lookup_inptpolicy')]"))
        flags.append(self.wait_until_hidden_of_element_by_xpath("//div[@class='three-quarters']/../.."))
        flags.append(self.wait_until_revealed_of_element_by_xpath("//td[text()='Policy Id:']/../../../../../.."))

        flags.append(self.ms_input_content_by_field_name('inpt member policy update',
                'policy lookup expiry date',
                row['policy end date']))

        flags.append(self.ms_input_content_by_field_name('inpt member policy update',
                'policy lookup num',
                row['policy num']))
        flags.append(self.ms_input_content_by_field_name('inpt member policy update',
                'policy lookup search'))
        string_for_printing = (row.to_dict().__str__().replace('{', '').replace('}', '').replace("'", ''))
        try:
            total_rows = self.wait.until(lambda x: x.find_elements_by_xpath("//a[contains(@id, 'lookupInptPolicySearchResult')]"))
        except se.TimeoutException:
            try:
                self.wait.until(lambda x: x.find_elements_by_xpath("//td[text()='Policy Id:']/../../../../../..//td[contains(text(), 'No records found.')]"))
            except se.TimeoutException:
                self.logger.error('RPA.ms_renewal_edit_policy',
                    'Renewing policy for %s failed: '
                    'time is out when searching for plans.' % string_for_printing)
            else:
                self.logger.error('RPA.ms_renewal_edit_policy',
                    'Renewing policy for %s failed: '
                    'no matched policy found.' % string_for_printing)
            return False
        else:
            if len(total_rows) > 1:
                self.logger.error('RPA.ms_renewal_edit_policy',
                    'Renewing policy for %s failed: '
                    'found %d policies and cannot proceed.' % (string_for_printing,
                        len(total_rows)))
                return False
            else:
                total_rows[0].click()
                flags.append(self.wait_until_hidden_of_element_by_xpath("//td[text()='Policy Id:']/../../../../../.."))

                flags.append(self.ms_complete_inputting(INPT_MEMBER_POLICY_UPDATE_FIELD_NAME_TO_XPATH_MAPPER['update'],
                        'inpt member policy update',
                        {
                            'Validation Error: Value is required': False,
                            'Successfully saved record into database': True
                        }))
        return True

    def ms_chg_dob(self, row):
        '''
        Operation function for action chg::gender_dob_maritalstatus

        :param dict row:                          Data fields attached to the action
        :rtype:                                        bool
        '''
        flags = list()
        if not self.fields_in_row(['dob'], row, "chg::dob"):
            return False

        primary_search_flag = self.ms_search_member_policy(row, ["dob"])
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False
        try:
            mm_id = self.get_content_on_current_page_by_xpath(INPT_MEMBER_POLICY_ADD_FIELD_NAME_TO_XPATH_MAPPER["dob"])
        except se.NoSuchElementException:
            self.logger.error('MemberSetupRPA.ms_chg_nric_ic_passport',
                'Cannot find out the member using MM ID (member ID) at xpath %s.' % INPT_MEMBER_POLICY_VIEW_FIELD_NAME_TO_XPATH_MAPPER["mm id"])
            return False
        except Exception as error:
            self.logger.error('MemberSetupRPA.ms_chg_nric_ic_passport',
                'Cannot find out the member using MM ID (member ID) at xpath %s: %s' % (INPT_MEMBER_POLICY_VIEW_FIELD_NAME_TO_XPATH_MAPPER["mm id"],
                    error))
            return False
        # self.go_to(MARC_MEMBER_SEARCH_ADDRESS)
        # self.ms_input_content_by_field_name(
        #     'member search',
        #     'member id',
        #     mm_id
        # )
        # self.driver.find_element_by_xpath("//span[contains(text(),
        # 'Search')]/..").click()
        mm_search_flag = self.ms_search_member('member search',
            pd.Series({'member id': mm_id}),
            'Mm Id',
            ['member id'],
            'Edit Record')
        if mm_search_flag:
            flags.append(mm_search_flag)
        else:
            return False
        fields = ['dob']
        for field in fields:
            content = self.ms_get_content_by_field_name('member update',
                field)
            if row[field] != "" and content != row[field]:
                flags.append(self.ms_input_content_by_field_name('member update',
                        field,
                        row[field]))
            else:
                self.logger.warning('MemberSetupRPA.ms_chg_dob',
                    '%s %s in the data is the same as what is stored in MARC.' % (field,
                        content))
        
        flags.append(self.ms_complete_inputting(MEMBER_UPDATE_FIELD_NAME_TO_XPATH_MAPPER['update'],
                'member update',
                {
                    'Validation Error: Value is required': False,
                    'Successfully saved record into database': True
                }))
        return all(flags)

    def ms_chg_handphone_phone_mobile(self, row):
        '''
        Operation function for action chg::handphone_phone_mobile

        :param dict row:                          Data fields attached to the action
        :rtype:                                        bool
        '''
        flags = list()
        if not self.fields_in_row(["phone"], row, "chg::handphone_phone_mobile"):
            return False

        primary_search_flag = self.ms_search_member_policy(row, record_action="Edit Record")
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False
        phone = self.ms_get_content_by_field_name('inpt member policy update',
            'phone')
        if phone != row['phone']:
            flags.append(self.ms_input_content_by_field_name('inpt member policy update',
                    'phone',
                    row['phone']))
           
            flags.append(self.ms_complete_inputting(INPT_MEMBER_POLICY_UPDATE_FIELD_NAME_TO_XPATH_MAPPER['update'],
                    'inpt member policy update',
                    {
                        'Validation Error: Value is required': False,
                        'Record updated successfully': True
                    }))
        else:
            self.logger.warning('MemberSetupRPA.ms_chg_phone',
                'Phone %s in the data is the same as what is stored in MARC.' % phone)
        return all(flags)

    def ms_chg_plan_rpa(self, row):
        '''
        Operation function for action chg::plan

        :param dict row:                          Data fields attached to the action
        :rtype:                                        bool
        '''
        flags = list()
        if not self.fields_in_row(["marc"], row, "chg::plan"):
            return False

        primary_search_flag = self.ms_search_member_policy(row)
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False

        try:
            int_plan_code = self.get_content_on_current_page_by_xpath(INPT_MEMBER_POLICY_VIEW_FIELD_NAME_TO_XPATH_MAPPER["Int Plan Code (Id)"])
        except se.NoSuchElementException:
            self.logger.error('MemberSetupRPA.ms_chg_plan_rpa',
                'Cannot find out the member using MM ID (member ID) at xpath %s.' % INPT_MEMBER_POLICY_VIEW_FIELD_NAME_TO_XPATH_MAPPER["mm id"])
            return False
        except Exception as error:
            self.logger.error('MemberSetupRPA.ms_chg_plan_rpa',
                'Cannot find out the member using MM ID (member ID) at xpath %s: %s' % (INPT_MEMBER_POLICY_VIEW_FIELD_NAME_TO_XPATH_MAPPER["mm id"],
                    error))
            return False

        mm_search_plan_flag = self.ms_search_plan_rpa('plan search',
            pd.Series({'int plan code(id)': int_plan_code}),
            'int plan code(id)',
            ['int plan code(id)'],
            'Update Record')
        
        if mm_search_plan_flag:
            flags.append(mm_search_plan_flag)
        else:
            return False

        flags.append(self.ms_input_content_by_field_name('plan search',
                      'ext plan code',
                      row['external plan code']))

        flags.append(self.ms_complete_inputting(MEMBER_UPDATE_FIELD_NAME_TO_XPATH_MAPPER['update'],
                 'plan search',
                 {
                     'Validation Error: Value is required': False,
                     'Successfully saved record into database': True
                 }))

        try:
            self.wait_until_hidden_of_element_by_xpath(LOADING_INDICATOR)
        except se.TimeoutException:
            try:
                self.wait_until_hidden_of_element_by_xpath(LOADING_INDICATOR)
            except se.TimeoutException:
                self.logger.error('RPA.ms_chg_plan_rpa',
                    'time out for searching for plans')


        return all(flags)

    def ms_chg_plan(self, row):
        '''
        Operation function for action chg::plan

        :param dict row:                          Data fields attached to the action
        :rtype:                                        bool
        '''
        flags = list()
        if not self.fields_in_row(["external plan code", "date received by aan"], row, "chg::plan"):
            return False

        primary_search_flag = self.ms_search_member_policy(row)
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False

        self.ms_input_content_by_field_name('inpt member policy view',
            'upgrade')

        self.wait_until_revealed_of_element_by_xpath('.//span[text()="Upgrade Inpt Member Plan"]/../..')

        if 'internal plan code id' in row:
            flags.append(self.ms_input_content_by_field_name('inpt member policy view',
                    'upgrade plan int plan code',
                    row['internal plan code id']))
        flags.append(self.ms_input_content_by_field_name('inpt member policy view',
                'upgrade plan ext plan code',
                row['external plan code']))

        try:
          temp = findall("^([0-9]{4})-([0-9]{2})-([0-9]{2})$", row['date received by aan'])[0]
          row['date received by aan'] = "%s/%s/%s" % (temp[2], temp[1], temp[0])
        except Exception as error:
          error

        flags.append(self.ms_input_content_by_field_name('inpt member policy view',
                'upgrade plan upgrade date',
                row['date received by aan']))

        sleep(1)

        flags.append(self.ms_input_content_by_field_name('inpt member policy view',
                'upgrade plan search'))

        try:
            self.wait_until_hidden_of_element_by_xpath(LOADING_INDICATOR)
        except se.TimeoutException:
            try:
                self.wait_until_hidden_of_element_by_xpath(LOADING_INDICATOR)
            except se.TimeoutException:
                self.logger.error('RPA.ms_chg_plan',
                    'time out for searching for plans')

        total_rows = []
        try:
            total_rows = self.wait.until(lambda x: x.find_elements_by_xpath("//a[contains(@id, 'commUpgradeSearchResult')]"))
        except se.TimeoutException:
            try:
                self.wait.until(lambda x: x.find_elements_by_xpath("//*[contains(text(), 'No records found.')]"))
            except se.TimeoutException:
                self.logger.error('RPA.ms_chg_plan',
                    'time out for searching for plans')
                return False
            else:
                self.logger.error('RPA.ms_chg_plan',
                    'no matching plan found')

                self.driver.find_element_by_xpath("//span[text()=\"Upgrade Inpt Member Plan\"]/../a[@class='ui-dialog-titlebar-icon ui-dialog-titlebar-close ui-corner-all' and @role=\"button\"]").click()
                return False

        if len(total_rows) > 1:
            self.logger.error('RPA.ms_chg_plan',
                'plan lookup: more than one result')
            return False
        elif total_rows == 0:
            self.logger.error('RPA.ms_chg_plan',
                'plan lookup: no results')
            return False

        else:
            
            status = self.ms_complete_inputting('//a[contains(@id, \'commUpgradeSearchResult\')]',
                'inpt member policy view',
                {
                    'Can\'t upgrade to a same plan.': 1,
                    'InptPolicy/Plan not found for this group policy.': 2,
                    'Upgrade successful.': 3
                })

            if status == 1:
                self.logger.warning('RPA.ms_chg_plan',
                    'plan already in place for that user')
            elif status == 2:
                self.logger.error('RPA.ms_chg_plan',
                    'plan lookup: more than one result')
                return False
            elif status == 3:
                self.logger.info('RPA.ms_chg_plan',
                    'upgrade plan successful')
            else:
                self.logger.error('RPA.ms_chg_plan',
                    'unknown result from adding plan')
                return False

        return all(flags)

    def ms_chg_principalname(self, row):
        '''
        Operation function for action chg::principalname

        :param dict row:                          Data fields attached to the action
        :rtype:                                        bool
        '''
        flags = list()
        if not self.fields_in_row(["principal name", "policy end date"], row, "chg::principalname"):
            return False

        primary_search_flag = self.ms_search_member_policy(row, ["principal name"], record_action="Edit Record")
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False

        flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                'principal add'))

        self.wait_until_revealed_of_element_by_xpath('.//span[text()="Inpt Principal Look Up"]/../..')

        flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                'principal lookup full name',
                row['principal name']))
        
        if 'principal nric' in row:
            flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                    'principal lookup nric',
                    row['principal nric']))
        
        if 'principal other ic' in row:
            flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                    'principal lookup other ic',
                    row['principal other ic']))

        flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                'principal lookup pol end date',
                row['policy end date']))

        flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                'principal lookup search',))

        self.wait_until_hidden_of_element_by_xpath(LOADING_INDICATOR)
        total_rows = []
        try:
            total_rows = self.wait.until(lambda x: x.find_elements_by_xpath("//a[contains(@id, 'lookupPrincipalSearchResult')]"))
        except se.TimeoutException:
            try:
                self.wait.until(lambda x: x.find_elements_by_xpath("//*[contains(text(), 'No records found.')]"))
            except se.TimeoutException:
                self.logger.error('RPA.ms_chg_principalname',
                    'time out for searching for principals')
                self.driver.find_element_by_xpath("//span[text()=\"Inpt Principal Look Up\"]/../a[@class='ui-dialog-titlebar-icon ui-dialog-titlebar-close ui-corner-all' and @role=\"button\"]").click()
            else:
                self.logger.error('RPA.ms_chg_principalname',
                    'no matching principal found')
                self.driver.find_element_by_xpath("//span[text()=\"Inpt Principal Look Up\"]/../a[@class='ui-dialog-titlebar-icon ui-dialog-titlebar-close ui-corner-all' and @role=\"button\"]").click()

        if len(total_rows) > 1:
            self.logger.error('RPA.ms_chg_principalname',
                'principal lookup: more than one result')
            return False
        elif len(total_rows) == 0:
            self.logger.error('RPA.ms_chg_principalname',
                'principal lookup: no results')
            return False
        else:
            total_rows[0].click()
            self.wait_until_hidden_of_element_by_xpath('.//span[text()="Inpt Principal Look Up"]/../..')

       
        flags.append(self.ms_complete_inputting(INPT_MEMBER_POLICY_UPDATE_FIELD_NAME_TO_XPATH_MAPPER['update'],
                'inpt member policy update',
                {
                    'Validation Error: Value is required': False,
                    'Record updated successfully': True
                }))

        return all(flags)

    def ms_chg_update_all(self, row):
        '''
        Operation function for action chg::update_all

        :param dict row:                          Data fields attached to the action
        :rtype:                                        bool
        '''
        flags = list()
        if not self.fields_in_row(["external plan code", "principal name"], row, "chg::update_all"):
            return False

        primary_search_flag = self.ms_search_member_policy(row)
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False
        
        ext_plan_code = self.ms_get_content_by_field_name('inpt member policy view',
            'ext plan code')
        if ext_plan_code != row['external plan code']:
            flags.append(self.ms_chg_plan(row))

        flags.append(self.ms_input_content_by_field_name('inpt member policy view',
                'edit',))

        equivalent = True
        principal_full_name = self.ms_get_content_by_field_name('inpt member policy add',
            'principal full name')
        if principal_full_name == row['principal name']:
            equivalent = False

        if 'principal nric' in row:
            principal_nric = self.ms_get_content_by_field_name('inpt member policy add',
                'principal nric')
            if principal_nric != row['principal nric']:
                equivalent = False

        if 'principal other ic' in row:
            principal_other_ic = self.ms_get_content_by_field_name('inpt member policy add',
                'principal other ic')
            if principal_other_ic != row['principal other ic']:
                equivalent = False

        if not equivalent:
            flags.append(self.ms_chg_principalname(row))
            return all(flags)

        return all(flags)

    def ms_chg_update_member_rpa(self, row, base):
        '''
          Operation function for action update_member
          full name, nric, otheric and dob
        '''
        flags = list()
        if not self.fields_in_row(["marc"], row, "chg::update_member"):
            return False

        try:
            fields_to_change = (row["marc"].split(";")[row["action"].split(";").index("update_member")]).split("_")[0:]
            print('fields_to_change: {0}'.format(fields_to_change))
        except IndexError:
            self.logger.error('RPA.ms_chg_update_member_rpa',
                'marc parameters not found')
        except:
            self.logger.error('RPA.ms_chg_update_member_rpa',
                'malformed action or marc field')
        cleaned_fields = []
        # Error and conflict handling

        
        if len(fields_to_change) == 0:
            self.logger.error('RPA.ms_chg_update_member_rpa',
                'no MARC parameter given')
            return False
        for field in fields_to_change:
            if field.lower() not in ["name", "dob", "nric",
                             "otheric"]:
                self.logger.error('RPA.ms_chg_update_fields_rpa',
                    'invalid action %s in MARC parameter' % field)
                return False
            elif field.lower() in ["name", "dob", "nric",
                             "otheric"]:
                cleaned_fields.append(field.lower())

        #search criteria
        primary_search_flag = self.ms_search_member_policy(row, base, record_action="View Record")
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False

        try:
            mm_id = self.get_content_on_current_page_by_xpath(INPT_MEMBER_POLICY_VIEW_FIELD_NAME_TO_XPATH_MAPPER["mm id"])
        except se.NoSuchElementException:
            self.logger.error('MemberSetupRPA.ms_chg_update_member_rpa',
                'Cannot find out the member using MM ID (member ID) at xpath %s.' % INPT_MEMBER_POLICY_VIEW_FIELD_NAME_TO_XPATH_MAPPER["mm id"])
            return False
        except Exception as error:
            self.logger.error('MemberSetupRPA.ms_chg_update_member_rpa',
                'Cannot find out the member using MM ID (member ID) at xpath %s: %s' % (INPT_MEMBER_POLICY_VIEW_FIELD_NAME_TO_XPATH_MAPPER["mm id"],
                    error))
            return False

        mm_search_flag = self.ms_search_member_rpa('member search',
            pd.Series({'member id': mm_id}),
            'Mm Id',
            ['member id'],
            'Edit Record',base)
        
        if mm_search_flag:
            flags.append(mm_search_flag)
        else:
            return False

        for field in cleaned_fields:
          #name_dob_nric_otheric;
          #print(field)
          if field == "name":
            try:
               flags.append(self.ms_input_content_by_field_name('member update',
                      'member full name',
                      base,
                      row['member full name']))
               
            except Exception as error:
              self.logger.error('ms_chg_update_member_rpa: member full name', str(error))
          elif field == "dob":
              try:
                  temp = findall("^([0-9]{4})-([0-9]{2})-([0-9]{2})$", row["dob"])[0]
                  update_dob = "%s/%s/%s" % (temp[2], temp[1], temp[0])
              except Exception as error:
                  error
              flags.append(self.ms_input_content_by_field_name('member update',
                      'dob',
                      base,
                      update_dob))
              
          elif field == "nric":
            national_id = self.ms_get_content_by_field_name('member update',
                'nric')
            if national_id != row['nric']:
                if row['nric']=="nan":
                   self.logger.info('MemberSetupRPA.ms_update_member_rpa',
                                    'nric is None.')
                else:
                   flags.append(self.ms_input_content_by_field_name('member update',
                        'nric',
                        base,
                        row['nric']))
                   
            else:
                self.logger.warning('MemberSetupRPA.ms_update_member_rpa',
                    'Member nric %s in the data is the same as what is stored in MARC.' % national_id)
          elif field == "otheric":
            other_ic = self.ms_get_content_by_field_name('member update',
                'otheric')
            if other_ic != row['otheric']:
                if row['otheric']=="nan":
                  self.logger.info('MemberSetupRPA.ms_update_member_rpa',
                                   'otheric is None.')
                else:
                  flags.append(self.ms_input_content_by_field_name('member update',
                        'otheric',
                        base,
                        row['otheric']))
                  
            else:
                self.logger.warning('MemberSetupRPA.ms_update_member_rpa',
                    'Member other id %s in the data is the same as what is stored in MARC.' % other_ic)
            
        flags.append(self.ms_complete_inputting(RPA_INPT_MEMBER_UPDATE_TO_XPATH_MAPPER['update'],
                 'member update',
                 {
                     'Validation Error: Value is required': False,
                     'Successfully saved record into database': True
                 },base))
        try:
            self.wait_until_hidden_of_element_by_xpath(LOADING_INDICATOR)
        except se.TimeoutException:
            try:
                self.wait_until_hidden_of_element_by_xpath(LOADING_INDICATOR)
            except se.TimeoutException:
                self.logger.error('RPA.ms_chg_plan_rpa',
                    'time out for searching for plans')
        return all(flags)


    def ms_chg_update_fields_rpa(self, row, base):
        '''
        Operation function for action chg::update_fields

        :param dict row:                          Data fields attached to the action
        :rtype:                                        bool
        '''
        flags = list()
        if not self.fields_in_row(["marc"], row, "chg::update_fields"):
            return False

        try:
            fields_to_change = (row["marc"].split(";")[row["action"].split(";").index("update_fields")]).split("_")[0:]
            print('fields_to_change: {0}'.format(fields_to_change))
            
        except IndexError:
            self.logger.error('RPA.ms_chg_update_fields_rpa',
                'marc parameters not found')
        except:
            self.logger.error('RPA.ms_chg_update_fields_rpa',
                'malformed action or marc field')
        cleaned_fields = []
        # Error and conflict handling

        
        if len(fields_to_change) == 0:
            self.logger.error('RPA.ms_chg_update_fields_rpa',
                'no MARC parameter given')
            return False
        for field in fields_to_change:
            if field.lower() not in ["joindate", "planattachdate", "relationship",
                             "employee", "extref", "plan", "principal"]:
                self.logger.error('RPA.ms_chg_update_fields_rpa',
                    'invalid action %s in MARC parameter' % field)
                return False
            elif field.lower() in ["joindate", "planattachdate", "relationship",
                             "employee", "extref", "plan", "principal"]:
                cleaned_fields.append(field.lower())

        #search criteria
        primary_search_flag = self.ms_search_member_policy(row, base, record_action="Edit Record")
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False
               
        for field in cleaned_fields:
            if field == "joindate":
                 try:
                    #self.wait_for_user_confirmation()
                    temp = findall("^([0-9]{4})-([0-9]{2})-([0-9]{2})$", row["first join date"])[0]
                    update_first_join_date = "%s/%s/%s" % (temp[2], temp[1], temp[0])
                 except Exception as error:
                    error
                 flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                    'first join date',
                    base,
                    update_first_join_date))
                 #self.wait_for_user_confirmation()
            elif field == "planattachdate":
                try:
                    #self.wait_for_user_confirmation()
                    temp = findall("^([0-9]{4})-([0-9]{2})-([0-9]{2})$", row["plan attach date"])[0]
                    update_plan_attach_date = "%s/%s/%s" % (temp[2], temp[1], temp[0])
                except Exception as error:
                    error
                flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                    'plan attach date',
                    base,
                    update_plan_attach_date))
                #self.wait_for_user_confirmation()
            elif field == "relationship":
                flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                    'relationship',
                    base,
                    row["relationship"]))
            elif field == "employee":
                 if row['employee id']=="nan":
                   self.logger.info('MemberSetupRPA.ms_chg_update_fields_rpa',
                                    'employee id is None.')
                 else:
                    flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                      'employee id',
                      base,
                      row["employee id"]))
                 #self.wait_for_user_confirmation()
            elif field == "extref":
              try:
                #self.wait_for_user_confirmation()
                #print('ext ref id: {0}'.format(row['external ref id  aka client ']))
                if row['external ref id  aka client ']=="nan":
                    self.logger.info('MemberSetupRPA.ms_chg_update_fields_rpa',
                                     'extref is None.')
                else:
                   flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                    'external ref id aka client',
                    base,
                    row["external ref id  aka client "]))
                   #self.wait_for_user_confirmation()
              except Exception as error:
                print('ext ref id error: {0}'.format(error))
                
            elif field == "plan":
                #self.wait_for_user_confirmation()
                flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                'member plan add',base))

                self.wait_until_revealed_of_element_by_xpath('.//span[text()="Inpt Plan Look Up"]/../..')
                
                try:
                  #validate external plan code if empty then use internal plan code
                  if row["external plan code"] == "nan":
                    row['internal plan code id'] = int(float(row['internal plan code id']))
                    flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                               'plan lookup plan int ref',
                               base,
                               row["internal plan code id"]))
                  else:
                    flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                               'plan lookup plan ext ref',
                               base,
                               row["external plan code"]))
                except Exception as error:
                  print(error)

                try:
                  flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                             'plan lookup search',base))
                except Exception as error:
                  print(error)

                self.wait_until_hidden_of_element_by_xpath(LOADING_INDICATOR)

                total_rows = []
                try:
                    total_rows = self.wait.until(lambda x: x.find_elements_by_xpath("//a[contains(@id, 'lookupInptPlanSearchResult')]"))
                except se.TimeoutException:
                    try:
                        self.wait.until(lambda x: x.find_elements_by_xpath("//*[contains(text(), 'No records found.')]"))
                    except se.TimeoutException:
                        self.logger.error('RPA.ms_chg_update_fields_rpa',
                            'time out for searching for plan')
                        self.driver.find_element_by_xpath("//span[text()=\"Inpt Plan Look Up\"]/../a[@class='ui-dialog-titlebar-icon ui-dialog-titlebar-close ui-corner-all' and @role=\"button\"]").click()
                    else:
                        self.logger.info('RPA.ms_chg_update_fields_rpa',
                            'no matching plan found')
                        self.driver.find_element_by_xpath("//span[text()=\"Inpt Plan Look Up\"]/../a[@class='ui-dialog-titlebar-icon ui-dialog-titlebar-close ui-corner-all' and @role=\"button\"]").click()

                if len(total_rows) > 1:
                    self.logger.error('RPA.ms_chg_update_fields_rpa',
                        'plan lookup: more than one result')
                    #return False
                elif len(total_rows) == 0:
                    self.logger.error('RPA.ms_chg_update_fields_rpa',
                        'plan lookup: no results')
                    #return False
                else:
                    
                    total_rows[0].click()
                    
                    self.wait_until_hidden_of_element_by_xpath('.//span[text()="Inpt Plan Look Up"]/../..')
                #self.wait_for_user_confirmation()  
            elif field == "principal":
                #self.wait_for_user_confirmation()
                valid = False
                print('principal')
                #print(row)
                principle_fields = []
                print('principal nric value: {0}, principal type: {1}'.format(row["principal nric"], type(row["principal nric"])))
                if row["relationship"]== "P" or row["relationship"] == "Principal":
                  self.logger.info('MemberSetupRPA.ms_chg_update_fields_rpa',
                                     'Relationship is principal.')
                else:
                  if row["principal nric"]=="nan":
                    print('principal otheric value: {0}, principal type: {1}'.format(row["principal other ic"], type(row["principal other ic"])))
                    if row["principal other ic"]=="nan":
                      self.logger.info('RPA.ms_chg_update_fields_rpa',
                          'principle lookup: no results')
                      valid = False
                    else:
                      principle_fields.append("principal lookup other ic")
                      valid = True
                  else:
                    principle_fields.append("principal lookup nric")
                    valid = True

                  if valid == True:
                    principle_fields.append("principal lookup pol end date")

                    #print('principle_fields')
                    print('principle_fields: {0}'.format(principle_fields))

                    flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                    'principal add',base))

                    self.wait_until_revealed_of_element_by_xpath('.//span[text()="Inpt Principal Look Up"]/../..')

                    for principle_field in principle_fields:
                        if principle_field == "principal lookup other ic":
                          flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                                 principle_field,
                                 base,
                                 row["principal other ic"]))
                        elif principle_field == "principal lookup nric":
                          flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                                 principle_field,
                                 base,
                                 row["principal nric"]))
                        elif principle_field == "principal lookup pol end date":
                          try:
                              temp = findall("^([0-9]{4})-([0-9]{2})-([0-9]{2})$", row["plan expiry date"])[0]
                              row["plan expiry date"] = "%s/%s/%s" % (temp[2], temp[1], temp[0])
                              #self.wait_for_user_confirmation()
                          except Exception as error:
                              error
                          flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                                 principle_field,
                                 base,
                                 row["plan expiry date"]))

                    try:
                      flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                                 'principal lookup search', base))
                  
                    except Exception as error:
                      print(error)

                    self.wait_until_hidden_of_element_by_xpath(LOADING_INDICATOR)
                    total_rows = []
                    try:
                        total_rows = self.wait.until(lambda x: x.find_elements_by_xpath("//a[contains(@id, 'lookupPrincipalSearchResult')]"))
                    except se.TimeoutException:
                        try:
                            self.wait.until(lambda x: x.find_elements_by_xpath("//*[contains(text(), 'No records found.')]"))
                        except se.TimeoutException:
                            self.logger.info('RPA.ms_chg_update_fields_rpa',
                                'time out for searching for principals')
                            self.driver.find_element_by_xpath("//span[text()=\"Inpt Principal Look Up\"]/../a[@class='ui-dialog-titlebar-icon ui-dialog-titlebar-close ui-corner-all' and @role=\"button\"]").click()
                        else:
                            self.logger.info('RPA.ms_chg_update_fields_rpa',
                                'no matching principal found')
                            self.driver.find_element_by_xpath("//span[text()=\"Inpt Principal Look Up\"]/../a[@class='ui-dialog-titlebar-icon ui-dialog-titlebar-close ui-corner-all' and @role=\"button\"]").click()

                    if len(total_rows) > 1:
                        self.logger.info('RPA.ms_chg_update_fields_rpa',
                            'principal lookup: more than one result')
                        #return False
                    elif len(total_rows) == 0:
                        self.logger.info('RPA.ms_chg_update_fields_rpa',
                            'principal lookup: no results')
                        #return False
                    else:
                        total_rows[0].click()
                        self.wait_until_hidden_of_element_by_xpath('.//span[text()="Inpt Principal Look Up"]/../..')
                  else:
                    self.logger.info('RPA.ms_chg_update_fields_rpa',
                            'principle nric & otheric: None')
            else:
                self.logger.error('RPA.ms_chg_update_fields',
                    'error occurred with cleaned field $s' % field)
                #return False
            #self.wait_for_user_confirmation()

        self.wait_until_hidden_of_element_by_xpath(LOADING_INDICATOR)
        flags.append(self.ms_complete_inputting(RPA_INPT_MEMBER_POLICY_UPDATE_TO_XPATH_MAPPER['update'],
                'inpt member policy add',
                {
                    'Validation Error: Value is required': False,
                    'Validation Error: Value is required': False,
                    'Record updated successfully': True
                },base))
        #try:
        #    self.wait_until_hidden_of_element_by_xpath(LOADING_INDICATOR)
        #except se.TimeoutException:
        #    try:
        #        self.wait_until_hidden_of_element_by_xpath(LOADING_INDICATOR)
        #    except se.TimeoutException:
        #        self.logger.error('RPA.ms_chg_plan_rpa',
        #            'time out for searching for plans')
        return all(flags)

    def ms_chg_update_fields(self, row):
        '''
        Operation function for action chg::update_fields

        :param dict row:                          Data fields attached to the action
        :rtype:                                        bool
        '''
        flags = list()
        if not self.fields_in_row(["marc"], row, "chg::update_fields"):
            return False

        try:
            fields_to_change = (row["marc"].split(";")[row["action"].split(";").index("update_fields")]).split("_")[0:]
            #print('ms_chg_update_field')
            #print(fields_to_change)
        except IndexError:
            self.logger.error('RPA.ms_chg_update_fields',
                'marc parameters not found')
        except:
            self.logger.error('RPA.ms_chg_update_fields',
                'malformed action or marc field')
        cleaned_fields = []
        # Error and conflict handling

        
        if len(fields_to_change) == 0:
            self.logger.error('RPA.ms_chg_update_fields',
                'no MARC parameter given')
            return False
        for field in fields_to_change:
            if field.lower() not in ["name", "nric", "ic", "dob", 
                            "joindate", "planattachdate", "relationship",
                             "employee", "expdate", "extref"]:
                self.logger.error('RPA.ms_chg_update_fields',
                    'invalid action %s in MARC parameter' % field)
                return False
            elif field.lower() in ["name", "nric", "ic", "dob", 
                              "joindate", "planattachdate", "relationship",
                             "employee", "expdate", "extref"]:
                cleaned_fields.append(field.lower())
        if "name" in cleaned_fields and len(cleaned_fields) > 1:
            self.logger.error('RPA.ms_chg_update_fields',
                'actions on name can cause conflicts with other fields, do name seperately')
            return False
        for field in cleaned_fields:
            if field == "name":
                flags.append(self.ms_chg_name(row))
            elif field == "nric" or field == "ic":
                flags.append(self.ms_chg_nric_ic_passport(row))
            elif field == "dob":
                flags.append(self.ms_chg_dob(row))
            elif field == "joindate":
                flags.append(self.ms_chg_chg_first_join_date(row))
            elif field == "planattachdate":
                flags.append(self.ms_chg_plan_attach_date(row))
            elif field == "effdate":
                flags.append(self.ms_chg_chg_effective_date(row))
            elif field == "relationship":
                flags.append(self.ms_chg_relationship(row))
            elif field == "employee":
                flags.append(self.ms_chg_employee(row))
            elif field == "expdate":
                flags.append(self.ms_chg_expiry_date(row))
            elif field == "extref":
                flags.append(self.ms_chg_ext_ref(row))
            else:
                self.logger.error('RPA.ms_chg_update_fields',
                    'error occurred with cleaned field $s' % field)
                return False
        return all(flags)

    def ms_chg_vip_vvip(self, row):
        '''
        Operation function for action chg::vip_vvip

        :param dict row:                          Data fields attached to the action
        :rtype:                                        bool
        '''
        flags = list()
        primary_search_flag = self.ms_search_member_policy(row, record_action="Edit Record")
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False

        flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                'vip',
                True))

        
        flags.append(self.ms_complete_inputting(INPT_MEMBER_POLICY_ADD_FIELD_NAME_TO_XPATH_MAPPER['update'],
                'inpt member policy add',
                {
                    'Validation Error: Value is required': False,
                    'Successfully saved record into database': True
                }))

        return all(flags)

    def ms_chg_special_condition(self, row):
        '''
        Operation function for action chg::special_condition

        :param dict row:                          Data fields attached to the action
        :rtype:                                        bool
        '''
        flags = list()
        param = []
        try:
            param = (row["marc"].split(";")[row["action"].split(";").index("special_condition")]).split("_")[1:]
        except IndexError:
            self.logger.error('RPA.ms_chg_special_condition',
                'marc parameters not found')
        except:
            self.logger.error('RPA.ms_chg_special_condition',
                'malformed action or marc field')
        cleaned_fields = []
        # Error and conflict handling

        if len(param) == 0:
            self.logger.error('RPA.ms_chg_special_condition',
                'no MARC parameter given')
            return False

        if len(param) != 1:
            self.logger.error('RPA.ms_chg_special_condition',
                'more than one MARC parameter given for special condition')
            return False

        param = param[0]
        if param.lower() not in ["amend", "append", "overwrite"]:
            self.logger.error('RPA.ms_chg_special_condition',
                'Invalid MARC parameter for special condition change, must be amend/overwrite or append')
            return False

        primary_search_flag = self.ms_search_member_policy(row, record_action="Edit Record")
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False

        if param.lower() in ["amend", "overwrite"]:
            flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                    'special condition',
                    row["special condition"]))
        elif param.lower() in ["append"]:
            temp = self.ms_get_content_by_field_name('inpt member policy add',
                'special condition')
            temp = "update %s: %s\n%s" % (datetime.now().strftime("%d/%m/%Y"),
                row["special condition"],
                temp)
            flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                    'special condition',
                    temp))

        
        flags.append(self.ms_complete_inputting(INPT_MEMBER_POLICY_ADD_FIELD_NAME_TO_XPATH_MAPPER['update'],
                'inpt member policy add',
                {
                    'Validation Error: Value is required': False,
                    'Successfully saved record into database': True
                }))

        return all(flags)

    def ms_chg_chg_first_join_date(self, row):
        '''
        Operation function for action chg::chg_first_join_date

        :param dict row:                          Data fields attached to the action
        :rtype:                                        bool
        '''
        
        flags = list()
        primary_search_flag = self.ms_search_member_policy(row, record_action="Edit Record")
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False

        try:
          temp = findall("^([0-9]{4})-([0-9]{2})-([0-9]{2})$", row["first join date"])[0]
          row["first join date"] = "%s/%s/%s" % (temp[2], temp[1], temp[0])
        except Exception as error:
          error

        flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                'first join date',
                row["first join date"]))

        attach = self.ms_get_content_by_field_name('inpt member policy add',
            'first join date')

        attach = datetime.strptime(attach, "%d/%m/%Y")
        first_join = datetime.strptime(row["first join date"], "%d/%m/%Y")
        if first_join > attach:
            flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                    'first join date',
                    row["first join date"]))

        
        #flags.append(
        #    self.ms_complete_inputting(
        #        RPA_INPT_MEMBER_POLICY_UPDATE_TO_XPATH_MAPPER['update'],
        #        'inpt member policy add',
        #        {
        #            'Validation Error: Value is required': False,
        #            'Successfully saved record into database': True
        #        }
        #    )
        #)

        return all(flags)

    def ms_chg_plan_attach_date(self, row):
        '''
        Operation function for action chg::chg_first_join_date

        :param dict row:                          Data fields attached to the action
        :rtype:                                        bool
        '''
        
        flags = list()
        primary_search_flag = self.ms_search_member_policy(row, record_action="Edit Record")
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False

        try:
          temp = findall("^([0-9]{4})-([0-9]{2})-([0-9]{2})$", row["plan attach date"])[0]
          row["plan attach date"] = "%s/%s/%s" % (temp[2], temp[1], temp[0])
        except Exception as error:
          error

        flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                'plan attach date',
                row["plan attach date"]))

        attach = self.ms_get_content_by_field_name('inpt member policy add',
            'plan attach date')

        attach = datetime.strptime(attach, "%d/%m/%Y")
        first_join = datetime.strptime(row["plan attach date"], "%d/%m/%Y")
        if first_join > attach:
            flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                    'plan attach date',
                    row["plan attach date"]))

        
        #flags.append(
        #    self.ms_complete_inputting(
        #        RPA_INPT_MEMBER_POLICY_UPDATE_TO_XPATH_MAPPER['update'],
        #        'inpt member policy add',
        #        {
        #            'Validation Error: Value is required': False,
        #            'Successfully saved record into database': True
        #        }
        #    )
        #)

        return all(flags)
    
    def ms_chg_chg_effective_date(self, row):
        '''
        Operation function for action chg::chg_first_effective_date

        :param dict row:                          Data fields attached to the action
        :rtype:                                        bool
        '''
        flags = list()
        primary_search_flag = self.ms_search_member_policy(row, record_action="Edit Record")
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False

        flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                'policy eff date',
                row["policy eff date"]))

       
        flags.append(self.ms_complete_inputting(INPT_MEMBER_POLICY_ADD_FIELD_NAME_TO_XPATH_MAPPER['update'],
                'inpt member policy add',
                {
                    'Validation Error: Value is required': False,
                    'Successfully saved record into database': True
                }))

        return all(flags)

    def ms_new_vip_vvip(self, row):
        '''
        Operation function for action new::vip_vvip

        :param dict row:                          Data fields attached to the action
        :rtype:                                        bool
        '''
        flags = list()
        primary_search_flag = self.ms_search_member_policy(row, record_action="Edit Record")
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False

        flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                'vip',
                True))

       
        flags.append(self.ms_complete_inputting(INPT_MEMBER_POLICY_ADD_FIELD_NAME_TO_XPATH_MAPPER['update'],
                'inpt member policy add',
                {
                    'Validation Error: Value is required': False,
                    'Successfully saved record into database': True
                }))

        return all(flags)

    def ms_import_row(self, row):
        '''
        Operation function for import action

        :param dict row:                          Row to be imported to marc
        :rtype:                                   bool
        '''
        flags = list()
        try:
            temp = pd.Dataframe([row])
            tempPath = os.path.join(os.getcwd(), "temp.xlsx")
            create_excel(tempPath, temp)
            flags.append(self.up_upload_file(tempPath))
            os.remove("temp.xlsx")
            return all(flags)
        except:
            return False

    def ms_chg_relationship(self, row):
        '''
        Change relationship of member.

        :param dict row:                          Data fields attached to the action
        :rtype:                                   bool
        '''
        flags = list()
        primary_search_flag = self.ms_search_member_policy(row, record_action="Edit Record")
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False

        flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                'relationship',
                row["relationship"]))

       
        #flags.append(
        #    self.ms_complete_inputting(
        #        INPT_MEMBER_POLICY_ADD_FIELD_NAME_TO_XPATH_MAPPER['update'],
        #        'inpt member policy add',
        #        {
        #            'Validation Error: Value is required': False,
        #            'Successfully saved record into database': True
        #        }
        #    )
        #)

        return all(flags)

    def ms_chg_employee(self, row):
        '''
        Change employee ID of member.

        :param dict row:                          Data fields attached to the action
        :rtype:                                   bool
        '''
        flags = list()
        primary_search_flag = self.ms_search_member_policy(row, record_action="Edit Record")
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False

        flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                'employee id',
                row["employee id"]))

       
        #flags.append(
        #    self.ms_complete_inputting(
        #        INPT_MEMBER_POLICY_ADD_FIELD_NAME_TO_XPATH_MAPPER['update'],
        #        'inpt member policy add',
        #        {
        #            'Validation Error: Value is required': False,
        #            'Successfully saved record into database': True
        #        }
        #    )
        #)

        return all(flags)

    def ms_chg_expiry_date(self, row):
        '''
        Change expiry date of member's plan.

        :param dict row:                          Data fields attached to the action
        :rtype:                                   bool
        '''
        flags = list()
        primary_search_flag = self.ms_search_member_policy(row, record_action="Edit Record")
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False

        flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                'policy end date',
                row["policy end date"]))

       
        flags.append(self.ms_complete_inputting(INPT_MEMBER_POLICY_ADD_FIELD_NAME_TO_XPATH_MAPPER['update'],
                'inpt member policy add',
                {
                    'Validation Error: Value is required': False,
                    'Successfully saved record into database': True
                }))

        return all(flags)

    def ms_chg_ext_ref(self, row):
        '''
        Change external reference ID of member.

        :param dict row:                          Data fields attached to the action
        :rtype:                                   bool
        '''
        
        flags = list()
        primary_search_flag = self.ms_search_member_policy(row, record_action="Edit Record")
        if primary_search_flag:
            flags.append(primary_search_flag)
        else:
            return False

        flags.append(self.ms_input_content_by_field_name('inpt member policy add',
                'external ref id aka client',
                row["external ref id aka client"]))

       
        #flags.append(
        #    self.ms_complete_inputting(
        #        INPT_MEMBER_POLICY_ADD_FIELD_NAME_TO_XPATH_MAPPER['update'],
        #        'inpt member policy add',
        #        {
        #            'Validation Error: Value is required': False,
        #            'Successfully saved record into database': True
        #        }
        #    )
        #)

        return all(flags)

    def ms_act_on_member_rpa(self, action, row, base):
        
        if action == 'update_fields':
             return self.ms_chg_update_fields_rpa(row, base)
             #return True
        elif action == 'update_member':
             return self.ms_chg_update_member_rpa(row, base)
             #return True
        #elif action == 'plan':
        #     return self.ms_chg_plan_rpa(row)
        #elif action == 'principalname':
        #     return self.ms_chg_principalname(row)
        else:
            print('validate action:' + action)
            valid_actions = ['update_fields', 'update_member']
            self.logger.error('MemberSetupRPA.ms_act_on_member_rpa',
                'The action %s has not been implemented. '
                'Implemented actions: %s.' % (action,
                    ', '.join(valid_actions)))
            raise NotImplementedError('The action %s has not been implemented. '
                'Implemented actions: %s.' % (action,
                    ', '.join(valid_actions)))

    def ms_act_on_member(self, action, row):
        
        if action == 'dob':
             return self.ms_chg_dob(row)
            #return True
        elif action == 'handphone_phone_mobile':
             return self.ms_chg_handphone_phone_mobile(row)
            #return True
        elif action == 'name':
             return self.ms_chg_name(row)
            #return True
        elif action == 'nric_ic_passport':
             return self.ms_chg_nric_ic_passport(row)
            #return True
        elif action == 'plan':
             return self.ms_chg_plan(row)
            #return True
        elif action == 'principalname':
             return self.ms_chg_principalname(row)
            #return True
        elif action == 'update_all':
             return self.ms_chg_update_all(row)
            #return True
        elif action == 'update_fields':
             return self.ms_chg_update_fields(row)
            #return True
        elif action == 'vip_vvip':
             return self.ms_chg_vip_vvip(row)
            #return True
        elif action == 'special_condition':
             return self.ms_chg_special_condition(row)
            #return True
        elif action == 'chg_first_join_date':
             return self.ms_chg_chg_first_join_date(row)
            #return True
        elif action == 'vip_vvip':
             return self.ms_new_vip_vvip(row)
            #return True
        elif action == 'edit_policy':
             return self.ms_renewal_edit_policy(row)
             #return True
        elif action == 'cancel_delete':
             return self.ms_ter_cancel_delete(row)
             #return True
        elif action == 'effdate':
             return self.ms_chg_chg_effective_date(row)
             #return True
        elif action == 'import':
             return self.ms_import_row(row)
             #return True
        elif action == 'relationship':
             return self.ms_chg_relationship(row)
             #return True
        elif action == 'employee':
             return self.ms_chg_employee(row)
             #return True
        elif action == 'expdate':
             return self.ms_chg_expiry_date(row)
             #return True
        elif action == 'extref':
             return self.ms_chg_ext_ref(row)
        elif action == 'planattachdate':
             return self.ms_chg_plan_attach_date(row)
             #return True
        else:
            
            valid_actions = ['dob',
                'handphone_phone_mobile',
                'name',
                'nric_ic_passport',
                'plan',
                'principalname',
                'update_all',
                'update_fields',
                'vip_vvip',
                'vip_vvip',
                'chg_24hr_plan',
                'edit_policy',
                'cancel_delete',
                'alert',
                'effdate',
                'import',
                'planattachdate']
            self.logger.error('MemberSetupRPA.ms_act_on_member',
                'The action %s has not been implemented. '
                'Implemented actions: %s.' % (action,
                    ', '.join(valid_actions)))
            raise NotImplementedError('The action %s has not been implemented. '
                'Implemented actions: %s.' % (action,
                    ', '.join(valid_actions)))

    def ms_update_member(self, dataframe, base):
        """
        TODO: Alter to use promboom's input format

        Use a MemberSetupXLSXFileLoader, iterate rows in it and do necessary updates.
        :param MemberSetupXLSXFileLoader loader:        Loader of the input file
        :rtype:                                         bool
        """
        audit_log("Selenium Progress","Completed...", base)
        flags = list()
        
        dataframe.columns = map(str.lower, dataframe.columns)
        #print('rpa column: {0}'.format(dataframe.columns))
        #print(dataframe.dtypes)
        table = dataframe.to_dict(orient="records")
        idx = 0
        good = []
        bad = []
        alert = ""

        for row in table:
            row["status"] = ""
            row["action"] = ""
            row["marc"] = ""
            row["custom search"] = ""
            row["error"] = ""
            row["first join date dm"] = row["first join date"]
            row["plan attach date dm"] = row["plan attach date"]
            row["plan expiry date dm"] = row["plan expiry date"]
            #Postgres action lookup
          
            keys = list(row)
         
            for key in keys:
            #    try:
            #      if len(findall("^[0-9]{4}-[0-9]{2}-[0-9]{2}$", row[key])) >
            #      0:
            #          temp = findall("^([0-9]{4})-([0-9]{2})-([0-9]{2})$",
            #          row[key])[0]
            #          row[key] = "%s/%s/%s" % (temp[2], temp[1], temp[0])
            #    except Exception as error:
            #      error
            #      #logging('ms_update_member',error)
                row[sub("[^a-zA-Z0-9]", " ", key)] = row.pop(key)
                
            try:
                actions, params, search = convert_action_code_to_actions(row["action code"], row["import type"], base)
                
            except:
                
                row["status"] = "Could not convert action code. "
                if not "import type" in row:
                  row["status"] = row["status"] + "Missing import type. "
                if not "action code" in row:
                  row["status"] = row["status"] + "Missing action code. "
               
                continue
                
            row["marc"] = params
            row["action"] = actions
            row["custom search"] = search

            if row["nric"] != "nan":
              row["nric"] = '{:0>12}'.format(str(row["nric"]))
              
            if row["employee id"] != "nan":
              row["employee id"] = str(row["employee id"])
            else:
              row["employee id"]="nan"
          
            #print('nric: {0}'.format(row["nric"]))
            actions = actions.split(";")

            row["record_status"] = ""
            for action in actions:
              try:
                if action == "alert":
                    flags.append(True)
                    alert = alert + "Placeholder alert. "
                    row["status"] = "Completed action %s. " % action
                    row["record_status"] = "Unsuccessful"
                    continue
                    self.logger.clear_latest_error_messages()
                elif action == "import":
                    flags.append(True)
                    alert = alert + "import type"
                    row["status"] = "Completed action %s. " % action
                    row["record_status"] = "Successful"
                    continue
                    self.logger.clear_latest_error_messages()
                else:
                    try:
                      print('ms_update_member_action: ' + action)
                      flag = self.ms_act_on_member_rpa(action, row, base)
                      flags.append(flag)
                    except NotImplementedError:
                      alert = alert + 'MemberSetupRPA.ms_update_member: Action %s is not implemented. ' % action
                      flag = False
                      flags.append(flag)
                      row["status"] = "Action %s is not required update MARC. " % action
                    except:
                      alert = alert + 'MemberSetupRPA.ms_update_member: Error while executing action %s. ' % action
                      flag = False
                      flags.append(flag)
                      row["status"] = "Error encountered while executing action %s. " % action
                    row["status"] = "Completed action %s. " % action

                #defind record status for row status
                
                if flag:
                   good.append(row)
                   self.logger.info('MemberSetupRPA.ms_update_member',
                            'No. %d row is %s and is processed successfully.' % (idx + 1, action))
                   update_audit("Selenium Progress", "Completed...No. {0} row is {1} and is processed successfully.".format(idx+1, action), base)
                   row["record_status"] = "Successful"
                else:
                   row["error"] = self.logger.get_latest_error_messages()
                   bad.append(row)
                   row["record_status"] = "Unsuccessful"
                   alert = alert + 'MemberSetupRPA.ms_update_member: No. %d row is %s and is processed unsuccessful. ' % (idx + 1, action)
                   update_audit("Selenium Progress", "Completed...No. {0} row is {1} and is processed unsuccessfully.".format(idx+1, action), base)
              except Exception as error:
                print(error)
            idx += 1

            analytic_dm_by_row(row, base)
        

        self.go_to(MARC_DASHBOARD_ADDRESS)

        #print('Identify Row: {0}'.format(row))
       
        # Store good and bad
        out_table = pd.DataFrame(data=table)
        
        if all(flags):
          return out_table, "Update completed successfully"
        else:
          return out_table, alert

    def ms_get_content_by_field_name(self, where, field_name):
        """
        Get content of record in Member Setup.
        :param str where:           Which page to use
        :param str field_name:      Full name of the field
        :rtype:                     str
        """
        if where == 'member add':
            return self.get_content_on_current_page_by_field_name(MEMBER_ADD_FIELD_NAME_TO_XPATH_MAPPER,
                field_name)
        elif where == 'member search':
            return self.get_content_on_current_page_by_field_name(MEMBER_SEARCH_FIELD_NAME_TO_XPATH_MAPPER,
                field_name)
        elif where == 'member update':
            return self.get_content_on_current_page_by_field_name(RPA_INPT_MEMBER_UPDATE_TO_XPATH_MAPPER,
                field_name)
        elif where == 'inpt member policy add':
            return self.get_content_on_current_page_by_field_name(INPT_MEMBER_POLICY_ADD_FIELD_NAME_TO_XPATH_MAPPER,
                field_name)
        elif where == 'inpt member policy search':
            return self.get_content_on_current_page_by_field_name(INPT_MEMBER_POLICY_SEARCH_FIELD_NAME_TO_XPATH_MAPPER,
                field_name)
        elif where == 'inpt member policy update':
            return self.get_content_on_current_page_by_field_name(RPA_INPT_MEMBER_POLICY_UPDATE_TO_XPATH_MAPPER,
                field_name)
        elif where == 'inpt member policy view':
            return self.get_content_on_current_page_by_field_name(INPT_MEMBER_POLICY_VIEW_FIELD_NAME_TO_XPATH_MAPPER,
                field_name)
        else:
            self.logger.error('MemberSetupRPA.ms_get_content_by_field_name',
                'Used a wrong where %s. '
                'Choose one among %s.' % (where,
                    ', '.join(['member add',
                            'member search',
                            'member update',
                            'inpt member policy add',
                            'inpt member policy search',
                            'inpt member policy update',
                            'inpt member policy view',])))
            raise ValueError('Used a wrong where %s. '
                'Choose one among %s.' % (where,
                    ', '.join(['member add',
                            'member search',
                            'member update',
                            'inpt member policy add',
                            'inpt member policy search',
                            'inpt member policy update',
                            'inpt member policy view',])))

    def ms_input_content_by_field_name(self, where, field_name, base, content=None):
        """
        Input content of record in Member Setup.
        :param str where:           Which page to use
        :param str field_name:      Full name of the field
        :param str content:         Content to input
        :rtype:                     str
        """
        if where == 'member add':
            return self.input_on_current_page_by_field_name(MEMBER_ADD_FIELD_NAME_TO_XPATH_MAPPER,
                MEMBER_ADD_FIELD_NAME_TO_TYPE_MAPPER,
                field_name,
                base,
                content)
        elif where == 'member search':
            return self.input_on_current_page_by_field_name(RPA_INPT_MEMBER_UPDATE_TO_XPATH_MAPPER,
                MEMBER_SEARCH_FIELD_NAME_TO_TYPE_MAPPER,
                field_name,
                base,
                content)
        elif where == 'member update':
            return self.input_on_current_page_by_field_name(RPA_INPT_MEMBER_UPDATE_TO_XPATH_MAPPER,
                RPA_MEMBER_UPDATE_FIELD_NAME_TO_TYPE_MAPPER,
                field_name,
                base,
                content)
        elif where == 'inpt member policy add':
            return self.input_on_current_page_by_field_name(RPA_INPT_MEMBER_POLICY_UPDATE_TO_XPATH_MAPPER,
                INPT_MEMBER_POLICY_ADD_FIELD_NAME_TO_TYPE_MAPPER,
                field_name,
                base,
                content)
        elif where == 'plan search':
            return self.input_on_current_page_by_field_name(RPA_INPT_PLAN_UPDATE_TO_XPATH_MAPPER,
                RPA_INPT_PLAN_UPDATE_TO_TYPE_MAPPER,
                field_name,
                base,
                content)
        elif where == 'inpt member policy search':
            return self.input_on_current_page_by_field_name(INPT_MEMBER_POLICY_SEARCH_FIELD_NAME_TO_XPATH_MAPPER,
                INPT_MEMBER_POLICY_SEARCH_FIELD_NAME_TO_TYPE_MAPPER,
                field_name,
                base,
                content)
        elif where == 'inpt member policy update':
            return self.input_on_current_page_by_field_name(INPT_MEMBER_POLICY_UPDATE_FIELD_NAME_TO_XPATH_MAPPER,
                INPT_MEMBER_POLICY_UPDATE_FIELD_NAME_TO_TYPE_MAPPER,
                field_name,
                base,
                content)
        elif where == 'inpt member policy view':
            return self.input_on_current_page_by_field_name(INPT_MEMBER_POLICY_ADD_FIELD_NAME_TO_XPATH_MAPPER,
                INPT_MEMBER_POLICY_VIEW_FIELD_NAME_TO_TYPE_MAPPER,
                field_name,
                base,
                content)
        else:
            self.logger.error('MemberSetupRPA.ms_input_content_by_field_name',
                'Used a wrong where %s. '
                'Choose one among %s.' % (where,
                    ', '.join(['member add',
                            'member search',
                            'member update',
                            'inpt member policy add',
                            'inpt member policy search',
                            'inpt member policy upate',])))
            raise ValueError('Used a wrong where %s. '
                'Choose one among %s.' % (where,
                    ', '.join(['member add',
                            'member search',
                            'member update',
                            'inpt member policy add',
                            'inpt member policy search',
                            'inpt member policy upate',])))
