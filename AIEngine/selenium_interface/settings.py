#!/usr/bin/env python
#
# Created by Xixuan on 29/08/2018
#
# The file contains all settings.
#

# ============ Start of General Settings ============

# Paths
RELATIVE_PATH_TO_FOLDER_OF_LOGS = './rpa_logs'
BROWSER_CHOICE = "firefox"
DEFAULT_PATH_TO_DRIVER = 'selenium_interface/chromedriver.exe'
PATH_TO_FIREFOX_DRIVER = 'selenium_interface/geckodriver.exe'

# Addresses
MARC_ADDRESS = 'https://marcmy.asia-assistance.com/marcmy/'
MARC_DASHBOARD_ADDRESS = MARC_ADDRESS + 'pages/dashboard/dashboard_inpatient.xhtml'

MARC_MEMBER_SEARCH_ADDRESS = MARC_ADDRESS + 'pages/setup/member/member_search.xhtml'
MARC_MEMBER_ADD_ADDRESS = MARC_ADDRESS + 'pages/setup/member/member_add.xhtml'

MARC_INPT_MEMBER_POLICY_SEARCH_ADDRESS = MARC_ADDRESS + 'pages/setup/inptmember/memberpolicy_search.xhtml'
MARC_INPT_MEMBER_POLICY_ADD_ADDRESS = MARC_ADDRESS + 'pages/setup/inptmember/memberpolicy_add.xhtml'
MARC_INPT_MEMBER_POLICY_VIEW_ADDRESS = MARC_ADDRESS + 'pages/setup/inptmember/memberpolicy_view.xhtml'

MARC_POLICY_SEARCH_ADDRESS = MARC_ADDRESS + 'pages/setup/inptpolicy/inpt_policy_search.xhtml'
MARC_POLICY_ADD_ADDRESS = MARC_ADDRESS + 'pages/setup/inptpolicy/inpt_policy_create.xhtml'

MARC_INPT_BULK_UPLOAD_ADDRESS = MARC_ADDRESS + 'pages/dataimport/dataimp_inpt_runimport.xhtml'

MARC_INPT_PLAN_SEARCH_ADDRESS = MARC_ADDRESS + 'pages/setup/inptplan/inpt_plan_search.xhtml'

# Logs
LOG_PREFIX = 'rpa_log_'
LOG_FORMATTER = '%(asctime)s - %(levelname)s - %(message)s'

# Others  # TODO: Store password in encrpyted space
UID = 'simon.lee'
PW = 'M@rcpass2020'
DB_UID = "davidoso"
DB_PW = "root"
TIMEOUT = 7
SLEEP_TIME = 1
MAX_N_RETRY = 3
MONITOR_REFRESH_TIME = 4
ONLY_DISPLAY_ERRORS_IN_CONSOLE = False
TOGGLE_UPDATE_CONFIRMATION = True # Halt before update if true

# ============ End of General Settings ============

# ============ Start of Input/Output Files Settings ============

# Column names
COLS_TO_BE_DROPPED_IN_INPUT_XLSX_FOR_MS = [
    'Import Type'
]
COLS_TO_BE_DROPPED_IN_INPUT_XLSX_FOR_PM = [
    'Action'
]
COLS_TO_BE_LOWERCASED_IN_INPUT_XLSX_FOR_MS = [
    'status', 'action'
]
COLS_TO_BE_LOWERCASED_IN_INPUT_XLSX_FOR_PM = [
    'move type',
]

MEMBER_HANDSHAKE_LAYOUT = [
    'no',
    'import type',
    'member full name',
    'address 1',
    'address 2',
    'address 3',
    'address 4',
    'gender',
    'dob',
    'nric',
    'otheric',
    'external ref id aka client',
    'internal ref id (aka aan)',
    'employee id',
    'marital status',
    'race',
    'phone',
    'vip',
    'special condition',
    'relationship',
    'principal int ref id (aka aan)',
    'principal ext ref id (aka client)',
    'principal name',
    'principal nric',
    'principal other ic',
    'program id',
    'policy type',
    'policy num',
    'policy eff date',
    'policy end date',
    'previous policy num',
    'previous policy end date',
    'policy owner',
    'external plan code',
    'internal plan code id',
    'first join date',
    'plan attach date',
    'plan expiry date',
    'insurer branch',
    'insurer agency',
    'insurer agency code',
    'insurer mco fees',
    'ima service?',
    'ima limit',
    'date received by aan',
    'cancellation date',
    'status',
    'action',
    'marc',
    'card',
    'listing'
]

MARC_LAYOUT = [
    'No',
    'Import Type',
    'Member Full Name',
    'Address 1',
    'Address 2',
    'Address 3',
    'Address 4',
    'Gender',
    'DOB',
    'NRIC',
    'OtherIC',
    'Ext Ref',
    'Internal Ref Id (aka AAN)',
    'Employee ID',
    'Marital Status',
    'Race',
    'Phone',
    'VIP',
    'Special Condition',
    'Relationship',
    'Principal Int Ref Id (aka AAN)',
    'Principal Ext Ref Id (aka Client)',
    'Principal Name',
    'Principal NRIC',
    'Principal Other Ic',
    'Program Id',
    'Policy Type',
    'Policy Num',
    'Policy Eff Date',
    'Policy End Date',
    'Previous Policy Num',
    'Previous Policy End Date',
    'Policy Owner',
    'External Plan Code',
    'Internal Plan Code Id',
    'First Join Date',
    'Plan Attach Date',
    'Plan Expiry Date',
    'Insurer Branch',
    'Insurer Agency',
    'Insurer Agency Code',
    'Insurer MCO Fees',
    'IMA Service?',
    'IMA Limit',
    'Date Received by AAN',
    'Cancellation Date',
    'Insurer Agency Tel',
    'Benefit Code'
]

LIST_LAYOUT = [
    'No',
    'Plan',
    'Member Full Name',
    'IC No',
    'Policy Num',
    'Policy Eff Date',
    'Policy End Date',
    'Policy Owner',
    'Principal Name',
    'Date Received by AAN',
    'file_name'
]

CARD_LAYOUT = [
    'No',
    'Member Full Name',
    'Address 1',
    'Address 2',
    'Address 3',
    'Address 4',
    'Relationship',
    'IC No',
    'Plan',
    'Policy Num',
    'Policy Eff Date',
    'Policy End Date',
    'Policy Owner',
    'Principal Name',
    'Employee ID/Member No',
    'Sorting1',
    'Sorting2',
    'Sorting3',
    'Sorting4'

]

RPA_LAYOUT = [
    'No',
    'Import Type',
    'Member Full Name',
    'Address 1',
    'Address 2',
    'Address 3',
    'Address 4',
    'Gender',
    'DOB',
    'NRIC',
    'OtherIC',
    'External Ref Id aka Client',
    'Internal Ref Id (aka AAN)',
    'Employee ID',
    'Marital Status',
    'Race',
    'Phone',
    'VIP',
    'Special Condition',
    'Relationship',
    'Principal Int Ref Id (aka AAN)',
    'Principal Ext Ref Id (aka Client)',
    'Principal Name',
    'Principal NRIC',
    'Principal Other Ic',
    'Program Id',
    'Policy Type',
    'Policy Num',
    'Policy Eff Date',
    'Policy End Date',
    'Previous Policy Num',
    'Previous Policy End Date',
    'Policy Owner',
    'External Plan Code',
    'Internal Plan Code Id',
    'First Join Date',
    'Plan Attach Date',
    'Plan Expiry Date',
    'Insurer Branch',
    'Insurer Agency',
    'Insurer Agency Code',
    'Insurer MCO Fees',
    'IMA Service?',
    'IMA Limit',
    'Date Received by AAN',
    'Cancellation Date',
    'STATUS',
    'ACTION',
    'MARC',
    'CARD',
    'LISTING'
]

POLICY_HANDSHAKE_LAYOUT = [
    'no.',
    'action',
    'policy creation type',
    'enabled',
    'ext insurer code',
    'ext unique code',
    'policy num',
    'policy owner name',
    'policy owner local name',
    'policy owner nric',
    'policy owner gender',
    'policy owner email',
    'policy effective date',
    'policy expiry date',
    'insurer branch',
    'agency name',
    'agency local name',
    'agency code',
    'agency tel',
    'agency email',
    'agency location',
    'policy remarks',
    'move type',
    'move start date',
    'move end date',
    'move remarks',
    'int plan code (id)',
    'ext plan code',
    'plan name',
    'outcome'
]

# Others
MAX_N_TRIES_TO_FIND_HEADER = 4
KEY_WORD_TO_FIND_HEADER = 'policy num'
OUTPUT_FILE_PREFIX = 'rpa_'

# ============ End of Input/Output Files Settings ============

# ============ Start of Login/Logout Settings ============

LOGIN_FIELD_NAME_TO_XPATH_MAPPER = {
    'login': './/input[@title="login"]',
    'password': './/input[@title="password"]',
}

# ============ End of Login/Logout Settings ============

# RPA 3.0 Matrix Requirements
# Update only 
# name, nric, other id, dob, ext ref id, employee id, relationship, principal, plan, first join date, plan attach date
RPA_INPT_MEMBER_POLICY_UPDATE_TO_XPATH_MAPPER = {
    #INPT MEMBER POLICY-PLAN SETUP - RPA 3.0
    'employee id': '/html[1]/body[1]/div[1]/div[3]/form[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[3]/td[2]/input[1]',
    'first join date': '//*[@id="mmTab:imppFirstJoinDate_input"]',
    'plan attach date': '//*[@id="mmTab:imppPlanAttachDate_input"]',
    'relationship': '//*[@id="mmTab:imppRelationship"]',
    'external ref id aka client': '/html[1]/body[1]/div[1]/div[3]/form[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[2]/input[1]',
    #plan
    'member plan add': './/a[contains(@onclick,"lookup_inptplan")]',
    'plan lookup plan ext ref':'//*[@id="lookupInptPlanExtRef"]',
    'plan lookup plan int ref':'//input[@id="lookupInptPlanId"]',
    'plan int ref': '//*[@id="lookupInptPlanSearchResult:0:j_idt831"]',
    'plan lookup search':'/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/form[1]/table[1]/tbody[1]/tr[3]/td[5]/button[1]',
    #principal
    'principal add': './/a[contains(@onclick,"lookup_inptprincipal")]',
    'principal lookup other ic': '/html[1]/body[1]/div[1]/div[3]/div[4]/div[2]/form[1]/table[1]/tbody[1]/tr[1]/td[4]/input[1]',
    'principal lookup nric': '/html[1]/body[1]/div[1]/div[3]/div[4]/div[2]/form[1]/table[1]/tbody[1]/tr[2]/td[4]/input[1]',
    'principal lookup pol end date': '//*[@id="lookupInptPrincipalPolicyExpDate_input"]',
    'principal fullname': '//*[@id="lookupPrincipalSearchResult:0:j_idt859"]',
    'principal lookup search': '/html[1]/body[1]/div[1]/div[3]/div[4]/div[2]/form[1]/table[1]/tbody[1]/tr[4]/td[2]/button[1]',
    'update': '/html[1]/body[1]/div[1]/div[3]/form[1]/button[1]',
    'goto search': '//button[@id="j_idt775"]'
    }

RPA_INPT_MEMBER_UPDATE_TO_XPATH_MAPPER = {
#UPDATE MEMBER - RPA 3.0
    'member id':'/html[1]/body[1]/div[1]/div[3]/form[1]/table[1]/tbody[1]/tr[1]/td[2]/input[1]',
    'member full name': '/html[1]/body[1]/div[1]/div[3]/form[1]/div[1]/table[1]/tbody[1]/tr[2]/td[2]/input[1]',
    'dob':'/html[1]/body[1]/div[1]/div[3]/form[1]/div[1]/table[1]/tbody[1]/tr[3]/td[4]/span[1]/input[1]',
    'nric': '/html[1]/body[1]/div[1]/div[3]/form[1]/div[1]/table[1]/tbody[1]/tr[4]/td[2]/input[1]',
    'otheric': '/html[1]/body[1]/div[1]/div[3]/form[1]/div[1]/table[1]/tbody[1]/tr[5]/td[2]/input[1]',
    'update': '/html[1]/body[1]/div[1]/div[3]/form[1]/div[1]/button[2]',
    'search': '/html[1]/body[1]/div[1]/div[3]/form[1]/table[1]/tbody[1]/tr[4]/td[2]/button[1]',
    'goto search': '/html[1]/body[1]/div[1]/div[3]/form[1]/div[1]/button[1]'
    }

RPA_INPT_PLAN_UPDATE_TO_XPATH_MAPPER = {
   'int plan code(id)': '//*[@id="j_idt42"]',
   'ext plan code' : '//*[@id="j_idt60"]',
   'search': '//*[@id="j_idt56"]',
   'update': '//*[@id="j_idt1287"]'
}

# ============ Start of Member Setup Settings ============

IGNORED_ACTIONS = {
    'chg::bankname_bankaccount_bankdetail', 'chg::category', 'chg::costcentre_division_department',
    'chg::email', 'new::hire',  'new::pending', 'new::policy', 'new::reprint_print',
    'takeover::donotprint', 'takeover::vip_vvip', 'trf::reprint_print', 'chg::reprint_print', 'chg::'
}

MEMBER_SETUP_EQUIVALENCE_MAPPER = {
	# ('OtherIc', ): 'other ic'
    #('member full name', ): 'full name'
}

INPT_MEMBER_POLICY_SEARCH_FIELD_NAME_TO_XPATH_MAPPER = {
    'impp id': './/input[@id="imppId"]', # Not in handshake
    'member full name': './/td[text()="Full Name:"]/../td[2]/input',
    'member local name': './/td[text()="Local Name:"]/../td[2]/input', # Not in handshake
    'nric': './/td[text()="NRIC : "]/../td[2]/input',
    'otheric': './/td[text()="Other IC:"]/../td[2]/input',
    'gender': './/td[text()="Gender:"]/../td[4]/select',
    'employee id': './/td[text()="Employee Id:"]/../td[4]/input',
    'dob': './/td[text()="DOB:"]/../td[4]/span/input',
    'external ref id (aka client)': './/td[text()="Ext Ref Id:"]/../td[4]/input',
    'policy num': './/td[text()="Policy Num:"]/../td[4]/input',
    'plan name': './/td[text()="Plan Name:"]/../td[6]/input', # Not in handshake
    'internal plan code id': './/td[text()="Int Plan Id:"]/../td[6]/input',
    'external plan code': './/td[text()="Ext Plan Code:"]/../td[6]/input',
    'agent name': './/td[text()="Agent Name:"]/../td[6]/input',
    'agent local name': './/td[text()="Agent Local Name:"]/../td[6]/input',
    'agent code': './/td[text()="Agent Code:"]/../td[6]/input',
    'policy eff date': './/td[text()="Pol Eff Date:"]/../td[8]/span/input',
    'pol expiry date': './/td[text()="Pol Expiry Date:"]/../td[8]/span/input',
    'plan attach date': './/td[text()="Plan Attach Date:"]/../td[8]/span/input',
    'plan expiry date': './/td[text()="Plan Expiry Date:"]/../td[8]/span/input',
    'plan ext date': './/td[text()="Plan Ext Date:"]/../td[8]/span/input', # Not in handshake
    'insurer': './/td[text()="Insurer:"]/../td/select',
    'show principal only': './/td[text()="Show principal only"]/input',
    'show enabled only': './/td[text()="Show enabled only"]/input',
    'search': './/button[@id="searchBtn"]'
}

INPT_MEMBER_POLICY_SEARCH_FIELD_NAME_TO_TYPE_MAPPER = {
    'impp id': 'input',
    'member full name': 'input',
    'member local name': 'input',
    'nric': 'input',
    'otheric': 'input',
    'gender': 'select',
    'employee id': 'input',
    'dob': 'date_input',
    'external ref id (aka client)': 'input',
    'policy num': 'input',
    'plan name': 'input',  # Not in handshake
    'internal plan code id': 'input',
    'external plan code': 'input',
    'agent name': 'input',
    'agent local name': 'input',
    'agent code': 'input',
    'policy eff date': 'date_input',
    'pol expiry date': 'date_input',
    'plan attach date': 'date_input',
    'plan expiry date': 'date_input',
    'plan ext date': 'date_input',
    'insurer': 'select',
    'show principal only': 'checkbox',
    'show enabled only': 'checkbox',

    'search': 'button'
}

INPT_MEMBER_POLICY_VIEW_FIELD_NAME_TO_XPATH_MAPPER = {
    'edit': './/button[@title="Edit"]',

    # Member info
    'mm id': '//td[contains(text(), \'Mm Id\')]/../td[6]',

    #plan code id
    'Int Plan Code (Id)': '//td[contains(text(),"Int Plan Code (Id):")]/../td[2]',

    # Plan
    'plan ext plan code': './/div[@id="plan_info_container_left"]/table/tbody/tr/td[text()="Ext Plan Code:"]/../td[2]',

    # Upgrade inpt member plan
    'upgrade': './/button[@title="Upgrade"]',
    'upgrade plan policy id': './/td[text()="Policy Id:"]/../td[2]/input',
    'upgrade plan policy num': './/td[text()="Policy Num:"]/../td[2]/input',
    'upgrade plan plan name': './/td[text()="Plan Name:"]/../td[2]/input',
    'upgrade plan int plan code': './/td[text()="Int Plan Code:"]/../td[4]/input',
    'upgrade plan ext plan code': './/td[text()="Ext Plan Code:"]/../td[4]/input',
    'upgrade plan upgrade date': './/td[text()="Upgrade Date:"]/../td/span/input',
    'upgrade plan search': './/span[text()="Search"]/..',
}

INPT_MEMBER_POLICY_VIEW_FIELD_NAME_TO_TYPE_MAPPER = {
    'edit': 'button',

    # Plan
    'plan ext plan code': '',

    # Upgrade inpt member plan
    'upgrade': 'onclick',
    'upgrade plan policy id': 'input',
    'upgrade plan policy num': 'input',
    'upgrade plan plan name': 'input',
    'upgrade plan int plan code': 'input',
    'upgrade plan ext plan code': 'input',
    'upgrade plan upgrade date': 'date_input',
    'upgrade plan search': 'button',
}

MEMBER_SEARCH_FIELD_NAME_TO_XPATH_MAPPER = {
    'member id': './/*[text()="Member Id:"]/../../td/input',
    'client': './/td[text()="Client:"]/../td/select',
    'dob': './/td[text()="DOB:"]/../td/span/input',
    'full name': './/td[text()="Full Name:"]/../td[2]/input',
    'national id': './/td[text()="National Id:"]/../td[4]/input',
    'created date from': './/td[text()="Created Date:"]/../td/span[1]/input',
    'created date to': './/td[text()="Created Date:"]/../td/span[2]/input',
    'other national id': './/td[text()="Other National Id:"]/../td[4]/input',
    'search': './/span[contains(@class,"ui-icon-search")]/..'
}

MEMBER_SEARCH_FIELD_NAME_TO_TYPE_MAPPER = {
    'member id': 'input',
    'client': 'select',
    'dob': 'date_input',
    'full name': 'input',
    'national id': 'input',
    'created date from': 'date_input',
    'created date to': 'date_input',
    'other national id': 'input',
    'search': 'button'
}

MEMBER_UPDATE_FIELD_NAME_TO_XPATH_MAPPER = {
    'title': './/td[text()="*Title:"]/../td[4]/input',
    'full name': './/td[text()="*Full Name:"]/../td[2]/input',
    'client': './/td[text()="*Client:"]/../td/select',
    'bank name': './/td[text()="Bank Name:"]/../td[6]/input',
    'dob': './/td[text()="*DOB:"]/../td/span/input',
    'bank code': './/td[text()="Bank Code:"]/../td[6]/input',
    'national id': './/td[text()="*National Id:"]/../td[2]/input',
    'gender': './/td[text()="Gender:"]/../td/select',
    'bank swift code': './/td[text()="Bank Swift Code:"]/../td[6]/input',
    'other national id': './/td[text()="Other National Id:"]/../td[2]/input',
    'marital status': './/td[text()="Marital Status:"]/../td/select',
    'bank payee name': './/td[text()="Bank Payee Name:"]/../td[6]/input',
    'mm ext ref': './/td[text()="Mm Ext Ref:"]/../td[2]/input',
    'prefer language': './/td[text()="Prefer Language :"]/../td/select',
    'bank account num': './/td[text()="Bank Account Num:"]/../td[6]/input',
    'email (from data import)': './/td[text()="Email (From Data Import):"]/../td[4]/input',
    'update': './/span[text()="Update"]/..'
}

RPA_MEMBER_UPDATE_FIELD_NAME_TO_TYPE_MAPPER = {
    'title': 'input',
    'member full name': 'input',
    'dob': 'date_input',
    'nric': 'input',
    'otheric': 'input',
    'update': 'button'
}

RPA_INPT_PLAN_UPDATE_TO_TYPE_MAPPER = {
    'int plan code(id)':'input',
    'search':'button',
    'ext plan code':'input'
}

MEMBER_UPDATE_FIELD_NAME_TO_TYPE_MAPPER = {
    'title': 'input',
    'full name': 'input',
    'client': 'select',
    'bank name': 'input',
    'dob': 'date_input',
    'bank code': 'input',
    'national id': 'input',
    'gender': 'select',
    'bank swift code': 'input',
    'other national id': 'input',
    'marital status': 'select',
    'bank payee name': 'input',
    'mm ext ref': 'input',
    'prefer language': 'select',
    'bank account num': 'input',
    'email (from data import)': 'input',
    'update': 'button'
}

INPT_MEMBER_POLICY_ADD_FIELD_NAME_TO_XPATH_MAPPER = {
    #INPT MEMBER POLICY-PLAN SETUP - RPA 3.0
    'employee id': '//*[@id="mmTab:j_idt77"]',
    'first join date': '//*[@id="mmTab:imppFirstJoinDate"]',
    'plan attach date': '//*[@id="mmTab:imppPlanAttachDate_input"]',
    'relationship': '//*[@id="mmTab:imppRelationship"]',
    'external ref id (aka client)': '//*[@id="mmTab:j_idt69"]',
    #UPDATE MEMBER - RPA 3.0
    'member full name': '//*[@id="j_idt48"]',
    'dob':'//*[@id="j_idt60_input"]',
    'nric': '//*[@id="j_idt66"]',
    'otheric': '//*[@id="j_idt77"]',
    #Update Plan
    'upgrade plan ext plan code': '//*[@id="j_idt60"]',
    # Member lookup
    'lookup member': './/a[contains(@onclick,"lookup_member")]',
    'member local name': './/form[@name="lookupMemberForm"]/table/tbody/tr[2]/td[2]/input', # Not in handshake
    'member lookup search': './/button[contains(@onclick,"lookupMemberForm")]',
    # Input member policy plan
    'special condition': './/td[text()="Spec Cond:"]/../td[2]/textarea',
    'axa benefit code': './/td[text()="Axa Benefit Code:"]/../td[2]/input', # Not in handshake
    'address 1': './/td[text()="Address 1:"]/../td[4]/input',
    'address 2': './/td[text()="Address 2:"]/../td[4]/input',
    'address 3': './/td[text()="Address 3:"]/../td[4]/input',
    'address 4': './/td[text()="Address 4:"]/../td[4]/input',
    'exclusion': './/td[text()="Exclusion:"]/../td[4]/textarea', # Not in handshake
    'postal code': './/td[text()="Postal Code:"]/../td[6]/input', # Not in handshake
    'city': './/td[text()="City:"]/../td[6]/input', # Not in handshake
    'country': './/td[text()="Country:"]/../td[6]/input', # Not in handshake
    'phone': './/td[text()="Phone:"]/../td[6]/input',
    'impairment': './/td[text()="Impairment:"]/../td[6]/textarea', # Not in handshake
    'dep num': './/td[text()="Dep Num:"]/../td[8]/input', # Not in handshake
    'insurer mco fees': './/td[text()="MCO Fee:"]/../td[8]/input',
    'school name': './/td[text()="School Name:"]/../td[8]/input', # Not in handshake
    'school state': './/td[text()="School State:"]/../td[8]/input', # Not in handshake
    'plan mvmt': './/td[text()="Plan Mvmt"]/../td[8]/textarea', # Not in handshake
    'vip': './/td[contains(text(),"VIP")]/input[1]',
    'enabled': './/td[contains(text(),"VIP")]/input[2]', # Not in handshake
    'national id': './/td[contains(text(),"VIP")]/input[3]', # Not in handshake
    # Input policy details
    'creation type': './/td[text()="*Creation Type:"]/../td/select', # Not in handshake
    'member policy add': './/a[contains(@onclick,"lookup_inptpolicy")]',
    'policy lookup id': './/form[@id="lookupInptPolicyForm"]/table/tbody/tr[1]/td[2]/input',
    'policy lookup num': './/form[@id="lookupInptPolicyForm"]/table/tbody/tr[2]/td[2]/input',
    'policy lookup effective date': './/form[@id="lookupInptPolicyForm"]/table/tbody/tr[1]/td[4]/span/input',
    'policy lookup expiry date': './/form[@id="lookupInptPolicyForm"]/table/tbody/tr[2]/td/span/input',
    'policy lookup search': './/form[@id="lookupInptPolicyForm"]/table/tbody/tr[2]/td[5]/button',

    'policy num': './/td[text()="*Policy Num:"]/../td[2]/input',
    'policy owner': './/td[text()="Policy Owner:"]/../td[2]/input',
    'previous policy num': './/td[text()="Prev Pol Num:"]/../td[2]/input',
    'insurer branch': './/td[text()="Insurer Branch:"]/../td[4]/input',
    'insurer agency code': './/td[text()="Agency Code:"]/../td[4]/input',
    'insurer agency': './/td[text()="Agency Name:"]/../td[4]/input',
    'policy eff date': './/td[text()="*Policy Eff Date:"]/../td[6]/span/input',
    'policy end date': './/td[text()="*Policy Expiry Date:"]/../td[6]/span/input',
    'policy ext date': './/td[text()="Policy Ext Date:"]/../td[6]/span/input', # Not in handshake
    'policy remarks': './/td[text()="Policy Remarks:"]/../td[8]/textarea',

    # Input plan details
    'member plan add': './/a[contains(@onclick,"lookup_inptplan")]',
    'plan lookup policy id': './/form[@id="lookupInptPlanForm"]/table/tbody/tr[1]/td[2]/input', # Not in handshake
    'plan lookup policy num': './/form[@id="lookupInptPlanForm"]/table/tbody/tr[2]/td[2]/input',
    'plan lookup plan int ref': './/form[@id="lookupInptPlanForm"]/table/tbody/tr[1]/td[4]/input',
    'plan lookup plan ext ref': './/form[@id="lookupInptPlanForm"]/table/tbody/tr[2]/td[4]/input',
    'plan lookup plan name': './/form[@id="lookupInptPlanForm"]/table/tbody/tr[3]/td[4]/input', # Not in handshake
    'plan lookup search': './/form[@id="lookupInptPlanForm"]/table/tbody/tr[3]/td[5]/button',
    'plan expiry date': './/td[text()="*Plan Expiry Date:"]/../td[6]/span/input',

    # Principal details
    'principal add': './/a[contains(@onclick,"lookup_inptprincipal")]',
    'principal lookup impp id': './/form[@id="lookupInptPrincipalForm"]/table/tbody/tr[1]/td[2]/input', # Not in handshake
    'principal lookup full name': './/form[@id="lookupInptPrincipalForm"]/table/tbody/tr[2]/td[2]/input',
    'principal lookup local name': './/form[@id="lookupInptPrincipalForm"]/table/tbody/tr[3]/td[2]/input', # Not in handshake
    'principal lookup other ic': './/form[@id="lookupInptPrincipalForm"]/table/tbody/tr[1]/td[4]/input',
    'principal lookup nric': './/form[@id="lookupInptPrincipalForm"]/table/tbody/tr[2]/td[4]/input',
    'principal lookup ext ref': './/form[@id="lookupInptPrincipalForm"]/table/tbody/tr[3]/td[4]/input',
    'principal lookup policy num': './/form[@id="lookupInptPrincipalForm"]/table/tbody/tr[1]/td[6]/input',
    'principal lookup pol end date': './/form[@id="lookupInptPrincipalForm"]/table/tbody/tr[2]/td[6]/span/input',
    'principal lookup search': './/form[@id="lookupInptPrincipalForm"]/table/tbody/tr[4]/td/button',
    'principal full name': './/a[contains(@onclick,"lookup_inptprincipal")]/../../td[4]/input',
    'principal nric': './/a[contains(@onclick,"lookup_inptprincipal")]/../../td[8]/input',
    'principal other ic': './/a[contains(@onclick,"lookup_inptprincipal")]/../../td[10]/input',

    'update': '//*[@id="updateButton"]',
    'save': './/button[@id="saveButton"]',
    'goto search': '//*[@id="j_idt775"]'
}


INPT_MEMBER_POLICY_ADD_FIELD_NAME_TO_TYPE_MAPPER = {
    # Member lookup
    'lookup member': 'onclick',
    'member full name': 'input',
    'member local name': 'input',
    'nric': 'input',
    'otheric': 'input',
    'member lookup search': 'button',
    # Input member policy plan
    'external ref id aka client': 'input',
    'employee id': 'input',
    'relationship': 'select',
    'special condition': 'input',
    'axa benefit code': 'input',
    'address 1': 'input',
    'address 2': 'input',
    'address 3': 'input',
    'address 4': 'input',
    'exclusion': 'input',
    'postal code': 'input',
    'city': 'input',
    'country': 'input',
    'phone': 'input',
    'impairment': 'input',
    'dep num': 'input',
    'insurer mco fees': 'input',
    'school name': 'input',
    'school state': 'input',
    'plan mvmt': 'input',
    'vip': 'checkbox',
    'enabled': 'checkbox',
    'national id': 'checkbox',
    # Input policy details
    'creation type': 'select',
    'member policy add': 'onclick',
    'policy lookup id': 'input',
    'policy lookup num': 'input',
    'policy lookup effective date': 'date_input',
    'policy lookup expiry date': 'date_input',
    'policy lookup search': 'button',

    'policy num': 'input',
    'policy owner': 'input',
    'previous policy num': 'input',
    'insurer branch': 'input',
    'insurer agency code': 'input',
    'insurer agency': 'input',
    'policy eff date': 'date_input',
    'policy end date': 'date_input',
    'policy ext date': 'date_input',
    'policy remarks': 'input',

    # Input plan details
    'member plan add': 'onclick',
    'plan int ref': 'onclick',
    'plan lookup policy id': 'input',
    'plan lookup policy num': 'input',
    'plan lookup plan int ref': 'input',
    'plan lookup plan ext ref': 'input',
    'plan lookup plan name': 'input',
    'plan lookup search': 'button',
    'first join date': 'date_input',
    'plan attach date': 'date_input',
    'plan expiry date': 'date_input',
    'ext plan code': 'input',
    'plan search': 'button',

    # Principal details
    'principal add': 'onclick',
    'principal lookup impp id': 'input',
    'principal lookup full name': 'input',
    'principal lookup local name': 'input',
    'principal lookup other ic': 'input',
    'principal lookup nric': 'input',
    'principal lookup ext ref': 'input',
    'principal lookup policy num': 'input',
    'principal lookup pol end date': 'date_input',
    'principal lookup search': 'button',

    'update': 'button',
    'save': 'button',
    'goto search': 'button'
}

MEMBER_ADD_FIELD_NAME_TO_XPATH_MAPPER = {
    'member full name': './/td[text()="*Full Name:"]/../td[2]/input',
    'chinese name': './/td[text()="Chinese Name:"]/../td[2]/input', # Not in handshake
    'national id': './/td[text()="*National Id:"]/../td[2]/input', # Not in handshake (!)
    'other national id': './/td[text()="Other National Id:"]/../td[2]/input', # Not in handshake
    'mm ext ref': './/td[text()="Mm Ext Ref:"]/../td[2]/input', # External ref id (aka client)???
    'title': './/td[text()="*Title:"]/../td[4]/input', # Not in handshake
    'client': './/td[text()="*Client:"]/../td[4]/select', # Principal id or name?
    'dob': './/td[text()="*DOB:"]/../td[4]/span/input',
    'gender': './/td[text()="Gender:"]/../td/select',
    'marital status': './/td[text()="Marital Status:"]/../td/select',
    'prefer language': './/td[text()="Prefer Language :"]/../td/select', # Not in handshake
    'bank name': './/td[text()="Bank Name:"]/../td[6]/input', # Not in handshake
    'bank code': './/td[text()="Bank Code:"]/../td[6]/input', # Not in handshake
    'bank swift code': './/td[text()="Bank Swift Code:"]/../td[6]/input', # Not in handshake
    'bank payee name': './/td[text()="Bank Payee Name:"]/../td[6]/input', # Not in handshake
    'bank account num': './/td[text()="Bank Account Num:"]/../td[6]/input', # Not in handshake
    'email': './/td[text()="Email (From Data Import):"]/../td[4]/input', # Not in handshake
    'blacklisted': './/td[contains(text(),"blacklisted")]/../td/input',
    'goto search': './/span[text()="Goto Search"]/..',
    'reset': './/span[text()="Reset"]/..',
    'save': './/span[text()="Save"]/..',
}

MEMBER_ADD_FIELD_NAME_TO_TYPE_MAPPER = {
    'member full name': 'input',
    'chinese name': 'input',
    'national id': 'input',
    'other national id': 'input',
    'mm ext ref': 'input',
    'title': 'input',
    'client': 'select',
    'dob': 'date_input',
    'gender': 'select',
    'marital status': 'select',
    'prefer language': 'select',
    'bank name': 'input',
    'bank code': 'input',
    'bank swift code': 'input',
    'bank payee name': 'input',
    'bank account num': 'input',
    'email': 'input',
    'blacklisted': 'checkbox',
    'goto search': 'button',
    'reset': 'button',
    'save': 'button',
}

INPT_MEMBER_POLICY_UPDATE_FIELD_NAME_TO_XPATH_MAPPER = INPT_MEMBER_POLICY_ADD_FIELD_NAME_TO_XPATH_MAPPER

INPT_MEMBER_POLICY_UPDATE_FIELD_NAME_TO_TYPE_MAPPER = INPT_MEMBER_POLICY_ADD_FIELD_NAME_TO_TYPE_MAPPER

# ============ End of Member Setup Settings ============

# ============ Start of Policy Manipulation Settings ============

INPT_POLICY_EQUIVALENCE_MAPPER = {
    ('enabled', 'policy enabled'): 'enabled',
    ('policy remark', ): 'policy remarks',
    ('remarks', 'movement remark', 'movement remarks', 'move remark'): 'move remarks',
    ('int plan id', 'int plan code'): 'int plan code (id)'
}

INPT_POLICY_CREATE_FIELD_NAME_TO_XPATH_MAPPER = {
    # Policy Details
    'policy creation type': './/td[text()="*Policy Creation Type:"]/../td/select',
    'policy num': './/td[text()="*Policy Num:"]/../td/input',
    'ext unique code': './/td[text()="Ext Unique Code:"]/../td/input',
    'ext insurer code': './/td[text()="Ext Insurer Code:"]/../td/input',
    'policy effective date': './/td[text()="*Policy Effective Date:"]/../td/span/input',
    'policy expiry date': './/td[text()="*Policy Expiry Date:"]/../td/span/input',
    'policy extension date': './/td[text()="Policy Extension Date:"]/../td/span/input',
    'policy cancellation date': './/td[text()="Policy Cancellation Date:"]/../td/span/input',
    'policy paid to date': './/td[text()="Policy Paid to Date:"]/../td/span/input',
    'policy reinstatement date': './/td[text()="Policy Reinstatement Date:"]/../td/span/input',
    'policy lg suspension date': './/td[text()="Policy LG Suspension Date:"]/../td/span/input',
    'policy termination': './/td[text()="Policy Termination:"]/../td/input',
    'suspend flag': './/td[text()="Suspend Flag:"]/../td/input',
    'policy remarks': './/td[text()="Remark:"]/../td/textarea',
    'enabled': './/td[text()="Enabled?"]/input[@type="checkbox"]',
    # Policy Owner
    'policy owner name': './/td[text()="Policy Owner Name:"]/../td/input',
    'policy owner local name': './/td[text()="Policy Owner Local Name:"]/../td/input',
    'policy owner nric': './/td[text()="Policy Owner NRIC:"]/../td/input',
    'policy owner email': './/td[text()="Policy Owner Email:"]/../td/input',
    'policy owner gender': './/td[text()="Policy Owner Gender:"]/../td/select',
    # AGENCY
    'insurer branch': './/td[text()="Insurer Branch:"]/../td/input',
    'agency location': './/td[text()="Agency Location:"]/../td/input',
    'agency name': './/td[text()="Agency Name:"]/../td/input',
    'agency local name': './/td[text()="Agency Local Name:"]/../td/input',
    'agency code': './/td[text()="Agency Code:"]/../td/input',
    'agency tel': './/td[text()="Agency Tel:"]/../td/input',
    'agency email': './/td[text()="Agency Email:"]/../td/input',
    # PREV POLICY
    'prev policy id': './/input[@id="inptPolPrevPolicyId"]',    # PF('prevpolicy_lookupbox').show()
    'prev policy num': './/input[@id="inptPolPrevPolicyNum"]',
    'prev policy eff date': './/input[@id="inptPolPrevPolicyEffDate"]',
    'prev policy exp date': './/input[@id="inptPolPrevPolicyExpDate"]',
    'prev policy ext date': './/input[@id="inptPolPrevPolicyExtDate"]',
    'prev lookup policy id': './/input[@id="lookupPrevInptPolicyId"]',
    'prev lookup policy effective date': './/input[@id="lookupPrevInptPolicyEffDate"]',
    'prev lookup policy num': './/input[@id="lookupPrevInptPolicyNum"]',
    'prev lookup policy expiry date': './/input[@id="lookupPrevInptPolicyExpDate"]',
    'prev lookup search': './/input[@id="lookupPrevInptPolicyEffDate"]/../../../tr/td/button',
    # ATTACH PLANS
    'int plan id': './/input[@id="inptPlanId"]',        # PF('plan_lookupbox').show()
    'int plan name': './/input[@id="inptPlanName"]',    # PF('plan_lookupbox').show()
    'lookup int plan code (id)': './/input[@id="lookupInptPlanId"]',
    'lookup ext plan code': './/input[@id="lookupInptPlanExtRef"]',
    'lookup plan name': './/input[@id="lookupInptPlanName"]',
    'lookup search': './/input[@id="lookupInptPlanName"]/../button',
    'add': './/input[@id="inptPlanName"]/../../td/button',
    # Final Actions
    'save': './/span[text()="Save"]/..',
    'clear': './/span[text()="Clear"]/..',
    'goto search': './/span[text()="Goto Search"]/..',
}

INPT_POLICY_CREATE_FIELD_NAME_TO_TYPE_MAPPER = {
    # Policy Details
    'policy creation type': 'select',
    'policy num': 'input',
    'ext unique code': 'input',
    'ext insurer code': 'input',
    'policy effective date': 'date_input',
    'policy expiry date': 'date_input',
    'policy extension date': 'date_input',
    'policy cancellation date': 'date_input',
    'policy paid to date': 'date_input',
    'policy reinstatement date': 'date_input',
    'policy lg suspension date': 'date_input',
    'policy termination': 'input',
    'suspend flag': 'input',
    'policy remarks': 'input',
    'enabled': 'checkbox',
    # policy Owner
    'policy owner name': 'input',
    'policy owner local name': 'input',
    'policy owner nric': 'input',
    'policy owner email': 'input',
    'policy owner gender': 'select',
    # AGENCY
    'insurer branch': 'input',
    'agency location': 'input',
    'agency name': 'input',
    'agency local name': 'input',
    'agency code': 'input',
    'agency tel': 'input',
    'agency email': 'input',
    # PREV POLICY
    'prev policy lookup': ' ',
    'prev lookup policy effective date': 'date_input',
    'prev lookup policy num': 'input',
    'prev lookup policy expiry date': 'date_input',
    'prev lookup search': 'button',
    # ATTACH PLANS
    'int plan id': 'input',
    'int plan name': 'input',
}

INPT_POLICY_SEARCH_FIELD_NAME_TO_XPATH_MAPPER = {
    'policy id': './/strong[text()="Policy Id:"]/../../td[2]/input',
    'policy owner name': './/td[text()="Policy Owner:"]/../td[4]/input',
    'policy num': './/td[text()="Policy Num:"]/../td[2]/input',
    'policy creation type': './/select[@id="PolicyType"]',
    'policy enabled': './/td[text()="Show enabled only"]/input',
    'search': './/span[text()="Search"]/..',
    'add New': './/span[text()="Add New"]/..'
}

INPT_POLICY_SEARCH_FIELD_NAME_TO_TYPE_MAPPER = {
    'policy id': 'input',
    'policy owner name': 'input',
    'policy num': 'input',
    'policy creation type': 'select',
    'policy enabled': 'checkbox',
    'search': 'button',
    'add New': 'button'
}

INPT_POLICY_UPDATE_FIELD_NAME_TO_XPATH_MAPPER = {
    # Policy Details
    'policy creation type': './/td[text()="*Policy Creation Type:"]/../td/select',
    'policy num': './/td[text()="*Policy Num:"]/../td/input',
    'ext unique code': './/td[text()="Ext Unique Code:"]/../td/input',
    'ext insurer code': './/td[text()="Ext Insurer Code:"]/../td/input',
    'policy effective date': './/td[text()="*Policy Effective Date:"]/../td/span/input',
    'policy expiry date': './/td[text()="*Policy Expiry Date:"]/../td/span/input',
    'policy extension date': './/td[text()="Policy Extension Date:"]/../td/span/input',
    'policy cancellation date': './/td[text()="Policy Cancellation Date:"]/../td/span/input',
    'policy paid to date': './/td[text()="Policy Paid to Date:"]/../td/span/input',
    'policy reinstatement date': './/td[text()="Policy Reinstatement Date:"]/../td/span/input',
    'policy lg suspension date': './/td[text()="Policy LG Suspension Date:"]/../td/span/input',
    'policy termination': './/td[text()="Policy Termination:"]/../td/input',
    'suspend flag': './/td[text()="Suspend Flag:"]/../td/input',
    'policy remarks': './/td[text()="Remark:"]/../td/textarea',
    'enabled': './/td[text()="Enabled?"]/../td/input',
    # Policy Owner
    'policy owner name': './/td[text()="Policy Owner Name:"]/../td/input',
    'policy owner local name': './/td[text()="Policy Owner Local Name:"]/../td/input',
    'policy owner nric': './/td[text()="Policy Owner NRIC:"]/../td/input',
    'policy owner email': './/td[text()="Policy Owner Email:"]/../td/input',
    'policy owner gender': './/td[text()="Policy Owner Gender:"]/../td/select',
    # AGENCY
    'insurer branch': './/td[text()="Insurer Branch:"]/../td/input',
    'agency location': './/td[text()="Agency Location:"]/../td/input',
    'agency name': './/td[text()="Agency Name:"]/../td/input',
    'agency local name': './/td[text()="Agency Local Name:"]/../td/input',
    'agency code': './/td[text()="Agency Code:"]/../td/input',
    'agency tel': './/td[text()="Agency Tel:"]/../td/input',
    'agency email': './/td[text()="Agency Email:"]/../td/input',
    # PREV POLICY
    'prev policy lookup': './/a[@onclick="PF(\'prevpolicy_lookupbox\').show()"]',
    'prev lookup policy id': './/*[@id="lookupPrevInptPolicyId"]',
    'prev lookup policy effective date': './/*[@id="lookupPrevInptPolicyEffDate"]',
    'prev lookup policy num': './/*[@id="lookupPrevInptPolicyNum"]',
    'prev lookup policy expiry date': './/*[@id="lookupPrevInptPolicyExpDate"]',
    'prev lookup search': './/*[@id="inptPrevPolicyLookupForm"]/table/tbody/tr/td/button[@type="submit"]',
    # INPT POLICY MOVEMENT
    'move type': './/*[@id="ipdType"]',
    'move start date': './/*[@id="ipdStartDate_input"]',
    'move end date': './/*[@id="ipdEndDate_input"]',
    'move remarks': './/td[text()="Remarks:"]/../td/textarea',
    'add movement': './/span[text()="Add Movement"]/..',
    # ATTACH PLANS
    'attach plan lookup': './/a[@onclick="PF(\'plan_lookupbox\').show()"]',
    'int plan code (id)': './/*[@id="lookupInptPlanId"]',
    'ext plan code': './/*[@id="lookupInptPlanExtRef"]',
    'plan name': './/*[@id="lookupInptPlanName"]',
    'search': './/*[@id="lookupInptPlanName"]/../button',
    'add': './/td[text()="Int Plan Id:"]/../td/button',
    # Final Actions
    'update': './/span[text()="Update"]/..',
    'goto search': './/span[text()="Goto Search"]/..',
}

INPT_POLICY_UPDATE_FIELD_NAME_TO_TYPE_MAPPER = {
    # Policy Details
    'policy creation type': 'select',
    'policy num': 'input',
    'ext unique code': 'input',
    'ext insurer code': 'input',
    'policy effective date': 'date_input',
    'policy expiry date': 'date_input',
    'policy extension date': 'date_input',
    'policy cancellation date': 'date_input',
    'policy paid to date': 'date_input',
    'policy reinstatement date': 'date_input',
    'policy lg suspension date': 'date_input',
    'policy termination': 'input',
    'suspend flag': 'input',
    'policy remarks': 'input',
    'enabled': 'checkbox',
    # Policy Owner
    'policy owner name': 'input',
    'policy owner local name': 'input',
    'policy owner nric': 'input',
    'policy owner email': 'input',
    'policy owner gender': 'input',
    # AGENCY
    'insurer branch': 'input',
    'agency location': 'input',
    'agency name': 'input',
    'agency local name': 'input',
    'agency code': 'input',
    'agency tel': 'input',
    'agency email': 'input',
    # PREV POLICY
    'prev policy lookup': 'onclick',
    'prev lookup policy id': 'input',
    'prev lookup policy effective date': 'date_input',
    'prev lookup policy num': 'input',
    'prev lookup policy expiry date': 'date_input',
    'prev lookup search': 'button',
    # INPT POLICY MOVEMENT
    'move type': 'select',
    'move start date': 'date_input',
    'move end date': 'date_input',
    'move remarks': 'input',
    'add movement': 'button',
    # ATTACH PLANS
    'attach plan lookup': 'onclick',
    'int plan code (id)': 'input',
    'ext plan code': 'input',
    'plan name': 'input',
    'search': 'button',
    'add': 'button',
    # Final Actions
    'update': 'button',
    'goto search': 'button',
}

# ============ End of Policy Manipulation Settings ============

# ================ Start of Bulk Import Settings ==============

BULK_IMPORT_FIELD_NAME_TO_XPATH_MAPPER = {
    'choose': './/span[text()="Choose"]/../input',
    'upload': './/button[contains(@class,"ui-fileupload-upload")]',
    'result': './/span[@class="ui-messages-info-detail"]'
}

BULK_IMPORT_FIELD_NAME_TO_TYPE_MAPPER = {
    'choose': 'file',
    'upload': 'button',
}

# ================ End of Bulk Import Settings ==============

# ================= General/Misc Settings ==================

LOADING_INDICATOR = './/div[@class=\"three-quarters\"]/../..'

# ================= Raw File Handler Settings ===============

PATH_TO_COMPANY_HANDLERS = "./rawHandler/handlers"

MARC_COLUMN_WIDTHS = {
    "A:A": 3,
    "B:B": 40,
    "C:C": 48,
    "D:D": 6,
    "E:E": 6,
    "F:F": 6,
    "G:G": 6,
    "H:H": 12,
    "I:I": 14,
    "J:J": 11,
    "K:K": 11,
    "L:L": 8,
    "M:M": 8,
    "N:N": 8,
    "O:O": 20,
    "P:P": 8,
    "Q:Q": 3,
    "R:R": 10,
    "S:S": 22,
    "T:T": 10,
    "U:U": 26,
    "V:V": 13,
    "W:W": 15,
    "X:X": 19,
    "Y:Y": 19,
    "Z:Z": 12,
    "AA:AA": 12,
    "AB:AB": 15,
    "AC:AC": 17,
    "AD:AD": 12,
    "AE:AE": 14,
    "AF:AF": 14,
    "AG:AG": 15,
    "AH:AH": 9,
    "AI:AI": 9,
    "AJ:AJ": 17,
    "AK:AK": 17,
    "AL:AL": 17,
    "AM:AM": 9,
    "AN:AN": 9,
    "AO:AO": 9,
    "AP:AP": 9,
    "AQ:AQ": 9,
    "AR:AR": 9,
    "AS:AS": 21,
    "AT:AT": 19,
    "AU:AU": 10,
    "AV:AV": 37,
    "AW:AW": 8,
    "AX:AX": 8,
    "AY:AY": 8
}

LISTING_COLUMN_WIDTHS = {
    "A:A": 1.43,
    "B:B": 7,
    "C:C": 30.86,
    "D:D": 8.29,
    "E:E": 6.86,
    "F:F": 7.43,
    "G:G": 6.14,
    "H:H": 30.43,
    "J:J": 8,
    "I:I": 27,
    "K:K": 8
}


MARC_QUERY = """SELECT impp.`imppId`, mm.`mmId`,impp.`imppEmployeeId`, mm.`mmFullName`, mm.`mmDOB`, mm.`mmMaritalStatus`, mm.`mmNRIC`, mm.`mmOtherIc`, mm.`mmGender`, impp.`imppRelationship`, impp.`imppExternalRefId`, impp.`imppVIP`,
CASE WHEN impp.`imppPrincipalId` IS NULL THEN impp.`imppId` ELSE impp.`imppPrincipalId` END AS "imppPrincipalId",
CASE WHEN impp.`imppPrincipalId` IS NULL THEN mm.`mmFullName` ELSE prpmm.mmFullName END AS "Principal Name",
CASE WHEN impp.`imppPrincipalId` IS NULL THEN mm.`mmNRIC` ELSE prpmm.mmNRIC END AS "Principal IC",
CASE WHEN impp.`imppPrincipalId` IS NULL THEN mm.`mmOtherIc` ELSE prpmm.mmOtherIc END AS "Principal Other IC",
CASE WHEN impp.`imppPrincipalId` IS NULL THEN impp.`imppExternalRefId` ELSE prpim.imppExternalRefId END AS "Principal Ext Ref",
impp.`imppFirstJoinDate`, impp.`imppPlanAttachDate`, impp.`imppPlanExpiryDate`, impp.`imppCancelDate`, pol.`inptPolPolicyNum`, pol.`inptPolOwner`, polplan.ippInptpolicyId, `inptPlanExtCode`,`inptPlanName`, polplan.ippInptPlanId, impp.`imppInptPolicyPlanId`,impp.`imppCreatedDate`, impp.`imppSpecialConditions`, impp.`imppAddress4`
FROM member mm INNER JOIN `inpt_member_policyplan` impp ON impp.`imppMemberId` = mm.mmid
INNER JOIN `inpt_policy_plan` polplan ON polplan.`ippId` = impp.`imppInptPolicyPlanId`
INNER JOIN inpt_plan plan ON plan.`inptPlanId` = polplan.`ippInptPlanId`
INNER JOIN `inpt_policy` pol ON pol.`inptPolId` = polplan.`ippInptPolicyId`
LEFT OUTER JOIN `inpt_member_policyplan` prpim ON impp.imppPrincipalId = prpim.`imppId`
LEFT OUTER JOIN member prpmm ON prpmm.mmId = prpim.imppMemberId
WHERE impp.`imppExternalRefId` = "?" AND impp.`imppCancelDate` IS NULL AND impp.`imppPlanExpiryDate` >= NOW()
ORDER BY 12 ASC, FIELD(impp.`imppRelationship`, 'P', 'S', 'C') ASC, 1 DESC"""

PRINCIPAL_QUERY = """SELECT mm.`mmFullName`, impp.`imppId`, mm.`mmId`,impp.`imppEmployeeId`, mm.`mmDOB`, mm.`mmMaritalStatus`, mm.`mmNRIC`, mm.`mmOtherIc`, mm.`mmGender`, impp.`imppRelationship`, impp.`imppExternalRefId`, impp.`imppVIP`,
CASE WHEN impp.`imppPrincipalId` IS NULL THEN impp.`imppId` ELSE impp.`imppPrincipalId` END AS "imppPrincipalId",
CASE WHEN impp.`imppPrincipalId` IS NULL THEN mm.`mmFullName` ELSE prpmm.mmFullName END AS "Principal Name",
CASE WHEN impp.`imppPrincipalId` IS NULL THEN mm.`mmNRIC` ELSE prpmm.mmNRIC END AS "Principal IC",
CASE WHEN impp.`imppPrincipalId` IS NULL THEN mm.`mmOtherIc` ELSE prpmm.mmOtherIc END AS "Principal Other IC",
CASE WHEN impp.`imppPrincipalId` IS NULL THEN impp.`imppExternalRefId` ELSE prpim.imppExternalRefId END AS "Principal Ext Ref",
CASE WHEN impp.`imppPrincipalId` IS NULL THEN mm.`mmGender` ELSE prpmm.`mmGender` END AS "Principal Gender",
impp.`imppFirstJoinDate`, impp.`imppPlanAttachDate`, impp.`imppPlanExpiryDate`, impp.`imppCancelDate`, pol.`inptPolPolicyNum`, pol.`inptPolOwner`, polplan.ippInptpolicyId, `inptPlanExtCode`,`inptPlanName`, polplan.ippInptPlanId, impp.`imppInptPolicyPlanId`,impp.`imppCreatedDate`, impp.`imppSpecialConditions`, impp.`imppAddress4`
FROM member mm INNER JOIN `inpt_member_policyplan` impp ON impp.`imppMemberId` = mm.mmid
INNER JOIN `inpt_policy_plan` polplan ON polplan.`ippId` = impp.`imppInptPolicyPlanId`
INNER JOIN inpt_plan plan ON plan.`inptPlanId` = polplan.`ippInptPlanId`
INNER JOIN `inpt_policy` pol ON pol.`inptPolId` = polplan.`ippInptPolicyId`
LEFT OUTER JOIN `inpt_member_policyplan` prpim ON impp.imppPrincipalId = prpim.`imppId`
LEFT OUTER JOIN member prpmm ON prpmm.mmId = prpim.imppMemberId
WHERE impp.`imppRelationship` = 'P' AND pol.`inptPolPolicyNum` = "?" AND impp.imppExternalRefId = '?' AND impp.`imppPlanExpiryDate` = '?'"""

OUT_PATHS = {
}
