import os
import pandas as pd
import numpy as np
import datetime
import shutil
from xlrd import XLRDError
from RCExcelTools import form_date

# Define the directories where supporting files are located.
DIRECTORIES = {'COMM_LOOKUPS_DIR': 'Z:\\Commissions Lookup', 'COMM_WORKING_DIR': 'Z:\\MK Working Commissions',
               'COMM_REPORTS_DIR': 'Z:\\MK Working Commissions\\Reports', 'DIGIKEY_DIR': 'W:\\'}
# If any directories aren't found, then replace them with the current working directory.
DIRECTORIES = {i: j if os.path.exists(j) else os.getcwd() for i, j in DIRECTORIES.items()}


# Set the numerical columns.
NUM_COLS = ['Quantity', 'Ext. Cost', 'Invoiced Dollars', 'Paid-On Revenue', 'Actual Comm Paid',
            'Unit Cost', 'Unit Price', 'CM Split', 'Year', 'Sales Commission', 'Split Percentage',
            'Commission Rate', 'Gross Rev Reduction', 'Shared Rev Tier Rate']


def load_salespeople_info():
    """Read in Salespeople Info. Return empty series if not found or if there's an error."""
    sales_info = pd.Series([])
    location = DIRECTORIES.get('COMM_LOOKUPS_DIR')
    try:
        sales_info = pd.read_excel(os.path.join(location, 'Salespeople Info.xlsx'), 'Info')
        # Make sure the required columns are present.
        cols = ['Salesperson', 'Sales Initials', 'Sales Percentage', 'Territory Cities', 'QQ Split']
        missing_cols = [i for i in cols if i not in list(sales_info)]
        if missing_cols:
            print('---\nThe following columns were not found in Salespeople Info: '
                  + ', '.join(missing_cols) + '\nPlease check for these columns and try again.')
            sales_info = pd.Series([])
        for col in ['Salesperson', 'Territory Cities', 'Sales Initials']:
            sales_info[col] = sales_info[col].fillna('')
        for col in ['Sales Percentage', 'QQ Split']:
            sales_info[col] = sales_info[col].fillna(0)
    except FileNotFoundError:
        print('---\nNo Salespeople Info file found!\n'
              'Please make sure Salespeople Info.xlsx is in the following directory:\n' + location)
    except XLRDError:
        print('---\nError reading sheet name for Salespeople Info.xlsx!\n'
              'Please make sure the main tab is named Info.')
    return sales_info


def load_principal_info():
    """Load the Principal Info file."""
    princ_info = pd.Series([])
    location = DIRECTORIES.get('COMM_LOOKUPS_DIR')
    try:
        princ_info = pd.read_excel(os.path.join(location, 'principalList.xlsx'), 'Principals')
    except FileNotFoundError:
        print('---\nNo Principal List file found!\n'
              'Please make sure Principal Info.xlsx is in the following directory:\n' + location)
    except XLRDError:
        print('---\nError reading sheet name for principalList.xlsx!\n'
              'Please make sure the main tab is named Principals.')
    return princ_info


def load_com_master():
    """Load and prepare the Commissions Master file. Return empty series if not found."""
    com_mast = pd.Series([])
    master_files = pd.Series([])
    location = DIRECTORIES.get('COMM_WORKING_DIR')
    file_path = os.path.join(location, 'Commissions Master.xlsx')
    today = datetime.datetime.today().strftime('%m-%d-%Y')
    run_com_backup = file_path.replace('.xlsx', '_BACKUP_' + str(today) + '.xlsx')
    print('Saving backup as: %s' % run_com_backup)
    shutil.copy(file_path, run_com_backup)
    try:
        com_mast = pd.read_excel(file_path, 'Master', dtype=str)
        master_files = pd.read_excel(file_path, 'Files Processed').fillna('')
        # Force numerical columns to be numeric.
        for col in NUM_COLS:
            try:
                com_mast[col] = pd.to_numeric(com_mast[col], errors='coerce').fillna(0)
            except KeyError:
                pass
        # Convert individual numbers to numeric in rest of columns.
        # Invoice/part numbers sometimes have leading zeros we'd like to keep, and
        # the INF gets read in as infinity, so remove these.
        mixed_cols = [col for col in list(com_mast) if col not in NUM_COLS
                      and col not in ['Invoice Number', 'Part Number', 'Principal']]
        for col in mixed_cols:
            com_mast[col] = pd.to_numeric(com_mast[col], errors='ignore')
        # Now remove the nans.
        com_mast.replace(['nan', np.nan], '', inplace=True)
        # Make sure all the dates are formatted correctly.
        for col in ['Invoice Date', 'Paid Date', 'Sales Report Date']:
            com_mast[col] = com_mast[col].map(lambda x: form_date(x))
        # Make sure that the CM Splits aren't blank or zero.
        com_mast['CM Split'] = com_mast['CM Split'].replace(['', '0', 0], 20)
        for col in ['CM Sales', 'Design Sales', 'Principal']:
            com_mast[col] = com_mast[col].map(lambda x: x.strip().upper())
    except FileNotFoundError:
        print('---\nNo Commissions Master file found!')
    except XLRDError:
        print('---\nCommissions Master tab names incorrect!\n'
              'Make sure the tabs are named Master and Files Processed.')
    return com_mast, master_files


def load_run_com(file_path):
    """Load and prepare the Running Commissions file. Return empty series if not found."""
    running_com, files_processed = pd.Series([]), pd.Series([])
    try:
        running_com = pd.read_excel(file_path, 'Master', dtype=str)
        files_processed = pd.read_excel(file_path, 'Files Processed').fillna('')
        for col in NUM_COLS:
            try:
                running_com[col] = pd.to_numeric(running_com[col], errors='coerce').fillna(0)
            except KeyError:
                pass
        # Convert individual numbers to numeric in rest of columns.
        mixed_cols = [col for col in list(running_com) if col not in NUM_COLS
                      and col not in ['Invoice Number', 'Part Number', 'Principal']]
        for col in mixed_cols:
            try:
                running_com[col] = pd.to_numeric(running_com[col], errors='ignore')
            except KeyError:
                pass
        # Now remove the nans.
        running_com.replace(['nan', np.nan], '', inplace=True)
        # Make sure all the dates are formatted correctly.
        running_com['Invoice Date'] = running_com['Invoice Date'].map(lambda x: form_date(x))
        # Make sure that the CM Splits aren't blank or zero.
        running_com['CM Split'] = running_com['CM Split'].replace(['', '0', 0], 20)
        for col in ['CM Sales', 'Design Sales', 'Principal']:
            running_com[col] = running_com[col].map(lambda x: x.strip().upper())
    except FileNotFoundError:
        print('---\nNo Running Commissions file found!')
    except XLRDError:
        print('---\nRunning Commissions tab names incorrect!\n'
              'Make sure the tabs are named Master and Files Processed.')
    return running_com, files_processed


# TODO: This can probably be combined with load_run_com.
def load_entries_need_fixing(file_dir):
    """Load and prepare the Entries Need Fixing file."""
    entries_need_fixing = None
    try:
        entries_need_fixing = pd.read_excel(file_dir, 'Data', dtype=str)
        # Convert entries to proper types, like above.
        for col in NUM_COLS:
            try:
                entries_need_fixing[col] = pd.to_numeric(entries_need_fixing[col], errors='coerce').fillna('')
            except KeyError:
                print('The following column was not found in ENF: ' + col
                      + '\nPlease check the column names and try again')
                return None
        mixed_cols = [col for col in list(entries_need_fixing) if col not in NUM_COLS
                      and col not in ['Invoice Number', 'Part Number', 'Principal']]
        for col in mixed_cols:
            try:
                entries_need_fixing[col] = entries_need_fixing[col].map(lambda x: pd.to_numeric(x, errors='ignore'))
            except KeyError:
                print('The following column was not found in ENF: ' + col
                      + '\nPlease check the column names and try again.')
                return None
        # Now remove the nans.
        entries_need_fixing.replace(['nan', np.nan], '', inplace=True)
        entries_need_fixing['Invoice Date'] = entries_need_fixing['Invoice Date'].map(lambda x: form_date(x))
        # Make sure that the CM Splits aren't blank or zero.
        entries_need_fixing['CM Split'] = entries_need_fixing['CM Split'].replace(['', '0', 0], 20)
    except FileNotFoundError:
        print('No matching Entries Need Fixing file found for this Running Commissions file!')
    except XLRDError:
        print('No sheet named Data found in Entries Need Fixing!')
    return entries_need_fixing


def load_acct_list():
    """Load and prepare the Account List file."""
    acct_list = pd.Series([])
    location = DIRECTORIES.get('COMM_LOOKUPS_DIR')
    try:
        acct_list = pd.read_excel(os.path.join(location, 'Master Account List.xlsx'), 'Allacct').fillna('')
        # Make sure the required columns are present.
        cols = ['SLS', 'ProperName']
        missing_cols = [i for i in cols if i not in list(acct_list)]
        if missing_cols:
            print('---\nThe following columns were not found in the Account List: '
                  + ', '.join(missing_cols) + '\nPlease check for these column '
                  'names (case-sensitive) and try again.')
            acct_list = pd.Series([])
    except FileNotFoundError:
        print('---\nNo Account List file found! Please make sure it is location in ' + location)
    except XLRDError:
        print('---\nAccount List tab names incorrect!\nMake sure the main tab is named Allacct.')
    return acct_list


def load_lookup_master():
    """Load and prepare the Lookup Master."""
    master_lookup = pd.Series([])
    location = DIRECTORIES.get('COMM_LOOKUPS_DIR')
    try:
        master_lookup = pd.read_excel(os.path.join(location, 'Lookup Master - Current.xlsx')).fillna('')
        # Make sure the required columns are present.
        lookup_cols = ['CM Sales', 'Design Sales', 'CM Split', 'Reported Customer', 'CM', 'Part Number',
                       'T-Name', 'T-End Cust', 'Last Used', 'Principal', 'City', 'Date Added']
        missing_cols = [i for i in lookup_cols if i not in list(master_lookup)]
        if missing_cols:
            print('---\nThe following columns were not found in the Lookup Master: '
                  + ', '.join(missing_cols) + '\nPlease check for these column names and try again.')
            return pd.Series([])
        # Set the CM Split to an int.
        master_lookup['CM Split'] = master_lookup['CM Split'].map(lambda x: int(x) if isinstance(x, float) else x)
    except FileNotFoundError:
        print('---\nNo Lookup Master found!\nPlease make sure Lookup Master - Current.xlsx is in ' + location)
    except XLRDError:
        print('Error reading sheet names.')
    return master_lookup


def load_root_customer_mappings():
    """Load and prepare the root customer mappings file."""
    customer_mappings = pd.Series([])
    location = DIRECTORIES.get('COMM_LOOKUPS_DIR')
    try:
        customer_mappings = pd.read_excel(os.path.join(location, 'rootCustomerMappings.xlsx'),
                                          'Sales Lookup').fillna('')
        # Check the column names.
        map_cols = ['Root Customer', 'Salesperson']
        missing_cols = [i for i in map_cols if i not in list(customer_mappings)]
        if missing_cols:
            print('The following columns were not detected in rootCustomerMappings.xlsx:\n%s' %
                  ', '.join(map(str, missing_cols)))
            customer_mappings = pd.Series([])
    except FileNotFoundError:
        print('---\nNo Root Customer Mappings file found!\n'
              'Please make sure rootCustomerMappings.xlsx is in the directory.\n')
    except XLRDError:
        print('Error reading sheet names.')
    return customer_mappings


def load_digikey_master():
    """Load and prepare the digikey insights master file."""
    digikey_master = pd.Series([])
    files_processed = pd.Series([])
    location = DIRECTORIES.get('DIGIKEY_DIR')
    try:
        digikey_master = pd.read_excel(os.path.join(location, 'Digikey Insight Master.xlsx'),
                                       'Master').fillna('')
        files_processed = pd.read_excel(os.path.join(location, 'Digikey Insight Master.xlsx'),
                                        'Files Processed').fillna('')
    except FileNotFoundError:
        print('---\nNo Digikey Insight Master file found!\n'
              'Please make sure Digikey Insight Master is in the directory.\n')
    except XLRDError:
        print('Error reading sheet names.')
    return digikey_master, files_processed
