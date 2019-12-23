import os
import pandas as pd
import numpy as np
from xlrd import XLRDError
from RCExcelTools import form_date

# Set the numerical columns.
num_cols = ['Quantity', 'Ext. Cost', 'Invoiced Dollars',
            'Paid-On Revenue', 'Actual Comm Paid', 'Unit Cost',
            'Unit Price', 'CM Split', 'Year', 'Sales Commission',
            'Split Percentage', 'Commission Rate',
            'Gross Rev Reduction', 'Shared Rev Tier Rate']


def load_salespeople_info(file_dir):
    """Read in Salespeople Info. Return empty series if not found."""
    sales_info = pd.Series([])
    if os.path.exists(file_dir + '\\Salespeople Info.xlsx'):
        try:
            sales_info = pd.read_excel(file_dir + '\\Salespeople Info.xlsx',
                                       'Info')
        except XLRDError:
            print('---\n'
                  'Error reading sheet name for Salespeople Info.xlsx!\n'
                  'Please make sure the main tab is named Info.\n')
    else:
        print('---\n'
              'No Salespeople Info file found!\n'
              'Please make sure Salespeople Info.xlsx is in the directory.\n')
    return sales_info


def load_com_master(file_dir):
    """Load and prepare the Commissions Master file. Return empty series if not found."""
    com_mast = pd.Series([])
    master_files = pd.Series([])
    try:
        com_mast = pd.read_excel(file_dir + '\\Commissions Master.xlsx',
                                 'Master Data', dtype=str)
        master_files = pd.read_excel(file_dir + '\\Commissions Master.xlsx',
                                     'Files Processed').fillna('')
    except FileNotFoundError:
        print('No Commissions Master file found!\n'
              '*Program Terminated*')
        return com_mast, master_files
    except XLRDError:
        print('Commissions Master tab names incorrect!\n'
              'Make sure the tabs are named Master Data and Files Processed.\n')
        return com_mast, master_files
    # Force numerical columns to be numeric.
    for col in num_cols:
        try:
            com_mast[col] = pd.to_numeric(com_mast[col],
                                          errors='coerce').fillna(0)
        except KeyError:
            pass
    # Convert individual numbers to numeric in rest of columns.
    mixed_cols = [col for col in list(com_mast) if col not in num_cols]
    # Invoice/part numbers sometimes have leading zeros we'd like to keep.
    mixed_cols.remove('Invoice Number')
    mixed_cols.remove('Part Number')
    # The INF gets read in as infinity, so skip the principal column.
    mixed_cols.remove('Principal')
    for col in mixed_cols:
        com_mast[col] = com_mast[col].map(
            lambda x: pd.to_numeric(x, errors='ignore'))
    # Now remove the nans.
    com_mast.replace(['nan', np.nan], '', inplace=True)
    # Make sure all the dates are formatted correctly.
    for col in ['Invoice Date', 'Paid Date', 'Sales Report Date']:
        com_mast[col] = com_mast[col].map(lambda x: form_date(x))
    # Make sure that the CM Splits aren't blank or zero.
    com_mast['CM Split'] = com_mast['CM Split'].replace(['', '0', 0], 20)
    return com_mast, master_files


def load_run_com(file_path):
    """Load and prepare the Running Commissions file.
    Return empty series if not found.
    """
    running_com = pd.Series([])
    files_processed = pd.Series([])
    try:
        running_com = pd.read_excel(file_path, 'Master', dtype=str)
        files_processed = pd.read_excel(file_path, 'Files Processed').fillna('')
    except FileNotFoundError:
        print('No Running Commissions file found!\n')
        return running_com, files_processed
    except XLRDError:
        print('Running Commissions tab names incorrect!\n'
              'Make sure the tabs are named Master and Files Processed.\n')
        return running_com, files_processed
    for col in num_cols:
        try:
            running_com[col] = pd.to_numeric(running_com[col],
                                             errors='coerce').fillna(0)
        except KeyError:
            pass
    # Convert individual numbers to numeric in rest of columns.
    mixed_cols = [col for col in list(running_com) if col not in num_cols]
    # Invoice/part numbers sometimes has leading zeros we'd like to keep.
    mixed_cols.remove('Invoice Number')
    mixed_cols.remove('Part Number')
    # The INF gets read in as infinity, so skip the principal column.
    mixed_cols.remove('Principal')
    for col in mixed_cols:
        try:
            running_com[col] = running_com[col].map(
                lambda x: pd.to_numeric(x, errors='ignore'))
        except KeyError:
            pass
    # Now remove the nans.
    running_com.replace(['nan', np.nan], '', inplace=True)
    # Make sure all the dates are formatted correctly.
    running_com['Invoice Date'] = running_com['Invoice Date'].map(
        lambda x: form_date(x))
    # Make sure that the CM Splits aren't blank or zero.
    running_com['CM Split'] = running_com['CM Split'].replace(['', '0', 0], 20)
    # Strip any extra spaces that made their way into salespeople columns.
    for col in ['CM Sales', 'Design Sales']:
        running_com[col] = running_com[col].map(lambda x: x.strip())
    return running_com, files_processed


def load_acct_list(file_dir):
    """Load and prepare the Account List file."""
    acctList = pd.Series([])
    try:
        acctList = pd.read_excel(file_dir + '\\Master Account List.xlsx',
                                 'Allacct')
    except FileNotFoundError:
        print('No Account List file found!\n')
    except XLRDError:
        print('Account List tab names incorrect!\n'
              'Make sure the main tab is named Allacct.\n')
    return acctList


def load_lookup_master(file_dir):
    """Load and prepare the Lookup master."""
    master_lookup = pd.Series([])
    if os.path.exists(file_dir + '\\Lookup Master - Current.xlsx'):
        master_lookup = pd.read_excel(file_dir + '\\Lookup Master - '
                                      'Current.xlsx').fillna('')
        # Check the column names.
        lookup_cols = ['CM Sales', 'Design Sales', 'CM Split',
                       'Reported Customer', 'CM', 'Part Number', 'T-Name',
                       'T-End Cust', 'Last Used', 'Principal', 'City',
                       'Date Added']
        miss_cols = [i for i in lookup_cols if i not in list(master_lookup)]
        if miss_cols:
            print('The following columns were not detected in '
                  'Lookup Master.xlsx:\n%s' %
                  ', '.join(map(str, missCols)))
    else:
        print('---\n'
              'No Lookup Master found!\n'
              'Please make sure Lookup Master - Current.xlsx is '
              'in the directory.\n')
    return master_lookup
