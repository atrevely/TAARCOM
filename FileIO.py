import os
import pandas as pd
import numpy as np
import datetime
import shutil
import logging
import GenerateMasterUtils as utils
from xlrd import XLRDError
from RCExcelTools import form_date, save_error, table_format

logger = logging.getLogger(__name__)


def load_salespeople_info():
    """Read in Salespeople Info. Return empty series if not found or if there's an error."""
    sales_info = pd.Series([])
    location = utils.DIRECTORIES.get('COMM_LOOKUPS_DIR')
    try:
        sales_info = pd.read_excel(os.path.join(location, 'Salespeople Info.xlsx'), sheet_name='Info')
        # Make sure the required columns are present.
        cols = ['Salesperson', 'Sales Initials', 'Sales Percentage', 'Territory Cities', 'QQ Split']
        missing_cols = [i for i in cols if i not in list(sales_info)]
        if missing_cols:
            logger.error(f'The following columns were not found in Salespeople Info: {', '.join(missing_cols)}'
                         '\nPlease check for these columns and try again.')
            sales_info = pd.Series([])
        for col in ['Salesperson', 'Territory Cities', 'Sales Initials']:
            sales_info[col] = sales_info[col].fillna('')
        for col in ['Sales Percentage', 'QQ Split']:
            sales_info[col] = sales_info[col].fillna(0)
    except FileNotFoundError:
        logger.error('No Salespeople Info file found!\n'
                     f'Please make sure Salespeople Info.xlsx is in the following directory: {location}')
    except XLRDError:
        logger.error('Error reading sheet name for Salespeople Info.xlsx! '
                     'Please make sure the main tab is named Info.')
    return sales_info


def load_principal_info():
    """Load the Principal Info file."""
    principal_info = pd.Series([])
    location = utils.DIRECTORIES.get('COMM_LOOKUPS_DIR')
    try:
        principal_info = pd.read_excel(os.path.join(location, 'principalList.xlsx'), sheet_name='Principals')
    except FileNotFoundError:
        logger.error('No Principal List file found! '
                     f'Please make sure Principal Info.xlsx is in the following directory: {location}')
    except XLRDError:
        logger.error('Error reading sheet name for principalList.xlsx! '
                     'Please make sure the main tab is named Principals.')
    return principal_info


def load_com_master():
    """Load and prepare the Commissions Master file. Return empty series if not found."""
    com_mast, master_files = pd.Series([]), pd.Series([])
    location = utils.DIRECTORIES.get('COMM_WORKING_DIR')
    file_path = os.path.join(location, 'Commissions Master.xlsx')
    today = datetime.datetime.today().strftime('%m-%d-%Y')
    run_com_backup = file_path.replace('.xlsx', f'_BACKUP_{str(today)}.xlsx')
    logger.info(f'Saving Commissions Master backup as: {run_com_backup}')
    shutil.copy(file_path, run_com_backup)

    try:
        com_mast = pd.read_excel(file_path, sheet_name='Master', dtype=str)
        master_files = pd.read_excel(file_path, sheet_name='Files Processed').fillna('')

        # Force numerical columns to be numeric.
        for col in utils.NUMERICAL_COLUMNS:
            try:
                com_mast[col] = pd.to_numeric(com_mast[col], errors='coerce').fillna(0)
            except KeyError:
                pass

        # Convert individual numbers to numeric in rest of columns.
        # Invoice/part numbers sometimes have leading zeros we'd like to keep, and
        # the INF gets read in as infinity, so remove these.
        mixed_cols = [col for col in list(com_mast) if col not in utils.NUMERICAL_COLUMNS
                      and col not in ['Invoice Number', 'Part Number', 'Principal']]
        for col in mixed_cols:
            com_mast[col] = pd.to_numeric(com_mast[col], errors='ignore')

        # Now remove the nans.
        com_mast.replace(to_replace=['nan', np.nan], value='', inplace=True)

        # Make sure all the dates are formatted correctly.
        for col in ['Invoice Date', 'Paid Date', 'Sales Report Date']:
            com_mast[col] = com_mast[col].map(lambda x: form_date(x))

        # Make sure that the CM Splits aren't blank or zero.
        com_mast['CM Split'] = com_mast['CM Split'].replace(['', '0', 0], 20)
        for col in ['CM Sales', 'Design Sales', 'Principal']:
            com_mast[col] = com_mast[col].map(lambda x: x.strip().upper())
    except FileNotFoundError:
        logger.error('No Commissions Master file found!')
    except XLRDError:
        logger.error('Commissions Master tab names incorrect! '
                     'Make sure the tabs are named Master and Files Processed.')
    return com_mast, master_files


def load_run_com(file_path):
    """Load and prepare the Running Commissions file. Return empty series if not found."""
    running_com, files_processed = pd.Series([]), pd.Series([])
    try:
        running_com = pd.read_excel(file_path, sheet_name='Master', dtype=str)
        files_processed = pd.read_excel(file_path, sheet_name='Files Processed').fillna('')

        for col in utils.NUMERICAL_COLUMNS:
            try:
                running_com[col] = pd.to_numeric(running_com[col], errors='coerce').fillna(0)
            except KeyError:
                pass

        # Convert individual numbers to numeric in rest of columns.
        mixed_cols = [col for col in list(running_com) if col not in utils.NUMERICAL_COLUMNS
                      and col not in ['Invoice Number', 'Part Number', 'Principal']]

        for col in mixed_cols:
            try:
                running_com[col] = pd.to_numeric(running_com[col], errors='ignore')
            except KeyError:
                pass

        # Now remove the nans.
        running_com.replace(to_replace=['nan', np.nan], value='', inplace=True)
        # Make sure all the dates are formatted correctly.
        running_com['Invoice Date'] = running_com['Invoice Date'].map(lambda x: form_date(x))
        # Make sure that the CM Splits aren't blank or zero.
        running_com['CM Split'] = running_com['CM Split'].replace(['', '0', 0], 20)
        for col in ['CM Sales', 'Design Sales', 'Principal']:
            running_com[col] = running_com[col].map(lambda x: x.strip().upper())
    except FileNotFoundError:
        logger.error('No Running Commissions file found!')
    except XLRDError:
        logger.error('Running Commissions tab names incorrect! '
                     'Make sure the tabs are named Master and Files Processed.')
    return running_com, files_processed


# TODO: This can probably be combined with load_run_com.
def load_entries_need_fixing(file_dir):
    """Load and prepare the Entries Need Fixing file."""
    entries_need_fixing = None
    try:
        entries_need_fixing = pd.read_excel(file_dir, sheet_name='Data', dtype=str)
        # Convert entries to proper types, like above.
        for col in utils.NUMERICAL_COLUMNS:
            try:
                entries_need_fixing[col] = pd.to_numeric(entries_need_fixing[col], errors='coerce').fillna('')
            except KeyError:
                logger.error(f'The following column was not found in ENF: {col}. '
                             f'Please check the column names and try again')
                return None
        mixed_cols = [col for col in list(entries_need_fixing) if col not in utils.NUMERICAL_COLUMNS
                      and col not in ['Invoice Number', 'Part Number', 'Principal']]
        for col in mixed_cols:
            try:
                entries_need_fixing[col] = entries_need_fixing[col].map(lambda x: pd.to_numeric(x, errors='ignore'))
            except KeyError:
                logger.error(f'The following column was not found in ENF: {col}. '
                             'Please check the column names and try again.')
                return None
        # Now remove the nans.
        entries_need_fixing.replace(to_replace=['nan', np.nan], value='', inplace=True)
        entries_need_fixing['Invoice Date'] = entries_need_fixing['Invoice Date'].map(lambda x: form_date(x))
        # Make sure that the CM Splits aren't blank or zero.
        entries_need_fixing['CM Split'] = entries_need_fixing['CM Split'].replace(['', '0', 0], 20)
    except FileNotFoundError:
        logger.error('No matching Entries Need Fixing file found for this Running Commissions file!')
    except XLRDError:
        logger.error('No sheet named Data found in Entries Need Fixing!')
    return entries_need_fixing


def load_acct_list():
    """Load and prepare the Account List file."""
    acct_list = pd.Series([])
    location = utils.DIRECTORIES.get('COMM_LOOKUPS_DIR')
    try:
        acct_list = pd.read_excel(os.path.join(location, 'Master Account List.xlsx'), sheet_name='Allacct').fillna('')
        # Make sure the required columns are present.
        cols = ['SLS', 'ProperName']
        missing_cols = [i for i in cols if i not in list(acct_list)]
        if missing_cols:
            logger.error(f'The following columns were not found in the Account List: {', '.join(missing_cols)}'
                         '\nPlease check for these column names (case-sensitive) and try again.')
            acct_list = pd.Series([])
    except FileNotFoundError:
        logger.error(f'No Account List file found! Please make sure it is location in {location}')
    except XLRDError:
        logger.error('Account List tab names incorrect! Make sure the main tab is named Allacct.')
    return acct_list


def load_lookup_master():
    """Load and prepare the Lookup Master."""
    master_lookup = pd.Series([])
    location = utils.DIRECTORIES.get('COMM_LOOKUPS_DIR')
    try:
        master_lookup = pd.read_excel(os.path.join(location, 'Lookup Master - Current.xlsx')).fillna('')
        # Make sure the required columns are present.
        lookup_cols = ['CM Sales', 'Design Sales', 'CM Split', 'Reported Customer', 'CM', 'Part Number',
                       'T-Name', 'T-End Cust', 'Last Used', 'Principal', 'City', 'Date Added']
        missing_cols = [i for i in lookup_cols if i not in list(master_lookup)]
        if missing_cols:
            logger.error(f'The following columns were not found in the Lookup Master: {', '.join(missing_cols)}'
                         '\nPlease check for these column names and try again.')
            return pd.Series([])
        # Set the CM Split to an int.
        master_lookup['CM Split'] = master_lookup['CM Split'].map(lambda x: int(x) if isinstance(x, float) else x)
    except FileNotFoundError:
        logger.error(f'No Lookup Master found! Please make sure Lookup Master - Current.xlsx is in {location}')
    except XLRDError:
        logger.error('Error loading the Lookup Master.')
    return master_lookup


def load_root_customer_mappings():
    """Load and prepare the root customer mappings file."""
    customer_mappings = pd.Series([])
    location = utils.DIRECTORIES.get('COMM_LOOKUPS_DIR')
    try:
        customer_mappings = pd.read_excel(os.path.join(location, 'rootCustomerMappings.xlsx'),
                                          sheet_name='Sales Lookup').fillna('')
        # Check the column names.
        map_cols = ['Root Customer', 'Salesperson']
        missing_cols = [i for i in map_cols if i not in list(customer_mappings)]
        if missing_cols:
            logger.error(f'The following columns were not detected in rootCustomerMappings.xlsx:'
                         f'\n{', '.join(map(str, missing_cols))}')
            customer_mappings = pd.Series([])
    except FileNotFoundError:
        logger.error('No Root Customer Mappings file found! '
                     'Please make sure rootCustomerMappings.xlsx is in the directory.')
    except XLRDError:
        logger.error('Error reading sheet names.')
    return customer_mappings


def load_digikey_master():
    """Load and prepare the digikey insights master file."""
    digikey_master = pd.Series([])
    files_processed = pd.Series([])
    location = utils.DIRECTORIES.get('DIGIKEY_DIR')
    try:
        digikey_master = pd.read_excel(os.path.join(location, 'Digikey Insight Master.xlsx'),
                                       sheet_name='Master').fillna('')
        files_processed = pd.read_excel(os.path.join(location, 'Digikey Insight Master.xlsx'),
                                        sheet_name='Files Processed').fillna('')
    except FileNotFoundError:
        logger.error('No Digikey Insight Master file found! '
                     'Please make sure Digikey Insight Master is in the directory.')
    except XLRDError:
        logger.error('Error reading sheet names.')
    return digikey_master, files_processed


def load_distributor_map():
    distributor_map = pd.Series([])
    location = utils.DIRECTORIES.get('COMM_LOOKUPS_DIR')
    # Read in the distributor map. Terminate if not found or if errors in file.
    if os.path.exists(os.path.join(location, 'distributorLookup.xlsx')):
        try:
            distributor_map = pd.read_excel(os.path.join(location, 'distributorLookup.xlsx'), sheet_name='Distributors')
        except XLRDError:
            logger.error('Error reading sheet name for distributorLookup.xlsx! '
                         'Please make sure the main tab is named Distributors.\n*Program terminated*')
            return
        # Check the column names.
        disty_map_cols = ['Corrected Dist', 'Search Abbreviation']
        missing_cols = [i for i in disty_map_cols if i not in list(distributor_map)]
        if missing_cols:
            logger.error(f'The following columns were not detected in distributorLookup.xlsx:'
                         f'\n{', '.join(map(str, missing_cols))}\n*Program terminated*')
            return
    else:
        logger.error('No distributor lookup file found! '
                     'Please make sure distributorLookup.xlsx is in the directory.\n*Program terminated*')
    return distributor_map


def save_excel_file(filename, tab_data, tab_names):
    """Save a file as an Excel spreadsheet."""
    if save_error(filename):
        logger.error(f'The following file is currently open in Excel: {filename}'
                     f'\nPlease close the file and try again.')
        return None
    if not isinstance(tab_data, list):
        tab_data = [tab_data]
    if not isinstance(tab_names, list):
        tab_names = [tab_names]
    assert len(tab_data) == len(tab_names), logger.error(f'Mismatch in size of tab data and tab names in {filename}.')
    # Add each tab to the document.
    with pd.ExcelWriter(filename, datetime_format='mm/dd/yyyy') as writer:
        for data, sheet_name in zip(tab_data, tab_names):
            data.to_excel(writer, sheet_name=sheet_name, index=False)
            table_format(sheet_data=data, sheet_name=sheet_name, workbook=writer)
