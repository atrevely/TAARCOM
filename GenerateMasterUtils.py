import os
import pandas as pd
import datetime
import logging
from dateutil.parser import parse

logger = logging.getLogger(__name__)

# Define the directories where supporting files are located.
TAARCOM_DIRECTORIES = {'COMM_LOOKUPS_DIR': 'Z:\\Commissions Lookup', 'COMM_WORKING_DIR': 'Z:\\MK Working Commissions',
                       'COMM_REPORTS_DIR': 'Z:\\MK Working Commissions\\Reports', 'DIGIKEY_DIR': 'W:\\'}
# If any directories aren't found, then replace them with the current working directory.
DIRECTORIES = {i: j if os.path.exists(j) else os.getcwd() for i, j in TAARCOM_DIRECTORIES.items()}

# Columns defined as containing numerical data.
NUMERICAL_COLUMNS = ['Quantity', 'Ext. Cost', 'Invoiced Dollars', 'Paid-On Revenue', 'Actual Comm Paid',
                     'Unit Cost', 'Unit Price', 'CM Split', 'Year', 'Sales Commission', 'Split Percentage',
                     'Commission Rate', 'Gross Rev Reduction', 'Shared Rev Tier Rate']


def get_column_names(field_mappings):
    """Generate the commission file column names in the correct order."""
    # Grab lookup table data names.
    column_names = list(field_mappings)

    # Add in non-lookup'd data names.
    column_names[0:0] = ['CM Sales', 'Design Sales']
    column_names[3:3] = ['T-Name', 'CM', 'T-End Cust']
    column_names[7:7] = ['Principal', 'Distributor']
    column_names[18:18] = ['CM Sales Comm', 'Design Sales Comm', 'Sales Commission']
    column_names[22:22] = ['Quarter Shipped', 'Month', 'Year']
    column_names.extend(['CM Split', 'Paid Date', 'From File', 'Sales Report Date', 'T-Notes', 'Unique ID'])

    # Make sure that column names are unique.
    unique_names = []
    for col_name in column_names:
        if col_name not in unique_names:
            unique_names.append(col_name)

    return unique_names


def filter_duplicate_files(filepaths, files_processed):
    """Check to ensure that no duplicate files were provided."""
    filenames = [os.path.basename(val) for val in filepaths]
    duplicates = list(set(filenames).intersection(files_processed['Filename']))
    filenames = [val for val in filenames if val not in duplicates]
    if duplicates:
        # Let us know we found duplicates and removed them.
        logger.warning(f'The following files are already in Running Commissions: {', '.join(map(str, duplicates))}'
                       '\nDuplicate files were removed from processing.')
    return filenames


def check_for_date_errors(date):
    # Check if the date is read in as a float/int, and convert to string.
    if isinstance(date, (float, int)):
        date = str(int(date))
    # Check if Pandas read it in as a Timestamp object.
    # If so, turn it back into a string (a bit roundabout, oh well).
    elif isinstance(date, (pd.Timestamp, datetime.datetime)):
        date = str(date)
    try:
        parse(date)
    except (ValueError, TypeError):
        # The date isn't recognized by the parser.
        return True
    except KeyError:
        logger.error('There is no Invoice Date column in Running Commissions! '
                     'Please check to make sure an Invoice Date column exists. '
                     'Note: Spelling, whitespace, and capitalization matter.')
        return True
    return False


def format_pct_numeric_cols(dataframe):
    # Convert applicable columns to numeric.
    numeric_cols = ['Quantity', 'Ext. Cost', 'Invoiced Dollars', 'Paid-On Revenue', 'Actual Comm Paid',
                    'Unit Cost', 'Unit Price']
    for col in numeric_cols:
        try:
            dataframe[col] = pd.to_numeric(dataframe[col], errors='coerce').fillna('')
        except KeyError:
            pass

    # Fix Commission Rate if it got read in as a decimal.
    pct_cols = ['Commission Rate', 'Split Percentage', 'Gross Rev Reduction', 'Shared Rev Tier Rate']
    for pct_col in pct_cols:
        try:
            # Remove '%' sign if present.
            col = dataframe[pct_col].astype(str).map(lambda x: x.strip('%'))
            # Convert to numeric.
            col = pd.to_numeric(col, errors='coerce')
            # Identify and correct which entries are not decimal.
            col[col > 1] /= 100
            dataframe[pct_col] = col.fillna(0)
        except (KeyError, TypeError):
            pass

    return dataframe
