import pandas as pd
import math
from dateutil.parser import parse
from RCExcelTools import tableFormat, formDate, saveError
from xlrd import XLRDError


# The main function.
def main(runCom):
    """Generates quarterly reports, then marks lines as paid."""
    # Set the directory for the data input/output.
    dataDir = 'Z:/MK Working Commissions/'

    print('Loading the data from Commissions Master...')

    # ---------------------------------------------
    # Load and prepare the Commissions Master file.
    # ---------------------------------------------
    # Load up the current Commissions Master file from the server.
    try:
        comMast = pd.read_excel(dataDir + 'Commissions Master.xlsx', 'Master',
                                dtype=str)
        masterFiles = pd.read_excel(dataDir + 'Commissions Master.xlsx',
                                    'Files Processed').fillna('')
    except FileNotFoundError:
        print('No Commissions Master file found!\n'
              '***')
        return
    except XLRDError:
        print('Commissions Master tab names incorrect!\n'
              'Make sure the tabs are named Master and Files Processed.\n'
              '***')
        return
    # Convert applicable columns to numeric.
    numCols = ['Quantity', 'Ext. Cost', 'Invoiced Dollars', 'Paid-On Revenue',
               'Actual Comm Paid', 'Unit Cost', 'Unit Price', 'CM Split',
               'Year', 'Sales Commission', 'Split Percentage',
               'Commission Rate', 'Gross Rev Reduction',
               'Shared Rev Tier Rate']
    for col in numCols:
        try:
            comMast[col] = pd.to_numeric(comMast[col],
                                         errors='coerce').fillna(0)
        except KeyError:
            pass
    # Convert individual numbers to numeric in rest of columns.
    mixedCols = [col for col in list(comMast) if col not in numCols]
    # Invoice/part numbers sometimes has leading zeros we'd like to keep.
    mixedCols.remove('Invoice Number')
    mixedCols.remove('Part Number')
    # The INF gets read in as infinity, so skip the principal column.
    mixedCols.remove('Principal')
    for col in mixedCols:
        comMast[col] = comMast[col].map(
                lambda x: pd.to_numeric(x, errors='ignore'))
    # Now remove the nans.
    comMast.replace('nan', '', inplace=True)

    # Make sure all the dates are formatted correctly.
    for col in ['Invoice Date', 'Paid Date', 'Sales Report Date']:
        comMast[col] = comMast[col].map(lambda x: formDate(x))
    # Make sure that the CM Splits aren't blank or zero.
    comMast['CM Split'] = comMast['CM Split'].replace(['', '0', 0], 20)

    # -------------------------------------------------
    # Filter down data to most recent finished quarter.
    # -------------------------------------------------
    # Determine the commissions months that are currently in the Master.
    commMonths = comMast['Comm Month'].unique()
    try:
        commMonths = [parse(str(i)) for i in commMonths if i != '']
    except ValueError:
        print('Error parsing dates in Comm Month column of Commissions Master!'
              '\nPlease check that all dates are in standard formatting and '
              'try again.\n***')
        return
    # Grab the most recent month in Commissions Master
    lastMonth = max(commMonths)
    # Find the most recent finished quarter.
    quarter = math.floor(int(lastMonth[-1:])/3)
    print('Most recent quarter detected as Q' + str(quarter)
          + '\n---')
    # Find months in quarter.
    monthsInQtr = [lastMonth[:-1] + str(quarter*3 - i) for i in [0, 1, 2]]
    # Grab the data for the current quarter.
    qtrCom = comMast[comMast['Comm Month'].isin(monthsInQtr)]
    










