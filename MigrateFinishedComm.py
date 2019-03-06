import pandas as pd
import datetime
from dateutil.parser import parse
import os.path
from xlrd import XLRDError


def tableFormat(sheetData, sheetName, wbook):
    """Formats the Excel output as a table with correct column formatting."""
    # Nothing to format, so return.
    if sheetData.shape[0] == 0:
        return
    sheet = wbook.sheets[sheetName]
    # Set default document format.
    docFormat = wbook.book.add_format({'font': 'Calibri',
                                       'font_size': 11})
    # Accounting format ($ XX.XX).
    acctFormat = wbook.book.add_format({'font': 'Calibri',
                                        'font_size': 11,
                                        'num_format': 44})
    # Comma format (XX,XXX).
    commaFormat = wbook.book.add_format({'font': 'Calibri',
                                         'font_size': 11,
                                         'num_format': 3})
    # Percent format, one decimal (XX.X%).
    pctFormat = wbook.book.add_format({'font': 'Calibri',
                                       'font_size': 11,
                                       'num_format': '0.0%'})
    # Date format (YYYY-MM-DD).
    dateFormat = wbook.book.add_format({'font': 'Calibri',
                                        'font_size': 11,
                                        'num_format': 14})
    # Format and fit each column.
    index = 0
    for col in sheetData.columns:
        # Match the correct formatting to each column.
        acctCols = ['Unit Price', 'Paid-On Revenue', 'Actual Comm Paid',
                    'Total NDS', 'Post-Split NDS', 'Cust Revenue YTD',
                    'Ext. Cost', 'Unit Cost', 'Total Commissions',
                    'Sales Commission', 'Invoiced Dollars']
        pctCols = ['Split Percentage', 'Commission Rate',
                   'Gross Rev Reduction', 'Shared Rev Tier Rate']
        coreCols = ['CM Sales', 'Design Sales', 'T-End Cust', 'T-Name',
                    'CM', 'Invoice Date']
        dateCols = ['Invoice Date', 'Paid Date', 'Sales Report Date',
                    'Date Added']
        if col in acctCols:
            formatting = acctFormat
        elif col in pctCols:
            formatting = pctFormat
        elif col in dateCols:
            formatting = dateFormat
        elif col == 'Quantity':
            formatting = commaFormat
        elif col == 'Invoice Number':
            # We're going to do some work in order to format the Invoice
            # Number as a number, yet keep leading zeros.
            for row in sheetData.index:
                invLen = len(sheetData.loc[row, col])
                # Figure out how many places the number goes to.
                numPadding = '0'*invLen
                invNum = pd.to_numeric(sheetData.loc[row, col],
                                       errors='ignore')
                invFormat = wbook.book.add_format({'font': 'Calibri',
                                                   'font_size': 11,
                                                   'num_format': numPadding})
                try:
                    sheet.write_number(row+1, index, invNum, invFormat)
                except TypeError:
                    pass
            # Move to the next column, as we're now done formatting
            # the Invoice Numbers.
            index += 1
            continue
        else:
            formatting = docFormat
        # Set column width and formatting.
        try:
            maxWidth = max(len(str(val)) for val in sheetData[col].values)
        except ValueError:
            maxWidth = 0
        # If column is one that always gets filled in, then keep it expanded.
        if col in coreCols:
            maxWidth = max(maxWidth, len(col), 10)
        # Don't let the columns get too wide.
        maxWidth = min(maxWidth, 50)
        # Extra space for '$'/'%' in accounting/percent format.
        if col in acctCols or col in pctCols:
            maxWidth += 2
        sheet.set_column(index, index, maxWidth+0.8, formatting)
        index += 1


def saveError(*excelFiles):
    """Check Excel files and return True if any file is open."""
    for file in excelFiles:
        try:
            open(file, 'r+')
        except FileNotFoundError:
            pass
        except PermissionError:
            return True
    return False


def formDate(inputDate):
    """Attemps to format a string as a date, otherwise ignores it."""
    try:
        outputDate = parse(str(inputDate)).date()
        return outputDate
    except (ValueError, OverflowError):
        return inputDate


# %% Main function.
def main(filepath, masterCom):
    """Appends a finished monthly Running Commissions file to the Master.

    Arguments:
    filepath -- path for opening Running Commissions (Excel file) to process.
    masterCom -- the Commissions Master file that holds historical data.
    """
    try:
        runCom = pd.read_excel(filepath, 'Master', dtype=str)
    except XLRDError:
        print('Error reading sheet name in Running Commissions file!\n'
              'Please make sure the main tab is named Master.\n'
              '***')
        return
    try:
        filesProcessed = pd.read_excel(filepath, 'Files Processed', dtype=str)
    except XLRDError:
        print('Error reading sheet name for  Running Commissions file!\n'
              'Please make sure the second tab is named Files Processed.\n'
              '***')
        return

    # Read in the Commissions Master. Exit if not found.
    if os.path.exists('Commissions Master.xlsx'):
        masterComm = pd.read_excel('Commissions Master.xlsx',
                                   'Master', dtype=str)
        masterFiles = pd.read_excel('Commissions Master.xlsx',
                                    'Files Processed', dtype=str)
        missCols = [i for i in set(masterComm).union(runCom) if
                    i not in list(masterComm) or i not in list(runCom)]
        masterCols = list(masterComm)
        fileCols = list(masterFiles)
        if missCols:
            print('The following columns were not detected in one of the two '
                  'files:\n%s' %
                  ', '.join(map(str, missCols))
                  + '\n***')
            return
    else:
        print('---\n'
              'No Commissions Master found!\n'
              'Please make sure Commissions Master.xlsx is '
              'in the directory.\n'
              '***')
        return

    # Read in the Master Lookup. Exit if not found.
    if os.path.exists('Lookup Master - Current.xlsx'):
        masterLookup = pd.read_excel('Lookup Master - Current.xlsx').fillna('')
        # Check the column names.
        lookupCols = ['CM Sales', 'Design Sales', 'CM Split',
                      'Reported Customer', 'CM', 'Part Number', 'T-Name',
                      'T-End Cust', 'Last Used', 'Principal', 'City',
                      'Date Added']
        missCols = [i for i in lookupCols if i not in list(masterLookup)]
        if missCols:
            print('The following columns were not detected in '
                  'Lookup Master.xlsx:\n%s' %
                  ', '.join(map(str, missCols))
                  + '\n***')
            return
    else:
        print('---\n'
              'No Lookup Master found!\n'
              'Please make sure Lookup Master - Current.xlsx is '
              'in the directory.\n'
              '***')
        return

    # Go through each line of the finished Running Commissions and use them to
    # update the Lookup Master.
    for row in runCom.index:
        # First match reported customer.
        repCust = str(runCom.loc[row, 'Reported Customer']).lower()
        POSCust = masterLookup['Reported Customer'].map(
                lambda x: str(x).lower())
        custMatches = masterLookup[repCust == POSCust]
        # Now match part number.
        partNum = str(runCom.loc[row, 'Part Number']).lower()
        PPN = masterLookup['Part Number'].map(lambda x: str(x).lower())
        fullMatches = custMatches[PPN == partNum]
        # Figure out if this entry is a duplicate of any existing entry.
        for matchID in fullMatches.index:
            matchCols = ['CM Sales', 'Design Sales', 'CM', 'T-Name',
                         'T-End Cust']
            duplicate = all(fullMatches.loc[matchID, i] == runCom.loc[row, i]
                            for i in matchCols)
            if duplicate:
                break
        # If it's not an exact duplicate, add it to the Lookup Master.
        if not duplicate:
            lookupCols = ['CM Sales', 'Design Sales', 'CM', 'T-Name',
                          'T-End Cust', 'Reported Customer', 'Principal',
                          'Part Number', 'City']
            newLookup = runCom.loc[row, lookupCols]
            newLookup['Date Added'] = datetime.datetime.now().date()
            newLookup['Last Used'] = datetime.datetime.now().date()
            masterLookup = masterLookup.append(newLookup, ignore_index=True)

    # Append the new Running Commissions.
    masterCom = masterCom.append(runCom, ignore_index=True)
    masterFiles = masterFiles.append(filesProcessed, ignore_index=True)
    masterCom = masterCom.loc[:, masterCols]
    masterFiles = masterFiles.loc[:, fileCols]

    # %% Get ready to save files.
    fname1 = 'Commissions Master.xlsx'
    fname2 = 'Lookup Master - Current.xlsx'

    if saveError(fname1, fname2):
        print('---\n'
              'One or more of these files are currently open in Excel:\n'
              'Running Commissions, Entries Need Fixing, Lookup Master.\n'
              'Please close these files and try again.\n'
              '***')
        return

    # Write the Commissions Master file.
    writer1 = pd.ExcelWriter(fname1, engine='xlsxwriter',
                             datetime_format='mm/dd/yyyy')
    masterCom.to_excel(writer1, sheet_name='Master', index=False)
    masterFiles.to_excel(writer1, sheet_name='Files Processed', index=False)
    # Format everything in Excel.
    tableFormat(masterCom, 'Master', writer1)
    tableFormat(masterFiles, 'Files Processed', writer1)

    # Write the Lookup Master.
    writer2 = pd.ExcelWriter(fname2, engine='xlsxwriter',
                             datetime_format='mm/dd/yyyy')
    masterLookup.to_excel(writer2, sheet_name='Lookup', index=False)
    # Format everything in Excel.
    tableFormat(masterLookup, 'Lookup', writer2)

    # Save the files.
    writer1.save()
    writer2.save()

    print('---\n'
          'Updates completed successfully!\n'
          '---\n'
          'Commissions Master updated.\n'
          'Lookup Master updated.\n'
          '+++')
