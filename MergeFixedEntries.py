import pandas as pd
import time
import datetime
from dateutil.parser import parse
import calendar
import math
import os.path
from xlrd import XLRDError


def tableFormat(sheetData, sheetName, wbook):
    """Formats the Excel output as a table with correct column formatting."""
    # Create the table.
    sheet = wbook.sheets[sheetName]
    header = [{'header': val} for val in sheetData.columns.tolist()]
    setStyle = {'header_row': True, 'style': 'TableStyleLight1',
                'columns': header}
    sheet.add_table(0, 0, len(sheetData.index),
                    len(sheetData.columns)-1, setStyle)
    # Set document formatting.
    docFormat = wbook.book.add_format({'font': 'Calibri',
                                       'font_size': 11})
    acctFormat = wbook.book.add_format({'font': 'Calibri',
                                        'font_size': 11,
                                        'num_format': 44})
    commaFormat = wbook.book.add_format({'font': 'Calibri',
                                         'font_size': 11,
                                         'num_format': 3})
    estFormat = wbook.book.add_format({'font': 'Calibri',
                                       'font_size': 11,
                                       'num_format': 44,
                                       'bg_color': 'yellow'})
    pctFormat = wbook.book.add_format({'font': 'Calibri',
                                       'font_size': 11,
                                       'num_format': '0.0%'})
    # Format and fit each column.
    i = 0
    for col in sheetData.columns:
        # Match the correct formatting to each column.
        acctCols = ['Unit Price', 'Paid-On Revenue', 'Actual Comm Paid',
                    'Total NDS', 'Post-Split NDS', 'Cust Revenue YTD',
                    'Ext. Cost', 'Unit Cost', 'Total Commissions',
                    'Sales Commission']
        pctCols = ['Split Percentage', 'Commission Rate',
                   'Gross Rev Reduction', 'Shared Rev Tier Rate']
        if col in acctCols:
            formatting = acctFormat
        elif col in pctCols:
            formatting = pctFormat
        elif col == 'Quantity':
            formatting = commaFormat
        elif col == 'Invoiced Dollars':
            # Highlight any estimates in Invoiced Dollars...
            for row in range(len(sheetData[col])):
                if sheetData.loc[row, 'Ext. Cost']:
                    sheet.write(row+1, i,
                                sheetData.loc[row, 'Invoiced Dollars'],
                                estFormat)
                else:
                    sheet.write(row+1, i,
                                sheetData.loc[row, 'Invoiced Dollars'],
                                acctFormat)
            # Formatting already done, so leave blank.
            formatting = []
        else:
            formatting = docFormat
        # Set column width and formatting.
        if sheetData.shape[0] > 0:
            maxWidth = max(len(str(val)) for val in sheetData[col].values)
            maxWidth = min(maxWidth, 50)
            sheet.set_column(i, i, maxWidth+0.8, formatting)
            i += 1
        else:
            sheet.set_column(i, i, len(col), formatting)


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


# The main function.
def main(runCom):
    """Replaces bad entries in Running Commissions with their fixed versions.

    Entries in Running Commissions which need attention are copied to the
    Entries Need Fixing file. This function merges fixed entries in the Need
    Fixing file into the Running Commissions file by overwriting the existing
    (bad) entry with the fixed one, then removing it from the Needs Fixing
    file.

    Additionally, this function maintains the Lookup Master by adding new
    entries when needed, and quarantining old entries that have not been
    used in 2+ years.
    """
    # Load up the current Running Commissions file.
    runningCom = pd.read_excel(runCom, 'Master').fillna('')
    filesProcessed = pd.read_excel(runCom, 'Files Processed').fillna('')
    comDate = runCom[-20:]

    # Load up the Entries Need Fixing file.
    if os.path.exists('Entries Need Fixing ' + comDate):
        try:
            fixList = pd.read_excel('Entries Need Fixing ' + comDate,
                                    'Data').fillna('')
        except XLRDError:
            print('Error reading sheet name for Entries Need Fixing.xlsx!\n'
                  'Please make sure the main tab is named Data.\n'
                  '***')
            return
    else:
        print('No Entries Need Fixing file found!\n'
              'Please make sure Entries Need Fixing.xlsx '
              'is in the directory.\n'
              '***')
        return

    # Read in the Master Lookup. Exit if not found.
    if os.path.exists('Lookup Master - Current.xlsx'):
        mastLook = pd.read_excel('Lookup Master - Current.xlsx').fillna('')
        # Check the column names.
        lookupCols = ['CM Sales', 'Design Sales', 'CM Split',
                      'Reported Customer', 'CM', 'Part Number', 'T-Name',
                      'T-End Cust', 'Last Used', 'Principal', 'City',
                      'Date Added']
        missCols = [i for i in lookupCols if i not in list(mastLook)]
        if missCols:
            print('The following columns were not detected in '
                  'Lookup Master.xlsx:\n%s' %
                  ', '.join(map(str, missCols))
                  + '\n***')
            return
    else:
        print('No Lookup Master found!\n'
              'Please make sure lookupMaster.xlsx is in the directory.\n'
              '***')
        return

    # Load the Quarantined Lookups.
    if os.path.exists('Quarantined Lookups.xlsx'):
        quarantined = pd.read_excel('Quarantined Lookups.xlsx').fillna('')
    else:
        print('No Quarantied Lookups file found!\n'
              'Please make sure Quarantined Lookups.xlsx '
              'is in the directory.\n'
              '***')
        return

    # Grab the lines that have been fixed.
    TNameFixed = fixList[fixList['T-End Cust'] != '']
    fixed = TNameFixed[TNameFixed['Invoice Date'] != '']
    # Return if there's nothing fixed.
    if fixed.shape[0] == 0:
        print('No new fixed entries detected.\n'
              '***')
        return

    # %%
    print('Writing fixed entries...')
    # Go through each entry that's fixed and replace it in Running Commissions.
    for row in fixed.index:
        dateError = False
        dateGiven = fixed.loc[row, 'Invoice Date']
        # Check if the date is read in as a float/int, and convert to string.
        if isinstance(fixed.loc[row, 'Invoice Date'], (float, int)):
            dateGiven = str(int(dateGiven))
        # Check if Pandas read it in as a Timestamp object.
        # If so, turn it back into a string (a bit roundabout, oh well).
        elif isinstance(dateGiven, pd.Timestamp):
            dateGiven = str(dateGiven)
        try:
            parse(dateGiven)
        except (ValueError, TypeError):
            # The date isn't recognized by the parser.
            dateError = True
        except KeyError:
            print('There is no Invoice Date column in Entries Need Fixing!\n'
                  'Please check to make sure an Invoice Date column exists.\n'
                  'Note: Spelling, whitespace, and capitalization matter.\n'
                  '---')
            dateError = True
        # If no error found in date, finish filling out the fixed entry.
        if not dateError:
            date = parse(dateGiven).date()
            # Make sure the date actually makes sense.
            currentYear = int(time.strftime('%Y'))
            if currentYear - date.year not in [0, 1]:
                dateError = True
            else:
                # Cast date format into mm/dd/yyyy.
                fixed.loc[row, 'Invoice Date'] = date
                # Fill in quarter/year/month data.
                fixed.loc[row, 'Year'] = date.year
                fixed.loc[row, 'Month'] = calendar.month_name[date.month][0:3]
                fixed.loc[row, 'Quarter Shipped'] = (str(date.year)
                                                     + 'Q'
                                                     + str(math.ceil(date.month/3)))

            # Replace the Running Commissions entry with the fixed one.
            RCIndex = fixed.loc[row, 'Running Com Index']
            runningCom.loc[RCIndex, :] = fixed.loc[row, list(runningCom)]

            # Append entry to Lookup Master, if applicable.
            # Check if entry is individual, misc, or unknown.
            skips = ['UNKNOWN', 'MISC', 'INDIVIDUAL']
            tName = fixed.loc[row, 'T-Name'].upper()
            if not any(i for i in skips if i in tName):
                # Match the part number.
                ppn = fixed.loc[row, 'Part Number']
                ppnMatch = mastLook[mastLook['Part Number'] == ppn]

                # Match the reported customer.
                repCust = fixed.loc[row, 'Reported Customer']
                custMatch = ppnMatch[ppnMatch['Reported Customer'] == repCust]

                # Check if there's already an entry for this customer/PPN.
                if len(custMatch) == 0:
                    # Create new lookup entry.
                    lookupEntry = fixList.loc[row, ['CM Sales', 'Design Sales',
                                                    'Reported Customer',
                                                    'T-End Cust', 'T-Name',
                                                    'CM', 'Principal',
                                                    'Part Number', 'City']]
                    lookupEntry['Date Added'] = datetime.datetime.now().date()
                    invDate = pd.Timestamp(fixList.loc[row, 'Invoice Date'])
                    lookupEntry['Last Used'] = invDate.strftime('%m/%d/%Y')
                    # Merge to the Lookup Master.
                    mastLook = mastLook.append(lookupEntry, ignore_index=True)

            # Delete the fixed entry from the Needs Fixing file.
            fixList.drop(row, inplace=True)

    # %%
    # Re-index the fix list and drop nans in Lookup Master
    fixList.reset_index(drop=True, inplace=True)
    mastLook.fillna('', inplace=True)
    # Check for entries that are too old and quarantine them.
    twoYearsAgo = datetime.datetime.today() - datetime.timedelta(days=720)
    try:
        lastUsed = mastLook['Last Used'].map(lambda x: pd.Timestamp(x))
        lastUsed = lastUsed.map(lambda x: x.strftime('%Y%m%d'))
    except AttributeError:
        print('Error reading one or more dates in the Lookup Master!\n'
              'Make sure the Last Used column is all MM/DD/YYYY format.\n'
              '***')
        return
    dateCutoff = lastUsed < twoYearsAgo.strftime('%Y%m%d')
    oldEntries = mastLook[dateCutoff].reset_index(drop=True)
    mastLook = mastLook[~dateCutoff].reset_index(drop=True)
    if oldEntries.shape[0] > 0:
        # Record the date we quarantined the entries.
        oldEntries.loc[:, 'Date Quarantined'] = datetime.datetime.now().date()
        # Add deprecated entries to the quarantine.
        quarantined = quarantined.append(oldEntries,
                                         ignore_index=True)
        # Notify us of changes.
        print(str(len(oldEntries))
              + ' entries quarantied for being more than 2 years old.\n')

    # Check if the files we're going to save are open already.
    fname1 = 'Running Commissions ' + time.strftime('%Y-%m-%d-%H%M') + '.xlsx'
    fname2 = 'Entries Need Fixing.xlsx'
    fname3 = 'Lookup Master - Current.xlsx'
    fname4 = 'Quarantined Lookups.xlsx'
    if saveError(fname1, fname2, fname3, fname4):
        print('---\n'
              'One or more files are currently open in Excel!\n'
              'Please close the files and try again.\n'
              '***')
        return

    # Write the Running Commissions file.
    writer1 = pd.ExcelWriter(fname1, engine='xlsxwriter')
    runningCom.to_excel(writer1, sheet_name='Master', index=False)
    filesProcessed.to_excel(writer1, sheet_name='Files Processed',
                            index=False)
    # Format as table in Excel.
    tableFormat(runningCom, 'Master', writer1)
    tableFormat(filesProcessed, 'Files Processed', writer1)

    # Write the Needs Fixing file.
    writer2 = pd.ExcelWriter(fname2, engine='xlsxwriter')
    fixList.to_excel(writer2, sheet_name='Data', index=False)
    # Format as table in Excel.
    tableFormat(fixList, 'Data', writer2)

    # Write the Lookup Master file.
    writer3 = pd.ExcelWriter(fname3, engine='xlsxwriter')
    mastLook.to_excel(writer3, sheet_name='Lookup', index=False)
    # Format as table in Excel.
    tableFormat(mastLook, 'Lookup', writer3)

    # Write the Quarantined Lookups file.
    writer4 = pd.ExcelWriter(fname4, engine='xlsxwriter')
    quarantined.to_excel(writer4, sheet_name='Lookup', index=False)
    # Format as table in Excel.
    tableFormat(quarantined, 'Lookup', writer4)

    # Save the files.
    writer1.save()
    writer2.save()
    writer3.save()
    writer4.save()

    print('---\n'
          'Fixed entries copied over successfully!\n'
          '+++')
