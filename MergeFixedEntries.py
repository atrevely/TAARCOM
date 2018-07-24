import pandas as pd
import time
import datetime
from dateutil.parser import parse
import calendar
import math


def tableFormat(sheetData, sheetName, wbook):
    """Formats the Excel output as a table."""
    # Create the table.
    sheet = wbook.sheets[sheetName]
    header = [{'header': val} for val in sheetData.columns.tolist()]
    setStyle = {'header_row': True, 'style': 'TableStyleMedium5',
                'columns': header, 'font': 'Century Gothic'}
    sheet.add_table(0, 0, len(sheetData.index),
                    len(sheetData.columns)-1, setStyle)
    # Set document formatting.
    docFormat = wbook.book.add_format({'font': 'Century Gothic',
                                       'font_size': 8})
    acctFormat = wbook.book.add_format({'font': 'Century Gothic',
                                        'font_size': 8,
                                        'num_format': 44})
    commaFormat = wbook.book.add_format({'font': 'Century Gothic',
                                         'font_size': 8,
                                         'num_format': 3})
    # Fit to the column width.
    i = 0
    for col in sheetData.columns:
        # Match the correct formatting to each column.
        acctCols = ['Unit Price', 'Invoiced Dollars', 'Paid-On Revenue',
                    'Actual Comm Paid', 'Total NDS', 'Post-Split NDS',
                    'Customer Revenue YTD']
        if col in acctCols:
            formatting = acctFormat
        elif col == 'Quantity':
            formatting = commaFormat
        else:
            formatting = docFormat
        # Set column width and formatting
        maxWidth = max([len(str(val)) for val in sheetData[col].values])
        sheet.set_column(i, i, maxWidth+0.8, formatting)
        i += 1


def saveError(*excelFiles):
    """Try saving Excel files and return True if any file is open."""
    for file in excelFiles:
        try:
            file.save()
        except IOError:
            return True
    return False


# The main function.
def main():
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
    # Load up the Entries Need Fixing file.
    fixList = pd.read_excel('Entries Need Fixing.xlsx', 'Data').fillna('')

    # Load up the current Running Commissions file.
    runningCom = pd.read_excel('Running Commissions.xlsx',
                               'Master').fillna('')
    filesProcessed = pd.read_excel('Running Commissions.xlsx',
                                   'Files Processed').fillna('')

    # Load up the Master Lookup.
    mastLook = pd.read_excel('Lookup Master 6-27-18.xlsx').fillna('')

    # Load the Quarantined Lookups.
    quarantinedLookups = pd.read_excel('Quarantined Lookups.xlsx').fillna('')

    # Grab the lines that have been fixed.
    dateFixed = fixList['Invoice Date'] != ''
    endCustFixed = fixList['T-End Cust'] != ''
    distFixed = fixList['Corrected Distributor'] != ''

    fixdDat = fixList[[x and y and z for x, y, z in zip(dateFixed,
                                                        endCustFixed,
                                                        distFixed)]]
    fixdDat.reset_index(inplace=True, drop=True)

    # %%
    # Go through each entry that's fixed and replace it in Running Commissions.
    for entry in range(len(fixdDat)):
        # Replace the Running Commissions entry with the fixed one.
        RCIndex = fixdDat.loc[entry, 'Running Com Index']
        runningCom.loc[RCIndex, :] = fixdDat.loc[entry, :]

        # Try parsing the date.
        dateError = 0
        try:
            parse(fixdDat.loc[entry, 'Invoice Date'])
        except ValueError:
            dateError = 1
        except TypeError:
            # Check if Pandas read it in as a Timestamp object.
            # If so, turn it back into a string.
            invDate = fixdDat.loc[entry, 'Invoice Date']
            if isinstance(invDate, pd.Timestamp):
                fixdDat.loc[entry, 'Invoice Date'] = str(invDate)
            else:
                dateError = 1
        # If no error found in date, finish filling out the fixed entry.
        if not dateError:
            date = parse(fixdDat.loc[entry, 'Invoice Date'])
            # Cast date format into mm/dd/yyyy.
            fixdDat.loc[entry, 'Invoice Date'] = date.strftime('%m/%d/%Y')
            # Fill in quarter/year/month data.
            fixdDat.loc[entry, 'Year'] = date.year
            fixdDat.loc[entry, 'Month'] = calendar.month_name[date.month][0:3]
            fixdDat.loc[entry, 'Quarter'] = (str(date.year)
                                             + 'Q'
                                             + str(math.ceil(date.month/3)))
            # Delete the fixed entry from the Needs Fixing file.
            fixIndex = fixList['Running Com Index']
            fixList.drop(fixList[fixIndex == RCIndex].index, inplace=True)

            # Append entry to Lookup Master, if applicable.
            if not fixdDat.loc[entry, 'Lookup Master Matches']:
                mastLook = mastLook.append(fixdDat.loc[entry, list(mastLook)],
                                           ignore_index=True).fillna('')
                # Record date that the new entry was added to Lookup Master.
                mastLook.loc[len(mastLook)-1, 'Date Added'] = time.strftime(
                        '%m/%d/%Y')

    # %%
    # Check if any entries are duplicates, then quarantine old versions.
    duplicates = mastLook.duplicated(subset=['POSCustomer', 'PPN'],
                                     keep='last')
    deprecatedEntries = mastLook[duplicates].reset_index(drop=True)
    mastLook = mastLook[~duplicates].reset_index(drop=True)
    # Check for entries that are too old and quarantine them.
    twoYearsAgo = datetime.datetime.today() - datetime.timedelta(days=720)
    dateCutoff = mastLook['Last Used'] < twoYearsAgo.strftime('%m/%d/%Y')
    oldEntries = mastLook[dateCutoff].reset_index(drop=True)
    mastLook = mastLook[~dateCutoff].reset_index(drop=True)
    # Record the date we quarantined the entries.
    deprecatedEntries.loc[:, 'Date Quarantined'] = time.strftime('%m/%d/%Y')
    oldEntries.loc[:, 'Date Quarantined'] = time.strftime('%m/%d/%Y')
    # Add deprecated entries to the quarantine.
    quarantinedLookups = quarantinedLookups.append(oldEntries,
                                                   ignore_index=True)
    quarantinedLookups = quarantinedLookups.append(deprecatedEntries,
                                                   ignore_index=True)
    # Notify us of changes.
    print(str(len(oldEntries))
          + 'entries quarantied for being more than 2 years old.\n'
          + str(len(deprecatedEntries))
          + 'entries quarantined for being deprecated (old duplicates).')

    # Write the Running Commissions file.
    writer1 = pd.ExcelWriter('Running Commissions '
                             + time.strftime('%Y-%m-%d-%H%M')
                             + '.xlsx', engine='xlsxwriter')
    runningCom.to_excel(writer1, sheet_name='Master', index=False)
    filesProcessed.to_excel(writer1, sheet_name='Files Processed',
                            index=False)
    # Format as table in Excel.
    tableFormat(runningCom, 'Master', writer1)
    tableFormat(filesProcessed, 'Files Processed', writer1)

    # Write the Needs Fixing file.
    writer2 = pd.ExcelWriter('Entries Need Fixing.xlsx', engine='xlsxwriter')
    fixList.to_excel(writer2, sheet_name='Data', index=False)
    # Format as table in Excel.
    tableFormat(fixList, 'Data', writer2)

    # Write the Lookup Master file.
    writer3 = pd.ExcelWriter('Lookup Master - Current.xlsx',
                             engine='xlsxwriter')
    mastLook.to_excel(writer3, sheet_name='Lookup', index=False)
    # Format as table in Excel.
    tableFormat(mastLook, 'Lookup', writer3)

    # Write the Quarantined Lookups file.
    writer4 = pd.ExcelWriter('Quarantined Lookups.xlsx', engine='xlsxwriter')
    quarantinedLookups.to_excel(writer4, sheet_name='Lookup', index=False)
    # Format as table in Excel.
    tableFormat(quarantinedLookups, 'Lookup', writer4)

    # Try saving the files, exit with error if any file is currently open.
    if saveError(writer1, writer2, writer3, writer4):
        print('---\n'
              'One or more files are currently open in Excel!\n'
              'Please close the files and try again.\n'
              '***')
        return

    # No errors, so save the files.
    writer1.save()
    writer2.save()
    writer3.save()
    writer4.save()
