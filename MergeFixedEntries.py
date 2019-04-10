import pandas as pd
import time
import datetime
from dateutil.parser import parse
import calendar
import math
import os.path
from xlrd import XLRDError
from RCExcelTools import tableFormat, saveError, formDate


# %% The main function.
def main(runCom):
    """Replaces incomplete entries in Running Commissions with final versions.

    Entries in Running Commissions which need attention are copied to the
    Entries Need Fixing file. This function merges fixed entries in the Need
    Fixing file into the Running Commissions file by overwriting the existing
    (bad) entry with the fixed one, then removing it from the Needs Fixing
    file.

    Additionally, this function maintains the Lookup Master by adding new
    entries when needed, and quarantining old entries that have not been
    used in 2+ years.
    """
    # -------------------------
    # Load in all of the files.
    # -------------------------
    # Load up the current Running Commissions file.
    runningCom = pd.read_excel(runCom, 'Master', dtype=str)
    runningCom.replace('nan', '', inplace=True)
    filesProcessed = pd.read_excel(runCom, 'Files Processed').fillna('')
    comDate = runCom[-20:]

    # Set the directory for saving output files.
    outDir = ''#'Z:/MK Working Commissions/'

    # Track commission dollars.
    try:
        comm = pd.to_numeric(runningCom['Actual Comm Paid'],
                             errors='raise').fillna(0)
        totComm = sum(comm)
    except ValueError:
        print('Non-numeric entry detected in Actual Comm Paid!\n'
              'Check the Actual Comm Paid column for bad data and try again.\n'
              '***')
        return

    # Load up the Entries Need Fixing file.
    if os.path.exists(outDir + 'Entries Need Fixing ' + comDate):
        try:
            fixList = pd.read_excel(outDir + 'Entries Need Fixing ' + comDate,
                                    'Data', dtype=str)
            fixList.replace('nan', '', inplace=True)
        except XLRDError:
            print('Error reading sheet name for Entries Need Fixing.xlsx!\n'
                  'Please make sure the main tab is named Data.\n'
                  '***')
            return
    else:
        print('No Entries Need Fixing file found!\n'
              'Please make sure Entries Need Fixing ' + comDate
              + ' is in the directory.\n'
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
                  'Lookup Master - Current.xlsx:\n%s' %
                  ', '.join(map(str, missCols))
                  + '\n***')
            return
    else:
        print('No Lookup Master found!\n'
              'Please make sure Lookup Master - Current.xlsx is '
              'in the directory.\n'
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

    # -----------------------------------------
    # Get the data that's ready to be migrated.
    # -----------------------------------------
    # Grab the lines that have an End Customer.
    endCustFixed = fixList[fixList['T-End Cust'] != '']
    # Grab entries where salespeople are filled in.
    CMSales = endCustFixed['CM Sales'] != ''
    DesignSales = endCustFixed['Design Sales'] != ''
    fixed = endCustFixed[[x or y for x, y in zip(CMSales, DesignSales)]]
    # Return if there's nothing fixed.
    if fixed.shape[0] == 0:
        print('No new fixed entries detected.\n'
              'Entries need a T-End Cust, Salespeople, and an Invoice Date '
              'in order to be eligible for migration to Running Commissions.\n'
              '***')
        return

    # %% Start the process of writing over fixed entries.
    print('Writing fixed entries...')
    # Go through each entry that's fixed and replace it in Running Commissions.
    for row in fixed.index:
        # -------------------------------
        # Make sure the date makes sense.
        # -------------------------------
        dateError = False
        dateGiven = fixed.loc[row, 'Invoice Date']
        # Check if the date is read in as a float/int, and convert to string.
        if isinstance(dateGiven, (float, int)):
            dateGiven = str(int(dateGiven))
        # Check if Pandas read it in as a Timestamp object.
        # If so, turn it back into a string (a bit roundabout, oh well).
        elif isinstance(dateGiven, (pd.Timestamp,  datetime.datetime)):
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
        # --------------------------------------------------------------
        # If no error found in date, finish filling out the fixed entry.
        # --------------------------------------------------------------
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
                Qtr = str(math.ceil(date.month/3))
                fixed.loc[row, 'Quarter Shipped'] = (str(date.year) + 'Q'
                                                     + Qtr)
            # Check for match in commission dollars.
            try:
                RCIndex = pd.to_numeric(fixed.loc[row, 'Running Com Index'],
                                        errors='raise')
            except ValueError:
                print('Error reading Running Com Index!\n'
                      'Make sure all values are numeric.\n'
                      '***')
                return
            comm = fixed.loc[row, 'Actual Comm Paid']
            if runningCom.loc[RCIndex, 'Actual Comm Paid'] == comm:
                # Replace the Running Commissions entry with the fixed one.
                runningCom.loc[RCIndex, :] = fixed.loc[row, list(runningCom)]
            else:
                print('Mismatch in commission dollars found in Entries '
                      'Need Fixing on row '
                      + str(row + 2)
                      + '\n***')
                return
            # Delete the fixed entry from the Needs Fixing file.
            fixList.drop(row, inplace=True)

    # %%
    # Make sure all the dates are formatted correctly.
    runningCom['Invoice Date'] = runningCom['Invoice Date'].map(
            lambda x: formDate(x))
    mastLook['Last Used'] = mastLook['Last Used'].map(lambda x: formDate(x))
    mastLook['Date Added'] = mastLook['Date Added'].map(lambda x: formDate(x))
    # Go through each column and convert applicable entries to numeric.
    cols = list(runningCom)
    # Invoice number sometimes has leading zeros we'd like to keep.
    cols.remove('Invoice Number')
    # The INF gets read in as infinity, so skip the principal column.
    cols.remove('Principal')
    for col in cols:
        runningCom[col] = pd.to_numeric(runningCom[col], errors='ignore')
    # Check to make sure commission dollars still match.
    comm = pd.to_numeric(runningCom['Actual Comm Paid'],
                         errors='coerce').fillna(0)
    if sum(comm) != totComm:
        print('Commission dollars do not match after fixing entries!\n'
              'Make sure Entries Need fixing aligns properly with '
              'Running Commissions.\n'
              'This error was likely caused by adding or removing rows '
              'in either file.\n'
              '***')
        return
    # Re-index the fix list and drop nans in Lookup Master.
    fixList.reset_index(drop=True, inplace=True)
    mastLook.fillna('', inplace=True)
    # Check for entries that are too old and quarantine them.
    twoYearsAgo = datetime.datetime.today() - datetime.timedelta(days=720)
    try:
        lastUsed = mastLook['Last Used'].map(lambda x: pd.Timestamp(x))
        lastUsed = lastUsed.map(lambda x: x.strftime('%Y%m%d'))
    except (AttributeError, ValueError):
        print('Error reading one or more dates in the Lookup Master!\n'
              'Make sure the Last Used column is all MM/DD/YYYY format.\n'
              '---')
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
              + ' entries quarantied for being more than 2 years old.\n'
              '---')

    # Check if the files we're going to save are open already.
    fname1 = outDir + 'Running Commissions ' + comDate
    fname2 = outDir + 'Entries Need Fixing ' + comDate
    fname3 = 'Lookup Master - Current.xlsx'
    fname4 = 'Quarantined Lookups.xlsx'
    if saveError(fname1, fname2, fname3, fname4):
        print('---\n'
              'One or more files are currently open in Excel!\n'
              'Please close the files and try again.\n'
              '***')
        return

    # Write the Running Commissions file.
    writer1 = pd.ExcelWriter(fname1, engine='xlsxwriter',
                             datetime_format='mm/dd/yyyy')
    runningCom.to_excel(writer1, sheet_name='Master', index=False)
    filesProcessed.to_excel(writer1, sheet_name='Files Processed',
                            index=False)
    # Format as table in Excel.
    tableFormat(runningCom, 'Master', writer1)
    tableFormat(filesProcessed, 'Files Processed', writer1)

    # Write the Needs Fixing file.
    writer2 = pd.ExcelWriter(fname2, engine='xlsxwriter',
                             datetime_format='mm/dd/yyyy')
    fixList.to_excel(writer2, sheet_name='Data', index=False)
    # Format as table in Excel.
    tableFormat(fixList, 'Data', writer2)

    # Write the Lookup Master file.
    writer3 = pd.ExcelWriter(fname3, engine='xlsxwriter',
                             datetime_format='mm/dd/yyyy')
    mastLook.to_excel(writer3, sheet_name='Lookup', index=False)
    # Format as table in Excel.
    tableFormat(mastLook, 'Lookup', writer3)

    # Write the Quarantined Lookups file.
    writer4 = pd.ExcelWriter(fname4, engine='xlsxwriter',
                             datetime_format='mm/dd/yyyy')
    quarantined.to_excel(writer4, sheet_name='Lookup', index=False)
    # Format as table in Excel.
    tableFormat(quarantined, 'Lookup', writer4)

    # Save the files.
    writer1.save()
    writer2.save()
    writer3.save()
    writer4.save()

    print('Fixed entries migrated successfully!\n'
          '+++')
