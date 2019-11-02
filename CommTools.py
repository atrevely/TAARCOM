import pandas as pd
from RCExcelTools import tableFormat, formDate, saveError
import datetime
from dateutil.parser import parse


def extractLookups(runningCom):
    """Scans a Running Commissions file for new Lookup Master entries and
    copies them over.
    """
    # Set the directory for saving output files.
    lookDir = 'Z:/Commissions Lookup/'

    # ----------------------------------------------
    # Load up the current Running Commissions file.
    # ----------------------------------------------
    runningCom = pd.read_excel(runCom, 'Master', dtype=str)
    # Convert numeric entries in columns that could have numbers in them.
    mixedCols = ['Reported Customer', 'CM', 'Part Number', 'T-Name',
                 'T-End Cust']
    for col in mixedCols:
        runningCom[col] = runningCom[col].map(
                lambda x: pd.to_numeric(x, errors='ignore'))
    runningCom.replace('nan', '', inplace=True)

    # ----------------------------------------------
    # Read in the Master Lookup. Exit if not found.
    # ----------------------------------------------
    if os.path.exists(lookDir + 'Lookup Master - Current.xlsx'):
        mastLook = pd.read_excel(lookDir +
                                 'Lookup Master - Current.xlsx').fillna('')
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
                  + '\n*Program Teminated*')
            return
    else:
        print('No Lookup Master found!\n'
              'Please make sure Lookup Master - Current.xlsx is '
              'in the directory.\n'
              '*Program Teminated*')
        return

    # ------------------------------------------------------------------------
    # Go through each line of the finished Running Commissions and use them to
    # update the Lookup Master.
    # ------------------------------------------------------------------------
    # Don't copy over INDIVIDUAL, MISC, or ALLOWANCE.
    noCopy = ['INDIVIDUAL', 'UNKNOWN', 'ALLOWANCE']
    paredID = [i for i in runningCom.index
               if not any(j in runningCom.loc[i, 'T-End Cust'].upper()
                          for j in noCopy)]
    for row in paredID:
        # First match reported customer.
        repCust = str(runningCom.loc[row, 'Reported Customer']).lower()
        POSCust = masterLookup['Reported Customer'].map(
                lambda x: str(x).lower())
        custMatches = masterLookup[repCust == POSCust]
        # Now match part number.
        partNum = str(runningCom.loc[row, 'Part Number']).lower()
        PPN = masterLookup['Part Number'].map(lambda x: str(x).lower())
        fullMatches = custMatches[PPN == partNum]
        # Figure out if this entry is a duplicate of any existing entry.
        duplicate = False
        for matchID in fullMatches.index:
            matchCols = ['CM Sales', 'Design Sales', 'CM', 'T-Name',
                         'T-End Cust']
            duplicate = all(
                    fullMatches.loc[matchID, i] == runningCom.loc[row, i]
                    for i in matchCols)
            if duplicate:
                break
        # If it's not an exact duplicate, add it to the Lookup Master.
        if not duplicate:
            lookupCols = ['CM Sales', 'Design Sales', 'CM Split', 'CM',
                          'T-Name', 'T-End Cust', 'Reported Customer',
                          'Principal', 'Part Number', 'City']
            newLookup = runningCom.loc[row, lookupCols]
            newLookup['Date Added'] = datetime.datetime.now().date()
            newLookup['Last Used'] = datetime.datetime.now().date()
            masterLookup = masterLookup.append(newLookup,
                                               ignore_index=True)

    # Save the Lookup Master.
    fname = lookDir + 'Lookup Master - Current.xlsx'
    if saveError(fname):
        print('---\n'
              'One or more of these files are currently open in Excel:\n'
              'Running Commissions, Entries Need Fixing, Lookup Master.\n'
              'Please close these files and try again.\n'
              '*Program Terminated*')
        return
    # Write the Lookup Master.
    writer = pd.ExcelWriter(fname, engine='xlsxwriter',
                            datetime_format='mm/dd/yyyy')
    masterLookup.to_excel(writer, sheet_name='Lookup', index=False)
    # Format everything in Excel.
    tableFormat(masterLookup, 'Lookup', writer)
    writer.save()
    print('---\n'
          'Lookup Master updated successfully!\n'
          '+++')


def reIndex(runningCom):
    """
    """
    # Set the directory for saving output files.
    outDir = 'Z:/MK Working Commissions/'
    lookDir = 'Z:/Commissions Lookup/'

    # ----------------------------------------------
    # Load up the current Running Commissions file.
    # ----------------------------------------------
    runningCom = pd.read_excel(runCom, 'Master', dtype=str)
    # Convert applicable columns to numeric.
    numCols = ['Quantity', 'Ext. Cost', 'Invoiced Dollars',
               'Paid-On Revenue', 'Actual Comm Paid', 'Unit Cost',
               'Unit Price', 'CM Split', 'Year', 'Sales Commission',
               'Split Percentage', 'Commission Rate',
               'Gross Rev Reduction', 'Shared Rev Tier Rate']
    for col in numCols:
        try:
            runningCom[col] = pd.to_numeric(runningCom[col],
                                            errors='coerce').fillna('')
        except KeyError:
            pass
    # Convert individual numbers to numeric in rest of columns.
    mixedCols = [col for col in list(runningCom) if col not in numCols]
    # Invoice/part numbers sometimes has leading zeros we'd like to keep.
    mixedCols.remove('Invoice Number')
    mixedCols.remove('Part Number')
    # The INF gets read in as infinity, so skip the principal column.
    mixedCols.remove('Principal')
    for col in mixedCols:
        runningCom[col] = runningCom[col].map(
                lambda x: pd.to_numeric(x, errors='ignore'))
    runningCom.replace('nan', '', inplace=True)
    # Round the Actual Comm Paid field.
    runningCom['Actual Comm Paid'] = runningCom['Actual Comm Paid'].map(
            lambda x: round(float(x), 2))
    filesProcessed = pd.read_excel(runCom, 'Files Processed').fillna('')
    comDate = runCom[-20:]

    # --------------------------------------
    # Load up the Entries Need Fixing file.
    # --------------------------------------
    if os.path.exists(outDir + 'Entries Need Fixing ' + comDate):
        try:
            fixList = pd.read_excel(outDir + 'Entries Need Fixing ' + comDate,
                                    'Data', dtype=str)
            # Convert entries to proper types, like above.
            for col in numCols:
                try:
                    fixList[col] = pd.to_numeric(fixList[col],
                                                 errors='coerce').fillna('')
                except KeyError:
                    pass
            for col in mixedCols:
                fixList[col] = fixList[col].map(
                        lambda x: pd.to_numeric(x, errors='ignore'))
            fixList.replace('nan', '', inplace=True)
            # Round the Actual Comm Paid field.
            fixList['Actual Comm Paid'] = fixList['Actual Comm Paid'].map(
                    lambda x: round(float(x), 2))
        except XLRDError:
            print('Error reading sheet name for Entries Need Fixing.xlsx!\n'
                  'Please make sure the main tab is named Data.\n'
                  '*Program Teminated*')
            return
    else:
        print('No Entries Need Fixing file found!\n'
              'Please make sure Entries Need Fixing ' + comDate
              + ' is in the directory ' + outDir + '.\n'
              '*Program Teminated*')
        return

    # ------------------------------------------------------
    # Go through and start matching the lines in ENF to RC.
    # ------------------------------------------------------
    for row in fixList.index:
        # Match the Reported Customer, Part Number, File, Commissions, and
        # Invoice Number.
        repCust = fixList.loc[row, 'Reported Customer']
        repCustMatch = runningCom[runningCom['Reported Customer'] == repCust]
        partNo = fixList.loc[row, 'Part Number']
        partNoMatch = repCustMatch[repCustMatch['Part Number'] == partNo]
        file = fixList.loc[row, 'From File']
        fileMatch = partNoMatch[partNoMatch['From File'] == file]
        comm = fixList.loc[row, 'Actual Comm Paid']
        commMatch = fileMatch[fileMatch['Actual Comm Paid'] == comm]
        invNo = fixList.loc[row, 'Invoice Number']
        invMatch = commMatch[commMatch['Invoice Number'] == invNo]
        # One match, we're good.
        if len(commMatch) == 1:
            fixList.loc[row, 'Running Com Index'] = str(commMatch.index[0])
        # Multiple matches, find and deal with exact duplicates.
        elif len(commMatch) > 1:
            fixList.loc[row, 'Running Com Index'] = ', '.join(
                    str(i) for i in invMatch.index)
        else:
            fixList.loc[row, 'Running Com Index'] = ''

    # -------------------------------------------
    # Deal with all of the multiple match lines.
    # -------------------------------------------
    multiMatches = fixList[fixList['Running Com Index'].str.contains(',')]
    for match in multiMatches['Running Com Index'].unique():
        matchIDs = match.split(', ')
        matchIDs = [int(i) for i in matchIDs]
        RCMatches = runningCom.loc[matchIDs]
        allMatch = all([all(RCMatches.iloc[0] == i) for i in RCMatches])

    # --------------------------------
    # Check for and clear collisions.
    # --------------------------------
    duplicates = fixList['Running Com Index'].duplicated()
    




def removeData(commMonth):
    """
    """
    # ---------------------------------------------
    # Load and prepare the Commissions Master file.
    # ---------------------------------------------
    dataDir = 'Z:/MK Working Commissions/'
    try:
        comMast = pd.read_excel(dataDir + 'Commissions Master.xlsx',
                                'Master Data', dtype=str)
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
    for col in numCols:
        try:
            comMast[col] = pd.to_numeric(comMast[col],
                                         errors='coerce').fillna(0)
        except KeyError:
            pass
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

    # -----------------------------------------------------
    # Now remove the data that matches the provided month.
    # -----------------------------------------------------
    
