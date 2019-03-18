import pandas as pd
import datetime
import os.path
from xlrd import XLRDError
from RCExcelTools import tableFormat, saveError


# Main function.
def main(filepath):
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
    masterComm = masterComm.append(runCom, ignore_index=True)
    masterFiles = masterFiles.append(filesProcessed, ignore_index=True)
    masterComm = masterComm.loc[:, masterCols]
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
    masterComm.to_excel(writer1, sheet_name='Master', index=False)
    masterFiles.to_excel(writer1, sheet_name='Files Processed', index=False)
    # Format everything in Excel.
    tableFormat(masterComm, 'Master', writer1)
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
