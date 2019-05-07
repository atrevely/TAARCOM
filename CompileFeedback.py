import pandas as pd
import os
import time
from xlrd import XLRDError


def tableFormat(sheetData, sheetName, wbook):
    """Formats the Excel output as a table with correct column formatting."""
    # Nothing to format, so return.
    if sheetData.shape[0] == 0:
        return
    sheet = wbook.sheets[sheetName]
    # Set document formatting.
    docFormat = wbook.book.add_format({'font': 'Calibri',
                                       'font_size': 11})
    acctFormat = wbook.book.add_format({'font': 'Calibri',
                                        'font_size': 11,
                                        'num_format': 44})
    commaFormat = wbook.book.add_format({'font': 'Calibri',
                                         'font_size': 11,
                                         'num_format': 3})
    newFormat = wbook.book.add_format({'font': 'Calibri',
                                       'font_size': 11,
                                       'bg_color': 'yellow'})
    movedFormat = wbook.book.add_format({'font': 'Calibri',
                                         'font_size': 11,
                                         'bg_color': '#FFCC99'})
    # Format and fit each column.
    i = 0
    # Columns which get shrunk down in reports.
    hideCols = ['Technology', 'Excel Part Link', 'Report Part Nbr Link',
                'MFG Part Description', 'Focus', 'Part Class Name',
                'Vendor ID', 'Invoice Detail Nbr', 'Assigned Account Rep',
                'Recipient', 'DKLI Report Date', 'Invoice Date Group',
                'Comments', 'Sales Channel']
    coreCols = ['Must Contact', 'End Product', 'How Contacted',
                'Information for Digikey']
    for col in sheetData.columns:
        acctCols = ['Unit Price', 'Invoiced Dollars']
        if col in acctCols:
            formatting = acctFormat
        elif col == 'Quantity':
            formatting = commaFormat
        else:
            formatting = docFormat
        maxWidth = max(len(str(val)) for val in sheetData[col].values)
        # Set maximum column width at 50.
        maxWidth = min(maxWidth, 50)
        if col in hideCols:
            maxWidth = 0
        elif col in coreCols:
            maxWidth = 25
        sheet.set_column(i, i, maxWidth+0.8, formatting)
        i += 1
    # Highlight new root customer and moved city rows.
    try:
        for row in sheetData.index:
            if sheetData.loc[row, 'Sales'] == '':
                sheet.write(row+1, 4, sheetData.loc[row, 'Root Customer..'],
                            newFormat)
            elif sheetData.loc[row, 'City on Acct List']:
                sheet.write(row+1, 4, sheetData.loc[row, 'Root Customer..'],
                            movedFormat)
                sheet.write(row+1, 24, sheetData.loc[row, 'City on Acct List'],
                            movedFormat)
    except KeyError:
        print('Error locating Sales and/or City on Acct List columns.\n'
              'Unable to highlight without these columns.\n'
              '---')


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
def main(filepaths):
    """Combine files into one finalized monthly Digikey file, and append it
    to the Digikey Insights Master file. Also updates the rootCustomerMappings
    file.

    Arguments:
    filepaths -- The filepaths to the files with new comments.
    """
    # ------------------
    # Load in the files.
    # ------------------
    # Set the directory paths to the server.
    lookDir = 'Z:/Commissions Lookup/'
    dataDir = 'Z:/Digikey Data/'
    # Load the Digikey Insights Master file.
    if os.path.exists(dataDir + 'Digikey Insight Master.xlsx'):
        insMast = pd.read_excel(dataDir + 'Digikey Insight Master.xlsx',
                                'Master').fillna('')
        filesProcessed = pd.read_excel(dataDir + 'Digikey Insight Master.xlsx',
                                       'Files Processed').fillna('')
    else:
        print('---\n'
              'No Digikey Insight Master file found!\n'
              'Please make sure Digikey Insight Master is in the directory.\n'
              '***')
        return

    # Load the Root Customer Mappings file.
    if os.path.exists(lookDir + 'rootCustomerMappings.xlsx'):
        try:
            rootCustMap = pd.read_excel(lookDir + 'rootCustomerMappings.xlsx',
                                        'Sales Lookup').fillna('')
        except XLRDError:
            print('---\n'
                  'Error reading sheet name for rootCustomerMappings.xlsx!\n'
                  'Please make sure the main tab is named Sales Lookup.\n'
                  '***')
            return
        # Check the column names.
        rootMapCols = ['Root Customer', 'Salesperson']
        missCols = [i for i in rootMapCols if i not in list(rootCustMap)]
        if missCols:
            print('The following columns were not detected in '
                  'rootCustomerMappings.xlsx:\n%s' %
                  ', '.join(map(str, missCols))
                  + '\n***')
            return
    else:
        print('---\n'
              'No Root Customer Mappings file found!\n'
              'Please make sure rootCustomerMappings.xlsx'
              'is in the directory.\n'
              '***')
        return

    # Strip the root off of the filepaths and leave just the filenames.
    filenames = [os.path.basename(i) for i in filepaths]

    # Load the Insight files.
    try:
        inputData = [pd.read_excel(i) for i in filepaths]
    except XLRDError:
        print('---\n'
              'Error reading in files!\n'
              '***')
        return

    # ----------------------------------------------
    # Combine the report data from each salesperson.
    # ----------------------------------------------
    # Make sure each filename has a salesperson initials.
    salespeople = ['CM', 'CR', 'DC', 'HS', 'IT', 'JC', 'JW', 'KC', 'LK',
                   'MG', 'MM', 'VD']
    initList = []
    for filename in filenames:
        inits = filename[0:2]
        initList.append(inits)
        if inits not in salespeople:
            print('Salesperson initials ' + inits + ' not recognized!\n'
                  'Make sure the first two letters of each filename are '
                  'salesperson initials (capitalized).\n'
                  '***')
            return
        elif inits in initList:
            print('Salesperson initials ' + inits + ' duplicated!\n'
                  'Make sure each salesperson has only one file.\n'
                  '***')
            return

    # Create the master dataframe to append to.
    finalData = pd.DataFrame(columns=inputData[0].colums)

    fileNum = 0
    for sheet in inputData:
        print('---\n'
              'Copying comments from file: ' + filenames[fileNum])

        # Grab only the salesperson's data.
        sales = filenames[fileNum][0:2]
        sheetData = sheet[sheet['Sales'] == sales]
        # Append data to the output dataframe.
        finalData = finalData.append(sheetData, ignore_index=True, sort=False)
        # Next file.
        fileNum += 1

    # Drop any unnamed columns that got processed.
    try:
        finalData = finalData.loc[:, ~sheet.columns.str.contains('^Unnamed')]
        finalData = finalData.loc[:, list(inputData[0])]
    except AttributeError:
        pass

    # -------------------------------------
    # Update the rootCustomerMappings file.
    # -------------------------------------
    for row in finalData.index:
        # Get root customer and salesperson.
        cust = sheet.loc[row, 'Root Customer..']
        salesperson = sheet.loc[row, 'Sales']
        if cust and salesperson:
            # Find match in rootCustomerMappings.
            custMatch = rootCustMap['Root Customer'] == cust
            if sum(custMatch) == 1:
                matchID = rootCustMap[custMatch].index
                # Input (possibly new) salesperson.
                rootCustMap.loc[matchID, 'Salesperson'] = salesperson
            elif not custMatch.any():
                # New customer (no match), so append to mappings.
                newCust = pd.DataFrame({'Root Customer': [cust],
                                        'Salesperson': [salesperson]})
                rootCustMap = rootCustMap.append(newCust, ignore_index=True,
                                                 sort=False)
            else:
                print('There appears to be a duplicate customer in'
                      ' rootCustomerMappings:\n'
                      + str(cust) + '\nPlease trim to one entry and try again.'
                      + '\n***')
                return

    # Append the new data to the Digikey Insight Master.
    insMast = insMast.append(finalData, ignore_index=True, sort=False)

    # Try saving the files, exit with error if any file is currently open.
    currentTime = time.strftime('%Y-%m-%d')
    fname1 = 'Digikey Insight Final ' + currentTime + '.xlsx'
    # Append the new file to files processed.
    newFile = pd.DataFrame(columns=filesProcessed.columns)
    newFile.loc[0, 'Filename'] = fname1
    filesProcessed = filesProcessed.append(newFile, ignore_index=True,
                                           sort=False)
    fname2 = 'Digikey Insight Master.xlsx'
    if saveError(fname1, fname2):
        print('---\n'
              'Insight Master and/or Final is currently open in Excel!\n'
              'Please close the file and try again.\n'
              '***')
        return

    # Write the Insight Master file.
    writer1 = pd.ExcelWriter(fname1, engine='xlsxwriter',
                             datetime_format='mm/dd/yyyy')
    finalData.to_excel(writer1, sheet_name='Master', index=False)
    # Format as table in Excel.
    tableFormat(finalData, 'Master', writer1)

    # Write the Insight Master file.
    writer2 = pd.ExcelWriter(fname2, engine='xlsxwriter',
                             datetime_format='mm/dd/yyyy')
    insMast.to_excel(writer2, sheet_name='Master', index=False)
    filesProcessed.to_excel(writer2, sheet_name='Files Processed',
                            index=False)
    # Format as table in Excel.
    tableFormat(insMast, 'Master', writer2)
    tableFormat(filesProcessed, 'Files Processed', writer2)

    # Save the file.
    writer1.save()
    writer2.save()

    print('---\n'
          'Updates completed successfully!\n'
          '---\n'
          'Digikey Master updated.\n'
          '+++')
