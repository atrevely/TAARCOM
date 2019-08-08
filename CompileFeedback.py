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
    # Set the autofilter for the sheet.
    sheet.autofilter(0, 0, sheetData.shape[0], sheetData.shape[1]-1)


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

    # Set the directory paths to the server.
    lookDir = 'Z:/Commissions Lookup/'
    dataDir = 'W:/'

    # ---------------------------------------
    # Load the Digikey Insights Master file.
    # ---------------------------------------
    if os.path.exists(lookDir + 'Digikey Insight Master.xlsx'):
        insMast = pd.read_excel(lookDir + 'Digikey Insight Master.xlsx',
                                'Master').fillna('')
        filesProcessed = pd.read_excel(lookDir + 'Digikey Insight Master.xlsx',
                                       'Files Processed').fillna('')
    else:
        print('---\n'
              'No Digikey Insight Master file found!\n'
              'Please make sure Digikey Insight Master is in the directory.\n'
              '*Program Terminated*')
        return

    # -----------------------------------
    # Load the Master Account List file.
    # -----------------------------------
    if os.path.exists(lookDir + 'Master Account List.xlsx'):
        try:
            mastAcct = pd.read_excel(lookDir + 'Master Account List.xlsx',
                                     'Allacct').fillna('')
        except XLRDError:
            print('---\n'
                  'Error reading sheet name for Master Account List.xlsx!\n'
                  'Please make sure the main tab is named Allacct.\n'
                  '*Program Terminated*')
            return
        # Check the column names.
        mastCols = ['ProperName', 'SLS', 'CITY']
        missCols = [i for i in mastCols if i not in list(mastAcct)]
        if missCols:
            print('The following columns were not detected in '
                  'Master Account List.xlsx:\n%s' %
                  ', '.join(map(str, missCols))
                  + '\nRemember to delete lines before the column '
                  'headers.\n*Program Terminated*')
            return
    else:
        print('---\n'
              'No Master Account List file found!\n'
              'Please make sure the Master Account List '
              'is in the directory.\n'
              '*Program Terminated*')
        return

    # --------------------------------------
    # Load the Root Customer Mappings file.
    # --------------------------------------
    if os.path.exists(lookDir + 'rootCustomerMappings.xlsx'):
        try:
            rootCustMap = pd.read_excel(lookDir + 'rootCustomerMappings.xlsx',
                                        'Sales Lookup').fillna('')
        except XLRDError:
            print('---\n'
                  'Error reading sheet name for rootCustomerMappings.xlsx!\n'
                  'Please make sure the main tab is named Sales Lookup.\n'
                  '*Program Terminated*')
            return
        # Check the column names.
        rootMapCols = ['Root Customer', 'Salesperson']
        missCols = [i for i in rootMapCols if i not in list(rootCustMap)]
        if missCols:
            print('The following columns were not detected in '
                  'rootCustomerMappings.xlsx:\n%s' %
                  ', '.join(map(str, missCols))
                  + '\n*Program Terminated*')
            return
    else:
        print('---\n'
              'No Root Customer Mappings file found!\n'
              'Please make sure rootCustomerMappings.xlsx'
              'is in the directory.\n'
              '*Program Terminated*')
        return

    # --------------------------------
    # Load the Salesperson Info file.
    # --------------------------------
    if os.path.exists(lookDir + 'Salespeople Info.xlsx'):
        try:
            salesInfo = pd.read_excel(lookDir + 'Salespeople Info.xlsx',
                                      'Info')
        except XLRDError:
            print('---\n'
                  'Error reading sheet name for Salespeople Info.xlsx!\n'
                  'Please make sure the main tab is named Info.\n'
                  '*Program terminated*')
            return
    else:
        print('---\n'
              'No Salespeople Info file found!\n'
              'Please make sure Salespeople Info.xlsx is in the directory.\n'
              '*Program terminated*')
        return

    # ------------------------
    # Load the Insight files.
    # ------------------------
    # Strip the root off of the filepaths and leave just the filenames.
    filenames = [os.path.basename(i) for i in filepaths]
    try:
        inputData = [pd.read_excel(i) for i in filepaths]
    except XLRDError:
        print('---\n'
              'Error reading in files!\n'
              '*Program Terminated*')
        return

    # ----------------------------------------------
    # Combine the report data from each salesperson.
    # ----------------------------------------------
    # Make sure each filename has a salesperson initials.
    salespeople = salesInfo['Sales Initials'].values
    initList = []
    for filename in filenames:
        inits = filename[0:2]
        if inits not in salespeople:
            print('Salesperson initials ' + inits + ' not recognized!\n'
                  'Make sure the first two letters of each filename are '
                  'salesperson initials (capitalized).\n'
                  '*Program Terminated*')
            return
        elif inits in initList:
            print('Salesperson initials ' + inits + ' duplicated!\n'
                  'Make sure each salesperson has at most one file.\n'
                  '*Program Terminated*')
            return
        initList.append(inits)
    # Create the master dataframe to append to.
    finalData = pd.DataFrame(columns=inputData[0].columns)
    # Copy over the comments.
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
                      + '\n*Program Terminated*')
                return

    # ----------------------------------------------------------------
    # Append the new data to the Digikey Insight Master, then update
    # the Current Salesperson.
    # ----------------------------------------------------------------
    mastCols = list(insMast)
    mastCols.remove('Current Sales')
    insMast = insMast.append(finalData[mastCols],
                             ignore_index=True, sort=False)
    insMast.fillna('', inplace=True)
    finalData.fillna('', inplace=True)
    # Go through each root customer and update current salesperson.
    for cust in insMast['Root Customer..'].unique():
        currentSales = ''
        # First check the Account List.
        acctMatch = mastAcct[mastAcct['ProperName'] == cust]
        if not acctMatch.empty:
            currentSales = acctMatch['SLS'].iloc[0]
        # Next try rootCustomerMappings.
        mapMatch = rootCustMap[rootCustMap['Root Customer'] == cust]
        if acctMatch.empty and not mapMatch.empty:
            currentSales = mapMatch['Current Sales'].iloc[0]
        # Update current salesperson.
        matchID = insMast[insMast['Root Customer..'] == cust].index
        insMast.loc[matchID, 'Current Sales'] = currentSales

    # ---------------------------------------------------------------------
    # Try saving the files, exit with error if any file is currently open.
    # ---------------------------------------------------------------------
    currentTime = time.strftime('%Y-%m-%d')
    fname1 = dataDir + 'Digikey Insight Final ' + currentTime + '.xlsx'
    # Append the new file to files processed.
    newFile = pd.DataFrame(columns=filesProcessed.columns)
    newFile.loc[0, 'Filename'] = fname1
    filesProcessed = filesProcessed.append(newFile, ignore_index=True,
                                           sort=False)
    fname2 = lookDir + 'Digikey Insight Master.xlsx'
    if saveError(fname1, fname2):
        print('---\n'
              'Insight Master and/or Final is currently open in Excel!\n'
              'Please close the file and try again.\n'
              '*Program Terminated*')
        return
    # Write the finished Insight file.
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
    # Save the files.
    writer1.save()
    writer2.save()
    print('---\n'
          'Updates completed successfully!\n'
          '---\n'
          'Digikey Master updated.\n'
          '+Program Complete+')
