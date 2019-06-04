import pandas as pd
import os
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
                                         'bg_color': '#FF9900'})
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


# %% The main function.
def main(filepath):
    """Looks up the salespeople for a Digikey Local Insight file.

    Arguments:
    filepath -- The filepath to the new Digikey Insight file.
    """
    # ----------------------------
    # Load in the necessary files.
    # ----------------------------
    # Set the directory paths to the server.
    lookDir = 'Z:/Commissions Lookup/'
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

    # Load the Master Account List file.
    if os.path.exists(lookDir + 'Master Account List.xlsx'):
        try:
            mastAcct = pd.read_excel(lookDir + 'Master Account List.xlsx',
                                     'Allacct').fillna('')
        except XLRDError:
            print('---\n'
                  'Error reading sheet name for Master Account List.xlsx!\n'
                  'Please make sure the main tab is named Allacct.\n'
                  '***')
            return
        # Check the column names.
        mastCols = ['ProperName', 'SLS', 'CITY']
        missCols = [i for i in mastCols if i not in list(mastAcct)]
        if missCols:
            print('The following columns were not detected in '
                  'Master Account List.xlsx:\n%s' %
                  ', '.join(map(str, missCols))
                  + '\nRemember to delete lines before the column '
                  'headers.\n***')
            return
    else:
        print('---\n'
              'No Master Account List file found!\n'
              'Please make sure the Master Account List '
              'is in the directory.\n'
              '***')
        return

    print('Looking up salespeople...')

    # Strip the root off of the filepath and leave just the filename.
    filename = os.path.basename(filepath)

    # Load the Insight file.
    insFile = pd.read_excel(filepath, None)
    insFile = insFile[list(insFile)[0]].fillna('')

    # -------------------------------------------
    # Clean up and match the new Digikey LI file.
    # -------------------------------------------
    # Switch the datetime objects over to strings.
    for col in list(insFile):
        try:
            insFile[col] = insFile[col].dt.strftime('%Y-%m-%d')
        except AttributeError:
            pass

    # Get the column list and input new columns.
    colNames = list(insFile)
    colNames[4:4] = ['Sales']
    colNames[6:6] = ['Must Contact', 'End Product', 'How Contacted',
                     'Information for Digikey']
    colNames[19:19] = ['Invoiced Dollars']
    colNames[25:25] = ['City on Acct List']
    colNames.extend(['TAARCOM Comments', 'New T-Cust'])

    # Calculate the Invoiced Dollars.
    try:
        qty = pd.to_numeric(insFile['Qty Shipped'], errors='coerce')
        price = pd.to_numeric(insFile['Unit Price'], errors='coerce')
        insFile['Invoiced Dollars'] = qty*price
        insFile['Invoiced Dollars'].fillna('', inplace=True)
    except KeyError:
        print('Error calculating Invoiced Dollars.\n'
              'Please make sure Qty Shipped and Unit Price columns '
              'are in the report.\n'
              '(Also check that the top line of the file contains '
              'the column names).\n'
              '***')
        return

    # Remove the 'Send' column, if present.
    try:
        colNames.remove('Send')
    except ValueError:
        pass

    if 'Root Customer..' not in colNames:
        print('Did not find a column named "Root Customer.."\n'
              'Please make sure this column exists and try again.\n'
              'Note: also check that row 1 of the file is the column headers.'
              '\n***')
        return

    # Go through each entry in the Insight file and look for a sales match.
    for row in insFile.index:
        # Check for individuals and CMs and note them in comments.
        if 'contract' in insFile.loc[row, 'Root Customer Class'].lower():
            insFile.loc[row, 'TAARCOM Comments'] = 'Contract Manufacturer'
        if 'individual' in insFile.loc[row, 'Root Customer Class'].lower():
            insFile.loc[row, 'TAARCOM Comments'] = 'Individual'
            # Assocaited cities to each salesperson.
            salesByCity = {'CM': ['Santa Francisco'],
                           'CR': ['San Jose', 'Morgan Hill'],
                           'DC': ['Oakland', 'Union City', 'Hayward',
                                  'Alameda', 'Berkeley', 'Petaluma',
                                  'Healdsburg', 'Santa Rosa', 'Rohnert Park'],
                           'HS': ['Sacramento', 'Auburn', 'Grass Valley',
                                  'Carson City', 'Reno'],
                           'JW': ['Sunnyvale', 'Moffett Field', 'Campbell',
                                  'Saratoga', 'Los Gatos', 'Cupertino',
                                  'Scotts Valley', 'Santa Cruz'],
                           'MG': ['Palo Alto', 'San Mateo', 'Belmont',
                                  'Stanford', 'San Carlos', 'Redwood City',
                                  'Menlo Park'],
                           'MM': ['Mountain View', 'Los Altos', 'Fremont',
                                  'Milpitas', 'San Francisco']}
            # Assign salesperson based on city.
            for key in salesByCity.keys():
                city = insFile.loc[row, 'CITY'].upper()
                if city in map(lambda x: x.upper(), salesByCity[key]):
                    insFile.loc[row, 'Sales'] = key
            # Done, so move to next line in file.
            continue
        cust = insFile.loc[row, 'Root Customer..']
        # Check for customer match in account list.
        acctMatch = mastAcct[mastAcct['ProperName'] == cust]
        if cust and len(acctMatch) == 1:
            # Check if the city is different from our account list.
            acctCity = acctMatch['CITY'].iloc[0].upper().split(", ")
            if insFile.loc[row, 'Customer City'].upper() not in acctCity:
                if len(acctCity) > 1:
                    acctCity = ', '.join(acctCity)
                insFile.loc[row, 'City on Acct List'] = acctCity
            # Copy over salesperson.
            insFile.loc[row, 'Sales'] = acctMatch['SLS'].iloc[0]
        else:
            # Look for match in rootCustMap file.
            salesMatch = rootCustMap['Root Customer'] == cust
            match = rootCustMap[salesMatch]
            if cust and len(match) == 1:
                # Match to salesperson if exactly one match is found.
                insFile.loc[row, 'Sales'] = match['Salesperson'].iloc[0]
            else:
                # Record that the customer is new.
                insFile.loc[row, 'New T-Cust'] = 'Y'
        # Convert applicable entries to numeric.
        for col in list(insFile):
            insFile.loc[row, col] = pd.to_numeric(insFile.loc[row, col],
                                                  errors='ignore')

    # Reorder columns and fill NaNs.
    insFile = insFile.loc[:, colNames].fillna('')

    # Try saving the files, exit with error if any file is currently open.
    outDir = 'C:/Users/kerry/Documents/disty data/Digikey/'
    fname1 = outDir + filename[:-5] + ' With Salespeople.xlsx'
    if saveError(fname1):
        print('---\n'
              'One or more files are currently open in Excel!\n'
              'Please close the files and try again.\n'
              '***')
        return

    # Write the Digikey Insight file, which now contains salespeople.
    writer1 = pd.ExcelWriter(fname1, engine='xlsxwriter',
                             datetime_format='mm/dd/yyyy')
    insFile.to_excel(writer1, sheet_name='Data', index=False)
    # Format in Excel.
    tableFormat(insFile, 'Data', writer1)

    # Save the files.
    writer1.save()

    print('---\n'
          'Salespeople successfully looked up!\n'
          'New file saved as in:\n ' + fname1 + '\n+++')
