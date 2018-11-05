import pandas as pd
from dateutil.parser import parse
from xlrd import XLRDError
import time
import calendar
import math
import os.path
import re


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
            # Highlight any estimates in Invoiced Dollars.
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
        maxWidth = max(len(str(val)) for val in sheetData[col].values)
        maxWidth = min(maxWidth, 50)
        sheet.set_column(i, i, maxWidth+0.8, formatting)
        i += 1


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


def tailoredPreCalc(princ, sheet, sheetName):
    """Do special pre-processing tailored to the principal input."""
    # Osram special processing.
    if princ == 'OSR':
        # Get rid of the Item column.
        try:
            sheet.rename(columns={'Item': 'Unmapped',
                                  'Material Number': 'Unmapped 2'},
                         inplace=True)
        except KeyError:
            pass
        # Combine Rep 1 % and Rep 2 %.
        if 'Rep 1 %' in list(sheet) and 'Rep 2 %' in list(sheet):
            for row in sheet.index:
                if sheet.loc[row, 'Rep 2 %']:
                    if not sheet.loc[row, 'Rep 1 %']:
                        sheet.loc[row, 'Rep 1 %'] = sheet.loc[row, 'Rep 2 %']
                    else:
                        print('Conflict between Rep 1 % and Rep 2 % columns.\n'
                              'Make sure each line has an entry in only '
                              'one of the two columns.\n'
                              '***')
                        return
    # ISSI special processing.
    if princ == 'ISS':
        try:
            # Rename the 'Commissions Due' column.
            sheet.rename(columns={'Commission Due': 'Unmapped',
                                  'Name': 'OEM/POS'}, inplace=True)
            print('Ignoring the Commissions Due column.')
        except KeyError:
            pass
    # ATS special processing.
    if princ == 'ATS':
        try:
            # Rename the 'Commissions Due' column.
            sheet.rename(columns={'Resale': 'Extended Resale',
                                  'Cost': 'Extended Cost'}, inplace=True)
            print('Matching Paid-On Revenue to the distributor.')
        except KeyError:
            pass
    # QRF special processing.
    if princ == 'QRF':
        if sheetName == 'OEM':
            # The column Name 11 needs to be deleted.
            try:
                sheet.rename(columns={'Name 11': 'Unmapped',
                                      'End Customer': 'Unmapped 2'},
                             inplace=True)
            except KeyError:
                pass
        elif sheetName == 'POS':
            # The column Customer is actually the Distributor.
            try:
                sheet.rename(columns={'Customer': 'Distributor',
                                      'BillDocNo': 'Unmapped',
                                      'End Customer': 'Unmapped 2'},
                             inplace=True)
            except KeyError:
                pass
    # INF special processing.
    if princ == 'INF':
        if 'Rep Group' in list(sheet):
            # Material Number is bad on this sheet.
            try:
                sheet.rename(columns={'Material Number': 'Unmapped'},
                             inplace=True)
            except KeyError:
                pass
        else:
            # A bunch of things are bad on this sheet.
            try:
                sheet.rename(columns={'Material Description': 'Unmapped1',
                                      'Sold To Name': 'Unmapped2',
                                      'Ship To Name': 'Unmapped3',
                                      'Item': 'Unmapped4',
                                      'End Name': 'Customer Name'},
                             inplace=True)
            except KeyError:
                pass


def tailoredCalc(princ, sheet, sheetName, distMap):
    """Do special processing tailored to the principal input."""
    # Make sure applicable entries exist and are numeric.
    invDol = True
    extCost = True
    try:
        sheet['Invoiced Dollars'] = pd.to_numeric(sheet['Invoiced Dollars'],
                                                  errors='coerce').fillna(0)
    except KeyError:
        invDol = False
    try:
        sheet['Ext. Cost'] = pd.to_numeric(sheet['Ext. Cost'],
                                           errors='coerce').fillna(0)
    except KeyError:
        extCost = False

    # Abracon special processing.
    if princ == 'ABR':
        invIn = 'Invoiced Dollars' in list(sheet)
        commNotIn = 'Actual Comm Paid' not in list(sheet)
        revIn = 'Paid-On Revenue' in list(sheet)
        commRateNotIn = 'Commission Rate' not in list(sheet)
        if invIn and commNotIn:
            # Input missing data. Commission Rate is always 3% here.
            sheet['Commission Rate'] = 0.03
            sheet['Paid-On Revenue'] = pd.to_numeric(sheet['Invoiced Dollars'],
                                                     errors='coerce')*0.7
            sheet['Actual Comm Paid'] = sheet['Paid-On Revenue']*0.03
            print('Columns added from Abracon special processing:\n'
                  'Commission Rate, Paid-On Revenue, '
                  'Actual Comm Paid\n'
                  '---')
        elif revIn and commRateNotIn:
            # Fill down Distributor for their grouping scheme.
            sheet['Distributor'].fillna(method='ffill', inplace=True)
            sheet['Distributor'].fillna('', inplace=True)
            # Calculate the Commission Rate.
            comPaid = pd.to_numeric(sheet['Actual Comm Paid'], errors='coerce')
            revenue = pd.to_numeric(sheet['Paid-On Revenue'], errors='coerce')
            comRate = round(comPaid/revenue, 3)
            sheet['Commission Rate'] = comRate
            print('Columns added from Abracon special processing:\n'
                  'Commission Rate\n'
                  '---')
    # ISSI special processing.
    if princ == 'ISS':
        if 'OEM/POS' in list(sheet):
            for row in range(len(sheet)):
                # Deal with OEM idiosyncrasies.
                if 'OEM' in sheet.loc[row, 'OEM/POS']:
                    # Put Sales Region into City.
                    sheet.loc[row, 'City'] = sheet.loc[row, 'Sales Region']
                    # Check for distributor in Customer
                    cust = sheet.loc[row, 'Reported Customer']
                    distName = re.sub('[^a-zA-Z0-9]', '', str(cust).lower())
                    # Find matches in the Distributor Abbreviations.
                    distMatches = [i for i in distMap['Search Abbreviation']
                                   if i in distName]
                    if len(distMatches) == 1:
                        # Copy to distributor column.
                        try:
                            sheet.loc[row, 'Distributor'] = cust
                        except KeyError:
                            pass
    # ATS special Processing.
    if princ == 'ATS':
        # Digikey and Mouser are paid on cost, not resale.
        for row in sheet.index:
            try:
                if sheet.loc[row, 'Distributor'] in ['DIGIKEY', 'MOUSER']:
                    sheet.loc[row, 'Paid-On Revenue'] = sheet.loc[row, 'Ext. Cost']
                else:
                    sheet.loc[row, 'Paid-On Revenue'] = sheet.loc[row, 'Invoiced Dollars']
            except KeyError:
                pass
    # ATP special Processing.
    if princ == 'ATP':
        # Load up the customer lookup file.
        if os.path.exists('ATPCustomerLookup.xlsx'):
            ATPCustLook = pd.read_excel('ATPCustomerLookup.xlsx',
                                        'Lookup').fillna('')
        else:
            print('No ATP Customer Lookup found!\n'
                  'Please make sure the Customer Lookup is in the directory.\n'
                  'Skipping tab.\n'
                  '---')
            return
        # Fill in commission rates and commission paid.
        if 'US' in sheetName and invDol:
            sheet['Commission Rate'] = 0.05
            sheet['Actual Comm Paid'] = sheet['Invoiced Dollars']*0.05
            print('Commission rate filled in for this tab: 5%\n'
                  '---')
            sheet['Reported Customer'].fillna(method='ffill', inplace=True)
            print('Correcting customer names.\n'
                  '---')
        elif 'TW' in sheetName and invDol:
            sheet['Commission Rate'] = 0.04
            sheet['Actual Comm Paid'] = sheet['Invoiced Dollars']*0.04
            print('Commission rate filled in for this tab: 4%\n'
                  '---')
            sheet['Reported Customer'].fillna(method='ffill', inplace=True)
            print('Correcting customer names.\n'
                  '---')
        elif 'POS' in sheetName and extCost:
            sheet['Commission Rate'] = 0.03
            sheet.rename(columns={'Reported End Customer':
                                  'Reported Customer'}, inplace=True)
            sheet['Actual Comm Paid'] = sheet['Ext. Cost']*0.03
            # Estimate invoiced dollars as 15% markup on cost.
            sheet['Invoiced Dollars'] = sheet['Ext. Cost']*1.15
            print('Commission rate filled in for this tab: 3%\n'
                  'Invoiced dollars estimated as 15% over Ext. Cost '
                  'for this tab.\n'
                  'ESTIMATED ENTRIES HIGHLIGHTED YELLOW.\n'
                  '---')
        else:
            print('-\n'
                  'Tab not labeled as US/TW/POS, '
                  'or Ext. Cost/Invoiced Dollars not found on this tab.\n'
                  '---')
        # Correct the customer names.
        try:
            sheet['Reported Customer'].fillna('', inplace=True)
        except KeyError:
            print('No Reported Customer column found!\n'
                  '---')
            return
        for row in range(len(sheet)):
            custName = sheet.loc[row, 'Reported Customer']
            # Find matches for the distName in the Distributor Abbreviations.
            custMatches = [i for i in ATPCustLook['Name'] if i in custName]
            if len(custMatches) == 1:
                # Find and input corrected distributor name.
                mloc = ATPCustLook['Name'] == custMatches[0]
                corrCust = ATPCustLook[mloc].iloc[0]['Corrected']
                sheet.loc[row, 'Reported Customer'] = corrCust
    # Mill-Max special Processing.
    if princ == 'MIL':
        invNum = True
        try:
            sheet['Invoice Number']
        except KeyError:
            print('Found no Invoice Numbers on this sheet.\n'
                  '---')
            invNum = False
        if 'Ext. Cost' in list(sheet) and not invDol:
            # Estimate invoiced dollars as 15% markup on cost.
            sheet['Invoiced Dollars'] = sheet['Ext. Cost']*1.15
            print('Invoiced dollars estimated as 15% over Ext. Cost '
                  'for this tab.\n'
                  'ESTIMATED ENTRIES HIGHLIGHTED YELLOW.\n'
                  '---')
            # Sometimes the Totals are written in the Part Number column.
            sheet = sheet[sheet['Part Number'] != 'Totals']
            sheet.reset_index(drop=True, inplace=True)
        elif 'Part Number' not in list(sheet) and invNum:
            # We need to load in the part number log.
            if os.path.exists('Mill-Max Invoice Log.xlsx'):
                MMaxLog = pd.read_excel('Mill-Max Invoice Log.xlsx',
                                        'Logs').fillna('')
                print('Looking up part numbers from invoice log.\n'
                      '---')
            else:
                print('No Mill-Max Invoice Log found!\n'
                      'Please make sure the Invoice Log is in the directory.\n'
                      'Skipping tab.\n'
                      '---')
                return
            # Input part number from Mill-Max Invoice Log.
            for row in range(len(sheet)):
                match = MMaxLog['Inv#'] == sheet.loc[row, 'Invoice Number']
                if sum(match) == 1:
                    partNum = MMaxLog[match].iloc[0]['Part Number']
                    sheet.loc[row, 'Part Number'] = partNum
                else:
                    print('Part number match error (invoice number either '
                          'not found or duplicated)!\n'
                          '---')
    # Osram special Processing.
    if princ == 'OSR':
        # For World Star POS tab, enter World Star as the distributor.
        if 'World' in sheetName:
            sheet['Distributor'] = 'World Star'
    # Cosel special Processing.
    if princ == 'COS':
        # Only work with the Details tab.
        if sheetName == 'Details' and extCost:
            print('Calculating commissions as 5% of Cost Ext.\n'
                  'Subtracting 2% off of commissions for Allied shipments.\n'
                  '---')
            for row in sheet.index:
                extenCost = sheet.loc[row, 'Ext. Cost']
                if sheet.loc[row, 'Distributor'] == 'ALLIED':
                    sheet.loc[row, 'Commission Rate'] = 0.049
                    sheet.loc[row, 'Actual Comm Paid'] = 0.049*extenCost
                else:
                    sheet.loc[row, 'Commission Rate'] = 0.05
                    sheet.loc[row, 'Actual Comm Paid'] = 0.05*extenCost
    # Globtek special Processing.
    if princ == 'GLO':
        # Make sure we have Invoiced Dollars.
        if invDol:
            sheet['Commission Rate'] = 0.05
            sheet['Actual Comm Paid'] = sheet['Invoiced Dollars']*0.05

    # Test the Commission Dollars to make sure they're correct.
    if 'Paid-On Revenue' in list(sheet):
        paidDol = sheet['Paid-On Revenue']
        if 'Shared Rev Tier Rate' in list(sheet):
            paidDol = paidDol*sheet['Shared Rev Tier Rate']
    elif not invDol and extCost:
        paidDol = sheet['Ext. Cost']
    elif invDol:
        paidDol = sheet['Invoiced Dollars']
    try:
        # Find percent error in Commission Dollars.
        realCom = paidDol*sheet['Commission Rate']
        realCom = realCom[realCom > 1]
        # Find commissions paid.
        if princ == 'ISS':
            # ISSI holds back 10% in Actual Comm Due.
            paidCom = sheet['Commission Due']
        else:
            paidCom = sheet['Actual Comm Paid']
        absError = abs(realCom - paidCom)
        if any(absError > 1):
            print('Greater than $1 error in Commission Dollars detected '
                  'in the following rows:')
            errLocs = absError[absError > 1].index.tolist()
            errLocs = [i+2 for i in errLocs]
            print(', '.join(map(str, errLocs))
                  + '\n---')
    except (KeyError, NameError):
        pass


# %%
# The main function.
def main(filepaths, runningCom, fieldMappings, inPrinc):
    """Processes commission files and appends them to Running Commissions.

    Columns in individual commission files are identified and appended to the
    Running Commissions under the appropriate column, as identified by the
    fieldMappings file. Entries are then passed through the Lookup Master in
    search of a match to Reported Customer + Part Number. Distributors are
    corrected to consistent names. Entries with missing information are copied
    to Entries Need Fixing for further attention.

    Arguments:
    filepaths -- paths for opening (Excel) files to process.
    runningCom -- current Running Commissions file (in Excel) to
                  which we are appending data.
    fieldMappings -- dataframe which links Running Commissions columns to
                     file data columns.
    principal -- the principal that supplied the commission file(s). Chosen
                 from the dropdown menu on the GUI main window.
    """
    # Grab lookup table data names.
    columnNames = list(fieldMappings)
    # Add in non-lookup'd data names.
    columnNames[0:0] = ['CM Sales', 'Design Sales', 'Quarter Shipped', 'Month',
                        'Year']
    columnNames[7:7] = ['T-End Cust', 'T-Name', 'CM',
                        'Principal', 'Corrected Distributor']
    columnNames[26:26] = ['Sales Commission']
    columnNames.extend(['CM Split', 'TEMP/FINAL', 'Paid Date', 'From File',
                        'Sales Report Date'])

    # Check to see if there's an existing Running Commissions to append to.
    if runningCom:
        finalData = pd.read_excel(runningCom, 'Master').fillna('')
        runComLen = len(finalData)
        filesProcessed = pd.read_excel(runningCom,
                                       'Files Processed').fillna('')
        print('Appending files to Running Commissions.')
        if list(finalData) != columnNames:
            print('---\n'
                  'Columns in Running Commissions '
                  'do not match fieldMappings.xlsx!\n'
                  'Please check column names and try again.\n'
                  '***')
            return
    # Start new Running Commissions.
    else:
        print('No Running Commissions file provided. Starting a new one.')
        runComLen = 0
        # These are our names for the data in Running Commissions.
        finalData = pd.DataFrame(columns=columnNames)
        filesProcessed = pd.DataFrame(columns=['Filename',
                                               'Total Commissions',
                                               'Date Added',
                                               'Paid Date'])

    # Strip the root off of the filepaths and leave just the filenames.
    filenames = [os.path.basename(val) for val in filepaths]
    # Check if we've duplicated any files.
    duplicates = list(set(filenames).intersection(filesProcessed['Filename']))
    # Don't let duplicate files get processed.
    filenames = [val for val in filenames if val not in duplicates]
    if duplicates:
        # Let us know we found duplictes and removed them.
        print('---\n'
              'The following files are already in Running Commissions:')
        for file in list(duplicates):
            print(file)
        print('Duplicate files were removed from processing.')
        # Exit if no new files are left.
        if not filenames:
            print('---\n'
                  'No new commissions files selected.\n'
                  'Please try selecting files again.\n'
                  '***')
            return

    # Read in each new file with Pandas and store them as dictionaries.
    # Each dictionary has a dataframe for each sheet in the file.
    inputData = [pd.read_excel(filepath, None) for filepath in filepaths]

    # Read in distMap. Exit if not found or if errors in file.
    if os.path.exists('distributorLookup.xlsx'):
        try:
            distMap = pd.read_excel('distributorLookup.xlsx', 'Distributors')
        except XLRDError:
            print('---\n'
                  'Error reading sheet name for distributorLookup.xlsx!\n'
                  'Please make sure the main tab is named Distributors.\n'
                  '***')
            return
        # Check the column names.
        distMapCols = ['Corrected Dist', 'Search Abbreviation']
        missCols = [i for i in distMapCols if i not in list(distMap)]
        if missCols:
            print('The following columns were not detected in '
                  'distributorLookup.xlsx:\n%s' %
                  ', '.join(map(str, missCols))
                  + '\n***')
            return
    else:
        print('---\n'
              'No distributor lookup file found!\n'
              'Please make sure distributorLookup.xlsx is in the directory.\n'
              '***')
        return

    # Read in file of entries that need fixing. Exit if not found.
    if os.path.exists('Entries Need Fixing.xlsx'):
        try:
            fixList = pd.read_excel('Entries Need Fixing.xlsx',
                                    'Data').fillna('')
        except XLRDError:
            print('---\n'
                  'Error reading sheet name for Entries Need Fixing.xlsx!\n'
                  'Please make sure the main tab is named Data.\n'
                  '***')
            return
    else:
        print('---\n'
              'No Entries Need Fixing file found!\n'
              'Please make sure Entries Need Fixing.xlsx'
              'is in the directory.\n'
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
              'Please make sure lookupMaster.xlsx is in the directory.\n'
              '***')
        return

    # %%
    # Iterate through each file that we're appending to Running Commissions.
    fileNum = 0
    for filename in filenames:
        # Grab the next file from the list.
        newData = inputData[fileNum]
        fileNum += 1
        print('______________________________________________________\n'
              'Working on file: ' + filename
              + '\n______________________________________________________')
        # Set total commissions for file back to zero.
        totalComm = 0

        # If principal is auto-detected, find it from filename.
        if inPrinc == 'AUTO-DETECT':
            principal = filename[0:3]
            print('Principal detected as: ' + principal)
        else:
            principal = inPrinc

        # Iterate over each dataframe in the ordered dictionary.
        # Each sheet in the file is its own dataframe in the dictionary.
        for sheetName in list(newData):
            # Grab next sheet in file.
            # Rework the index just in case it got read in wrong.
            sheet = newData[sheetName].reset_index(drop=True)
            # Make sure index is an integer, not a string.
            sheet.index = sheet.index.map(int)
            # Strip whitespace from column names.
            sheet.rename(columns=lambda x: str(x).strip(), inplace=True)
            # Clear out unnamed columns. Attribute error means it's an empty
            # sheet, so simply pass it along (it'll get dealt with).
            try:
                sheet = sheet.loc[:, ~sheet.columns.str.contains('^Unnamed')]
            except AttributeError:
                pass
            totalRows = sheet.shape[0]
            print('Found ' + str(totalRows) + ' entries in the tab '
                  + sheetName + '\n----------------------------------')
            # Do specialized pre-processing tailored to principlal.
            tailoredPreCalc(principal, sheet, sheetName)

            # Iterate over each column of data that we want to append.
            missingCols = []
            for dataName in list(fieldMappings):
                # Grab list of names that the data could potentially be under.
                nameList = fieldMappings[dataName].dropna().tolist()
                # Look for a match in the sheet column names.
                sheetColumns = list(sheet)
                columnName = [val for val in sheetColumns if val in nameList]

                # Let us know if we didn't find a column that matches,
                # or if we found too many columns that match,
                # then rename the column in the sheet to the master name.
                if not columnName:
                    missingCols.append(dataName)
                elif len(columnName) > 1:
                    print('Found multiple matches for ' + dataName
                          + '\nMatching columns: %s' %
                          ', '.join(map(str, columnName))
                          + '\nPlease fix column names and try again.\n'
                          '***')
                    return
                else:
                    sheet.rename(columns={columnName[0]: dataName},
                                 inplace=True)

            # Fix Commission Rate if it got read in as a decimal.
            try:
                numCom = pd.to_numeric(sheet['Commission Rate'],
                                       errors='coerce')
                sheet['Commission Rate'] = numCom
                decCom = sheet['Commission Rate'] > 0.9
                newCom = sheet.loc[decCom, 'Commission Rate']/100
                sheet.loc[decCom, 'Commission Rate'] = newCom.fillna('')
            except (KeyError, TypeError):
                pass
            # Fix Split Percentage if it got read in as a decimal.
            try:
                # Remove '%' sign if present.
                numSplit = sheet['Split Percentage'].astype(str).map(
                                lambda x: x.strip('%'))
                # Conver to numeric.
                numSplit = pd.to_numeric(numSplit, errors='coerce')
                sheet['Split Percentage'] = numSplit
                decSplit = sheet['Split Percentage'] > 1
                newSplit = sheet.loc[decSplit, 'Split Percentage']/100
                sheet.loc[decSplit, 'Split Percentage'] = newSplit.fillna('')
            except (KeyError, TypeError):
                pass
            # Fix Gross Rev Reduction if it got read in as a decimal.
            try:
                numRev = pd.to_numeric(sheet['Gross Rev Reduction'],
                                       errors='coerce')
                sheet['Gross Rev Reduction'] = numRev
                decRev = sheet['Gross Rev Reduction'] > 1
                newRev = sheet.loc[decRev, 'Gross Rev Reduction']/100
                sheet.loc[decRev, 'Gross Rev Reduction'] = newRev.fillna('')
            except (KeyError, TypeError):
                pass
            # Fix Shared Rev Tier Rate if it got read in as a decimal.
            try:
                numTier = pd.to_numeric(sheet['Shared Rev Tier Rate'],
                                        errors='coerce')
                sheet['Shared Rev Tier Rate'] = numTier
                decTier = sheet['Shared Rev Tier Rate'] > 1
                newTier = sheet.loc[decTier, 'Shared Rev Tier Rate']/100
                sheet.loc[decTier, 'Shared Rev Tier Rate'] = newTier.fillna('')
            except (KeyError, TypeError):
                pass
            # Do special processing for principal, if applicable.
            tailoredCalc(principal, sheet, sheetName, distMap)
            # Drop entries with emtpy part number.
            try:
                sheet.dropna(subset=['Part Number'], inplace=True)
            except KeyError:
                pass

            # Now that we've renamed all of the relevant columns,
            # append the new sheet to Running Commissions, where only the
            # properly named columns are appended.
            if sheet.columns.duplicated().any():
                dupes = sheet.columns[sheet.columns.duplicated()].unique()
                print('Two items are being mapped to the same column!\n'
                      'These columns contain duplicates: %s' %
                      ', '.join(map(str, dupes))
                      + '\n***')
                return
            elif 'Actual Comm Paid' not in list(sheet):
                # Tab has no commission data, so it is ignored.
                print('No commission dollars found on this tab.\n'
                      'Skipping tab.\n'
                      '-')
            elif 'Part Number' not in list(sheet):
                # Tab has no paart number data, so it is ignored.
                print('No part numbers found on this tab.\n'
                      'Skipping tab.\n'
                      '-')
            elif 'Invoice Date' not in list(sheet):
                # Tab has no date column, so report and exit.
                print('No Invoice Date column found for this tab.\n'
                      'Please make sure the Invoice Date is mapped.\n'
                      '***')
                return
            else:
                # Remove entries with no commissions dollars.
                # Coerce entries with bad data (non-numeric gets 0).
                sheet['Actual Comm Paid'] = pd.to_numeric(
                        sheet['Actual Comm Paid'],
                        errors='coerce').fillna(0)
                sheet = sheet[sheet['Actual Comm Paid'] != 0]

                # Add 'From File' column to track where data came from.
                sheet['From File'] = filename
                # Fill in principal.
                sheet['Principal'] = principal

                # Find matching columns.
                matchingColumns = [val for val in list(sheet)
                                   if val in list(fieldMappings)]
                if len(matchingColumns) > 0:
                    # Sum commissions paid on sheet.
                    print('Commissions for this tab: '
                          + '${:,.2f}'.format(sheet['Actual Comm Paid'].sum())
                          + '\n-')
                    totalComm += sheet['Actual Comm Paid'].sum()
                    # Strip whitespace from all strings in dataframe.
                    stringCols = [val for val in list(sheet)
                                  if sheet[val].dtype == 'object']
                    for col in stringCols:
                        sheet[col] = sheet[col].fillna('').astype(str).map(
                                lambda x: x.strip())

                    # Append matching columns of data.
                    appCols = matchingColumns + ['From File', 'Principal']
                    finalData = finalData.append(sheet[appCols],
                                                 ignore_index=True)
                    # Let us know which columns are missing.
                    if missingCols:
                        print('The following columns were not found: %s' %
                              ', '.join(map(str, missingCols))
                              + '\n-')
                else:
                    print('Found no data on this tab. Moving on.\n'
                          '-')

        if totalComm > 0:
            # Show total commissions.
            print('Total commissions for this file: '
                  '${:,.2f}'.format(totalComm))
            # Append filename and total commissions to Files Processed sheet.
            newFile = pd.DataFrame({'Filename': [filename],
                                    'Total Commissions': [totalComm],
                                    'Date Added': [time.strftime('%m/%d/%Y')],
                                    'Paid Date': ['']})
            filesProcessed = filesProcessed.append(newFile, ignore_index=True)
        else:
            print('No new data found in this file.\n'
                  'Moving on without adding file.')

    # %%
    # Fill NaNs left over from appending.
    finalData.fillna('', inplace=True)
    # Find matches in Lookup Master and extract data from them.
    # Let us know how many rows are being processed.
    numRows = '{:,.0f}'.format(len(finalData) - runComLen)
    if numRows == '0':
        print('---\n'
              'No new valid data provided.\n'
              'Please check the new files for missing '
              'data or column matches.\n'
              '***')
        return
    print('---\n'
          'Beginning processing on ' + numRows + ' rows of data.')
    finalData.reset_index(inplace=True, drop=True)

    # Iterate over each row of the newly appended data.
    for row in range(runComLen, len(finalData)):
        # Fill in the Sales Commission.
        salesComm = finalData.loc[row, 'Actual Comm Paid']*0.45
        finalData.loc[row, 'Sales Commission'] = salesComm
        # First match part number.
        partNum = str(finalData.loc[row, 'Part Number']).lower()
        PPN = masterLookup['Part Number'].map(lambda x: str(x).lower())
        partNoMatches = masterLookup[partNum == PPN]
        # Next match reported customer.
        repCust = str(finalData.loc[row, 'Reported Customer']).lower()
        POSCust = partNoMatches['Reported Customer'].map(lambda x: str(x).lower())
        custMatches = partNoMatches[repCust == POSCust].reset_index()
        # Record number of Lookup Master matches.
        lookMatches = len(custMatches)
        # Make sure we found exactly one match.
        if lookMatches == 1:
            custMatches = custMatches.iloc[0]
            # Grab primary and secondary sales people from Lookup Master.
            finalData.loc[row, 'CM Sales'] = custMatches['CM Sales']
            finalData.loc[row, 'Design Sales'] = custMatches['Design Sales']
            finalData.loc[row, 'T-Name'] = custMatches['T-Name']
            finalData.loc[row, 'CM'] = custMatches['CM']
            finalData.loc[row, 'T-End Cust'] = custMatches['T-End Cust']
            finalData.loc[row, 'CM Split'] = custMatches['CM Split']
            # Update usage in lookup Master.
            masterLookup.loc[custMatches['index'],
                             'Last Used'] = time.strftime('%m/%d/%Y')
            # Update OOT city if not already filled in.
            if custMatches['T-Name'][0:3] == 'OOT' and not custMatches['City']:
                masterLookup.loc[custMatches['index'],
                                 'City'] = finalData.loc[row, 'City']

        # Try parsing the date.
        dateError = False
        dateGiven = finalData.loc[row, 'Invoice Date']
        # Check if the date is read in as a float/int, and convert to string.
        if isinstance(finalData.loc[row, 'Invoice Date'], (float, int)):
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
            print('---'
                  'There is no Invoice Date column in Running Commissions!\n'
                  'Please check to make sure an Invoice Date column exists.\n'
                  'Note: Spelling, whitespace, and capitalization matter.\n'
                  '---')
            dateError = True
        # If no error found in date, fill in the month/year/quarter
        if not dateError:
            date = parse(dateGiven)
            # Make sure the date actually makes sense.
            currentYear = int(time.strftime('%Y'))
            if currentYear - date.year not in [0, 1]:
                dateError = True
            else:
                # Cast date format into mm/dd/yyyy.
                finalData.loc[row, 'Invoice Date'] = date.strftime('%m/%d/%Y')
                # Fill in quarter/year/month data.
                finalData.loc[row, 'Year'] = date.year
                finalData.loc[row, 'Month'] = calendar.month_name[date.month][0:3]
                finalData.loc[row, 'Quarter Shipped'] = (str(date.year) + 'Q'
                                                         + str(math.ceil(date.month/3)))

        # Find a corrected distributor match.
        # Strip extraneous characters and all spaces, and make lowercase.
        distName = re.sub('[^a-zA-Z0-9]', '',
                          str(finalData.loc[row, 'Distributor'])).lower()

        # Find matches for the distName in the Distributor Abbreviations.
        distMatches = [i for i in distMap['Search Abbreviation']
                       if i in distName]
        if len(distMatches) == 1:
            # Find and input corrected distributor name.
            mloc = distMap['Search Abbreviation'] == distMatches[0]
            corrDist = distMap[mloc].iloc[0]['Corrected Dist']
            finalData.loc[row, 'Corrected Distributor'] = corrDist
        elif not distName:
            finalData.loc[row, 'Corrected Distributor'] = ''
            distMatches = ['Empty']

        # Go through each column and convert applicable entries to numeric.
        cols = list(finalData)
        cols.remove('Principal')
        for col in cols:
            finalData.loc[row, col] = pd.to_numeric(finalData.loc[row, col],
                                                    errors='ignore')

        # If any data isn't found/parsed, copy entry to Fix Entries.
        if lookMatches != 1 or len(distMatches) != 1 or dateError:
            fixList = fixList.append(finalData.loc[row, :])
            fixList.loc[row, 'Running Com Index'] = row
            fixList.loc[row, 'Distributor Matches'] = len(distMatches)
            fixList.loc[row, 'Lookup Master Matches'] = lookMatches
            fixList.loc[row, 'Date Added'] = time.strftime('%m/%d/%Y')
            finalData.loc[row, 'TEMP/FINAL'] = 'TEMP'
        else:
            # Everything found, so entry is final.
            finalData.loc[row, 'TEMP/FINAL'] = 'FINAL'

        # Update progress every 1,000 rows.
        if row % 1000 == 0 and row > 0:
            print('Done with row ' '{:,.0f}'.format(row))

    # Reorder columns to match the desired layout in columnNames.
    finalData.fillna('', inplace=True)
    finalData = finalData.loc[:, columnNames]
    columnNames.extend(['Distributor Matches', 'Lookup Master Matches',
                        'Date Added', 'Running Com Index'])
    # Fix up the Entries Need Fixing file.
    fixList = fixList.loc[:, columnNames]
    fixList.index.name = 'Master Index'
    fixList.reset_index(inplace=True)
    fixList.fillna('', inplace=True)

    # %%
    # Check if the files we're going to save are open already.
    fname1 = 'Running Commissions ' + time.strftime('%Y-%m-%d-%H%M') + '.xlsx'
    fname2 = 'Entries Need Fixing.xlsx'
    fname3 = 'Lookup Master ' + time.strftime('%Y-%m-%d-%H%M') + '.xlsx'
    if saveError(fname1, fname2, fname3):
        print('---\n'
              'One or more files are currently open in Excel!\n'
              'Please close the files and try again.\n'
              '***')
        return

    # Write the Running Commissions file.
    writer1 = pd.ExcelWriter(fname1, engine='xlsxwriter')
    finalData.to_excel(writer1, sheet_name='Master', index=False)
    filesProcessed.to_excel(writer1, sheet_name='Files Processed',
                            index=False)
    # Format as table in Excel.
    tableFormat(finalData, 'Master', writer1)
    tableFormat(filesProcessed, 'Files Processed', writer1)

    # Write the Needs Fixing file.
    writer2 = pd.ExcelWriter(fname2, engine='xlsxwriter')
    fixList.to_excel(writer2, sheet_name='Data', index=False)
    # Format as table in Excel.
    tableFormat(fixList, 'Data', writer2)

    # Write the Lookup Master.
    writer3 = pd.ExcelWriter(fname3, engine='xlsxwriter')
    masterLookup.to_excel(writer3, sheet_name='Lookup', index=False)
    # Format as table in Excel.
    tableFormat(masterLookup, 'Lookup', writer3)

    # Save the files.
    writer1.save()
    writer2.save()
    writer3.save()

    print('---\n'
          'Updates completed successfully!\n'
          '---\n'
          'Running Commissions updated.\n'
          'Lookup Master updated.\n'
          'Entries Need Fixing updated.\n'
          '+++')
