import pandas as pd
import numpy as np
from dateutil.parser import parse
from xlrd import XLRDError
import time
import calendar
import math
import os.path
import re
import datetime
from RCExcelTools import tableFormat, saveError, formDate


def tailoredPreCalc(princ, sheet, sheetName):
    """Do special pre-processing tailored to the principal input. Primarily,
    this involves renaming columns that would get looked up incorrectly
    in the Field Mappings.

    This function modifies a dataframe inplace.
    """
    # Initialize the renameDict in case it doesn't get set later.
    renameDict = {}
    # ------------------------------
    # Osram special pre-processing.
    # ------------------------------
    if princ == 'OSR':
        renameDict = {'Item': 'Unmapped', 'Material Number': 'Unmapped 2',
                      'Customer Name': 'Unmapped 3',
                      'Sales Date': 'Unmapped 4'}
        sheet.rename(columns=renameDict, inplace=True)
        # Combine Rep 1 % and Rep 2 %.
        if 'Rep 1 %' in list(sheet) and 'Rep 2 %' in list(sheet):
            print('Copying Rep 2 % into empty Rep 1 % lines.\n'
                  '---')
            for row in sheet.index:
                if sheet.loc[row, 'Rep 2 %'] and not sheet.loc[row, 'Rep 1 %']:
                    sheet.loc[row, 'Rep 1 %'] = sheet.loc[row, 'Rep 2 %']
    # -----------------------------
    # ISSI special pre-processing.
    # -----------------------------
    if princ == 'ISS':
        renameDict = {'Commission Due': 'Unmapped', 'Name': 'OEM/POS'}
        sheet.rename(columns=renameDict, inplace=True)
    # ----------------------------
    # ATS special pre-processing.
    # ----------------------------
    if princ == 'ATS':
        renameDict = {'Resale': 'Extended Resale', 'Cost': 'Extended Cost'}
        sheet.rename(columns=renameDict, inplace=True)
    # ----------------------------
    # QRF special pre-processing.
    # ----------------------------
    if princ == 'QRF':
        if sheetName in ['OEM', 'OFF']:
            renameDict = {'End Customer': 'Unmapped 2', 'Item': 'Unmapped 3'}
            sheet.rename(columns=renameDict, inplace=True)
        elif sheetName == 'POS':
            renameDict = {'Company': 'Distributor', 'BillDocNo': 'Unmapped',
                          'End Customer': 'Unmapped 2', 'Item': 'Unmapped 3'}
            sheet.rename(columns=renameDict, inplace=True)
    # ----------------------------
    # INF special pre-processing.
    # ----------------------------
    if princ == 'INF':
        if 'Rep Group' in list(sheet):
            renameDict = {'Material Number': 'Unmapped'}
            sheet.rename(columns=renameDict, inplace=True)
            # Drop the RunRate row(s) on this sheet.
            try:
                ID = sheet[sheet['Comm Type'] == 'OffShoreRunRate'].index
                sheet.loc[ID, :] = ''
                print('Dropping any lines with Comm Type as OffShoreRunRate.\n'
                      '-')
            except KeyError:
                print('Found no Comm Type column!\n'
                      '-')
        else:
            renameDict = {'Material Description': 'Unmapped1',
                          'Sold To Name': 'Unmapped2',
                          'Ship To Name': 'Unmapped3', 'Item': 'Unmapped4',
                          'End Name': 'Customer Name'}
            sheet.rename(columns=renameDict, inplace=True)
    # ----------------------------
    # XMO special pre-processing.
    # ----------------------------
    if princ == 'XMO':
        renameDict = {'Amount': 'Commission', 'Commission Due': 'Unmapped'}
        sheet.rename(columns=renameDict, inplace=True)
    # Return the renameDict for future use in the matched raw file.
    return renameDict


def tailoredCalc(princ, sheet, sheetName, distMap):
    """Do special processing tailored to the principal input. This involves
    things like filling in commissions source as cost/resale, setting some
    commission rates that aren't specified in the data, etc.

    This function modifies a dataframe inplace.
    """
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
    # ---------------------------
    # Abracon special processing.
    # ---------------------------
    if princ == 'ABR':
        # Use the sheet names to figure out what processing needs to be done.
        if 'Adj' in sheetName:
            # Input missing data. Commission Rate is always 3% here.
            sheet['Commission Rate'] = 0.03
            sheet['Paid-On Revenue'] = pd.to_numeric(sheet['Invoiced Dollars'],
                                                     errors='coerce')*0.7
            sheet['Actual Comm Paid'] = sheet['Paid-On Revenue']*0.03
            # These are paid on resale.
            sheet['Comm Source'] = 'Resale'
            print('Columns added from Abracon special processing:\n'
                  'Commission Rate, Paid-On Revenue, '
                  'Actual Comm Paid\n'
                  '---')
        elif 'MoComm' in sheetName:
            # Fill down Distributor for their grouping scheme.
            sheet['Reported Distributor'].replace('', np.nan, inplace=True)
            sheet['Reported Distributor'].fillna(method='ffill', inplace=True)
            sheet['Reported Distributor'].fillna('', inplace=True)
            # Paid-On Revenue gets Invoiced Dollars.
            sheet['Paid-On Revenue'] = sheet['Invoiced Dollars']
            sheet['Comm Source'] = 'Resale'
            # Calculate the Commission Rate.
            comPaid = pd.to_numeric(sheet['Actual Comm Paid'], errors='coerce')
            revenue = pd.to_numeric(sheet['Paid-On Revenue'], errors='coerce')
            comRate = round(comPaid/revenue, 3)
            sheet['Commission Rate'] = comRate
            print('Columns added from Abracon special processing:\n'
                  'Commission Rate\n'
                  '---')
        else:
            print('Sheet not recognized!\n'
                  'Make sure the tab name contains either MoComm or Adj '
                  'in the name.\n'
                  'Continuing without extra ABR processing.\n'
                  '---')
    # -------------------------
    # ISSI special processing.
    # -------------------------
    if princ == 'ISS':
        if 'OEM/POS' in list(sheet):
            for row in sheet.index:
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
                            sheet.loc[row, 'Reported Distributor'] = cust
                        except KeyError:
                            pass
        # ISSI is paid on resale.
        sheet['Comm Source'] = 'Resale'
    # ------------------------
    # ATS special processing.
    # ------------------------
    if princ == 'ATS':
        # Try setting the Paid-On Revenue as the Invoiced Dollars.
        try:
            sheet['Paid-On Revenue'] = sheet['Invoiced Dollars']
        except KeyError:
            pass
        # Try setting the cost/resale by the distributor.
        try:
            for row in sheet.index:
                dist = str(sheet.loc[row, 'Reported Distributor']).lower()
                # Digikey and Mouser are paid on cost, not resale.
                if 'digi' in dist or 'mous' in dist:
                    sheet.loc[row, 'Comm Source'] = 'Cost'
                else:
                    sheet.loc[row, 'Comm Source'] = 'Resale'
        except KeyError:
            pass
    # ----------------------------
    # Mill-Max special processing.
    # ----------------------------
    if princ == 'MIL':
        invNum = True
        try:
            sheet['Invoice Number']
        except KeyError:
            print('Found no Invoice Numbers on this sheet.\n'
                  '---')
            invNum = False
        if extCost and not invDol:
            # Sometimes the Totals are written in the Part Number column.
            sheet.drop(sheet[sheet['Part Number'] == 'Totals'].index,
                       inplace=True)
            sheet.reset_index(drop=True, inplace=True)
            # These commissions are paid on cost.
            sheet['Paid-On Revenue'] = sheet['Ext. Cost']
            sheet['Comm Source'] = 'Cost'
        elif 'Part Number' not in list(sheet) and invNum:
            # We need to load in the part number log.
            lookDir = 'Z:/Commissions Lookup/'
            if os.path.exists(lookDir + 'Mill-Max Invoice Log.xlsx'):
                MMaxLog = pd.read_excel(lookDir + 'Mill-Max Invoice Log.xlsx',
                                        dtype=str).fillna('')
                print('Looking up part numbers from invoice log.\n'
                      '---')
            else:
                print('No Mill-Max Invoice Log found!\n'
                      'Please make sure the Invoice Log is in the '
                      'Commission Lookup directory.\n'
                      'Skipping tab.\n'
                      '---')
                return
            # Input part number from Mill-Max Invoice Log.
            for row in sheet.index:
                if sheet.loc[row, 'Invoice Number']:
                    match = MMaxLog['Inv#'] == sheet.loc[row, 'Invoice Number']
                    if sum(match) == 1:
                        partNum = MMaxLog[match].iloc[0]['Part Number']
                        sheet.loc[row, 'Part Number'] = partNum
                    else:
                        sheet.loc[row, 'Part Number'] = 'NOT FOUND'
            # These commissions are paid on resale.
            sheet['Comm Source'] = 'Resale'
    # --------------------------
    # Osram special processing.
    # --------------------------
    if princ == 'OSR':
        # For World Star POS tab, enter World Star as the distributor.
        if 'World' in sheetName:
            sheet['Reported Distributor'] = 'World Star'
        try:
            sheet['Paid-On Revenue'] = sheet['Invoiced Dollars']
        except KeyError:
            pass
        # Osram is paid on resale.
        sheet['Comm Source'] = 'Resale'
    # --------------------------
    # Cosel special processing.
    # --------------------------
    if princ == 'COS':
        # Only work with the Details tab.
        if sheetName == 'Details' and extCost:
            print('Calculating commissions as 5% of Cost Ext.\n'
                  'For Allied shipments, 4.9% of Cost Ext.\n'
                  '---')
            # Revenue is from cost.
            sheet['Paid-On Revenue'] = sheet['Ext. Cost']
            for row in sheet.index:
                extenCost = sheet.loc[row, 'Ext. Cost']
                if sheet.loc[row, 'Reported Distributor'] == 'ALLIED':
                    sheet.loc[row, 'Commission Rate'] = 0.049
                    sheet.loc[row, 'Actual Comm Paid'] = 0.049*extenCost
                else:
                    sheet.loc[row, 'Commission Rate'] = 0.05
                    sheet.loc[row, 'Actual Comm Paid'] = 0.05*extenCost
        # Cosel is paid on cost.
        sheet['Comm Source'] = 'Cost'
    # ----------------------------
    # Globtek special processing.
    # ----------------------------
    if princ == 'GLO':
        try:
            sheet['Paid-On Revenue'] = sheet['Invoiced Dollars']
        except KeyError:
            print('No Invoiced Dollars found on this sheet!\n')
        if 'Commission Rate' not in sheet.columns:
            sheet['Commission Rate'] = 0.05
        if 'Actual Comm Paid' not in sheet.columns:
            try:
                sheet['Actual Comm Paid'] = sheet['Paid-On Revenue']*0.05
            except KeyError:
                print('No Paid-On Revenue found, could not calculate '
                      'Actual Comm Paid.\n'
                      '---')
                return
        # Globtek is paid on resale.
        sheet['Comm Source'] = 'Resale'
    # --------------------------
    # RF360 special processing.
    # --------------------------
    if princ == 'QRF':
        try:
            sheet['Paid-On Revenue'] = sheet['Invoiced Dollars']
        except KeyError:
            pass
        # RF360 is paid on resale.
        sheet['Comm Source'] = 'Resale'
    # ------------------------
    # INF special processing.
    # ------------------------
    if princ == 'INF':
        # INF is paid on resale.
        sheet['Comm Source'] = 'Resale'
    # ------------------------
    # LAT special processing.
    # ------------------------
    if princ == 'LAT':
        # LAT is paid on resale.
        sheet['Comm Source'] = 'Resale'
    # ------------------------
    # SUR special processing.
    # ------------------------
    if princ == 'SUR':
        try:
            sheet['Paid-On Revenue'] = sheet['Invoiced Dollars']
        except KeyError:
            pass
        # SUR is paid on resale.
        sheet['Comm Source'] = 'Resale'
    # ------------------------
    # XMO special processing.
    # ------------------------
    if princ == 'XMO':
        try:
            sheet['Paid-On Revenue'] = sheet['Invoiced Dollars']
        except KeyError:
            pass
        # XMO is paid on resale.
        sheet['Comm Source'] = 'Resale'


# %% Main function.
def main(filepaths, runningCom, fieldMappings):
    """Processes commission files and appends them to Running Commissions.

    Columns in individual commission files are identified and appended to the
    Running Commissions under the appropriate column, as identified by the
    fieldMappings file. Entries are then passed through the Lookup Master in
    search of a match to Reported Customer and Part Number. Distributors are
    corrected to consistent names. Entries with missing information are copied
    to Entries Need Fixing for further attention.

    Arguments:
    filepaths -- paths for opening (Excel) files to process.
    runningCom -- current Running Commissions file (in Excel) onto which we are
                  appending data.
    fieldMappings -- dataframe which links Running Commissions columns to
                     file data columns.
    """
    # Grab lookup table data names.
    columnNames = list(fieldMappings)
    # Add in non-lookup'd data names.
    columnNames[0:0] = ['CM Sales', 'Design Sales']
    columnNames[3:3] = ['T-Name', 'CM', 'T-End Cust']
    columnNames[7:7] = ['Principal', 'Corrected Distributor']
    columnNames[18:18] = ['Sales Commission']
    columnNames[20:20] = ['Quarter Shipped', 'Month', 'Year']
    columnNames.extend(['CM Split', 'TEMP/FINAL', 'Paid Date', 'From File',
                        'Sales Report Date'])

    # Set the directories for outputting data and finding lookups.
    outDir = 'Z:/MK Working Commissions/'
    lookDir = 'Z:/Commissions Lookup/'

    # -------------------------------------------------------------------
    # Check to see if there's an existing Running Commissions to append
    # the new data onto. If so, we need to do some work to get it ready.
    # -------------------------------------------------------------------
    if runningCom:
        finalData = pd.read_excel(runningCom, 'Master', dtype=str)
        # Convert applicable columns to numeric.
        numCols = ['Quantity', 'Ext. Cost', 'Invoiced Dollars',
                   'Paid-On Revenue', 'Actual Comm Paid', 'Unit Cost',
                   'Unit Price', 'CM Split', 'Year', 'Sales Commission',
                   'Split Percentage', 'Commission Rate',
                   'Gross Rev Reduction', 'Shared Rev Tier Rate']
        for col in numCols:
            try:
                finalData[col] = pd.to_numeric(finalData[col],
                                               errors='coerce').fillna('')
            except KeyError:
                pass
        # Convert individual numbers to numeric in rest of columns.
        mixedCols = [col for col in list(finalData) if col not in numCols]
        # Invoice/part numbers sometimes has leading zeros we'd like to keep.
        mixedCols.remove('Invoice Number')
        mixedCols.remove('Part Number')
        # The INF gets read in as infinity, so skip the principal column.
        mixedCols.remove('Principal')
        for col in mixedCols:
            finalData[col] = finalData[col].map(
                    lambda x: pd.to_numeric(x, errors='ignore'))
        # Now remove the nans.
        finalData.replace('nan', '', inplace=True)
        runComLen = len(finalData)
        filesProcessed = pd.read_excel(runningCom,
                                       'Files Processed').fillna('')
        print('Appending files to Running Commissions.')
        # Make sure column names all match.
        if set(list(finalData)) != set(columnNames):
            missCols = [i for i in columnNames if i not in finalData]
            addCols = [i for i in finalData if i not in columnNames]
            print('---\n'
                  'Columns in Running Commissions '
                  'do not match fieldMappings.xlsx!\n'
                  'Missing columns:\n%s' %
                  ', '.join(map(str, missCols))
                  + '\nExtra (erroneous) columns:\n%s' %
                  ', '.join(map(str, addCols))
                  + '\n*Program terminated*')
            return
        # Load in the matching Entries Need Fixing file.
        comDate = runningCom[-20:]
        fixName = outDir + 'Entries Need Fixing ' + comDate
        try:
            fixList = pd.read_excel(fixName, 'Data', dtype=str)
            # Convert applicable columns to numeric.
            numCols = ['Quantity', 'Ext. Cost', 'Invoiced Dollars',
                       'Paid-On Revenue', 'Actual Comm Paid', 'Unit Cost',
                       'Unit Price', 'CM Split', 'Year', 'Sales Commission',
                       'Split Percentage', 'Commission Rate',
                       'Gross Rev Reduction', 'Shared Rev Tier Rate']
            for col in numCols:
                try:
                    fixList[col] = pd.to_numeric(fixList[col],
                                                 errors='coerce').fillna('')
                except KeyError:
                    pass
            # Convert individual numbers to numeric in rest of columns.
            mixedCols = [col for col in list(fixList) if col not in numCols]
            # Invoice number sometimes has leading zeros we'd like to keep.
            mixedCols.remove('Invoice Number')
            # The INF gets read in as infinity, so skip the principal column.
            mixedCols.remove('Principal')
            for col in mixedCols:
                fixList[col] = fixList[col].map(
                        lambda x: pd.to_numeric(x, errors='ignore'))
            # Now remove the nans.
            fixList.replace('nan', '', inplace=True)
        except FileNotFoundError:
            print('No matching Entries Need Fixing file found for this '
                  'Running Commissions file!\n'
                  'Make sure ' + fixName
                  + ' is in the proper folder.\n'
                  '*Program terminated*')
            return
        except XLRDError:
            print('No sheet named Data found in Entries Need Fixing '
                  + fixName + '.xlsx!\n'
                  + '*Program terminated*')
            return
    # Start new Running Commissions.
    else:
        print('No Running Commissions file provided. Starting a new one.')
        runComLen = 0
        finalData = pd.DataFrame(columns=columnNames)
        filesProcessed = pd.DataFrame(columns=['Filename',
                                               'Total Commissions',
                                               'Date Added',
                                               'Paid Date'])

    # -------------------------------------------------------------------
    # Check to make sure we aren't duplicating files, then load in data.
    # -------------------------------------------------------------------
    # Strip the root off of the filepaths and leave just the filenames.
    filenames = [os.path.basename(val) for val in filepaths]
    # Check if we've duplicated any files.
    duplicates = list(set(filenames).intersection(filesProcessed['Filename']))
    # Don't let duplicate files get processed.
    filenames = [val for val in filenames if val not in duplicates]
    if duplicates:
        # Let us know we found duplictes and removed them.
        print('---\n'
              'The following files are already in Running Commissions:\n%s' %
              ', '.join(map(str, duplicates)))
        print('Duplicate files were removed from processing.')
        # Exit if no new files are left.
        if not filenames:
            print('---\n'
                  'No new commissions files selected.\n'
                  'Please try selecting files again.\n'
                  '*Program terminated*')
            return
    # Read in each new file with Pandas and store them as dictionaries.
    # Each dictionary has a dataframe for each sheet in the file.
    inputData = [pd.read_excel(filepath, None, dtype=str)
                 for filepath in filepaths]

    # --------------------------------------------------------------
    # Read in distMap. Terminate if not found or if errors in file.
    # --------------------------------------------------------------
    if os.path.exists(lookDir + 'distributorLookup.xlsx'):
        try:
            distMap = pd.read_excel(lookDir + 'distributorLookup.xlsx',
                                    'Distributors')
        except XLRDError:
            print('---\n'
                  'Error reading sheet name for distributorLookup.xlsx!\n'
                  'Please make sure the main tab is named Distributors.\n'
                  '*Program terminated*')
            return
        # Check the column names.
        distMapCols = ['Corrected Dist', 'Search Abbreviation']
        missCols = [i for i in distMapCols if i not in list(distMap)]
        if missCols:
            print('The following columns were not detected in '
                  'distributorLookup.xlsx:\n%s' %
                  ', '.join(map(str, missCols))
                  + '\n*Program terminated*')
            return
    else:
        print('---\n'
              'No distributor lookup file found!\n'
              'Please make sure distributorLookup.xlsx is in the directory.\n'
              '*Program terminated*')
        return

    # ------------------------------------------------------------------------
    # Read in the Lookup Master. Terminate if not found or if errors in file.
    # ------------------------------------------------------------------------
    if os.path.exists(lookDir + 'Lookup Master - Current.xlsx'):
        masterLookup = pd.read_excel(lookDir + 'Lookup Master - '
                                     'Current.xlsx').fillna('')
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
                  + '\n*Program terminated*')
            return
    else:
        print('---\n'
              'No Lookup Master found!\n'
              'Please make sure Lookup Master - Current.xlsx is '
              'in the directory.\n'
              '*Program terminated*')
        return

    # %% Done loading in the data and supporting files, now go to work.
    # Iterate through each file that we're appending to Running Commissions.
    fileNum = 0
    for filename in filenames:
        # Grab the next file from the list.
        newData = inputData[fileNum]
        fileNum += 1
        print('_'*54 + '\nWorking on file: ' + filename + '\n' + '_'*54)
        # Initialize total commissions for this file.
        totalComm = 0

        # -------------------------------------------------------------------
        # Detect principal from filename, terminate if not on approved list.
        # -------------------------------------------------------------------
        principal = filename[0:3]
        print('Principal detected as: ' + principal)
        princList = ['ABR', 'ATP', 'ATS', 'ATO', 'COS', 'EVE', 'GLO', 'INF',
                     'ISS', 'LAT', 'MIL', 'OSR', 'QRF', 'SUR', 'TRI', 'TRU']
        if principal not in princList:
            print('Principal supplied is not valid!\n'
                  'Current valid principals: '
                  + ', '.join(map(str, princList))
                  + '\nRemember to capitalize the principal abbreviation at'
                  'start of filename.'
                  '\n*Program terminated*')
            return

        # ----------------------------------------------------------------
        # Iterate over each dataframe in the ordered dictionary.
        # Each sheet in the file is its own dataframe in the dictionary.
        # ----------------------------------------------------------------
        for sheetName in list(newData):
            # Rework the index just in case it got read in wrong.
            sheet = newData[sheetName].reset_index(drop=True)
            # Remove the 'nan' strings that got read in.
            sheet.replace('nan', '', inplace=True)
            # Make sure index is an integer, not a string.
            sheet.index = sheet.index.map(int)
            # Create a duplicate of the sheet that stays unchanged aside
            # from recording matches.
            rawSheet = sheet.copy(deep=True)
            # Figure out if we've already added in the matches row.
            if filename.split('.')[0][-7:] != 'Matched':
                rawSheet.index += 1
            # Strip whitespace from column names.
            sheet.rename(columns=lambda x: str(x).strip(), inplace=True)
            # Clear out unnamed columns. Attribute error means it's an empty
            # sheet, so simply pass it along (it'll get dealt with).
            try:
                sheet = sheet.loc[:, ~sheet.columns.str.contains('^Unnamed')]
            except AttributeError:
                pass
            # Do specialized pre-processing tailored to principlal.
            renameDict = tailoredPreCalc(principal, sheet, sheetName)
            totalRows = sheet.shape[0]
            print('Found ' + str(totalRows) + ' entries in the tab '
                  + sheetName + '\n----------------------------------')
            # Iterate over each column of data that we want to append.
            for dataName in list(fieldMappings):
                # Grab list of names that the data could potentially be under.
                nameList = fieldMappings[dataName].dropna().tolist()
                # Look for a match in the sheet column names.
                sheetColumns = list(sheet)
                columnName = [val for val in sheetColumns if val in nameList]
                # If we found too many columns that match,
                # then rename the column in the sheet to the master name.
                if len(columnName) > 1:
                    print('Found multiple matches for ' + dataName
                          + '\nMatching columns: %s' %
                          ', '.join(map(str, columnName))
                          + '\nPlease fix column names and try again.\n'
                          '*Program terminated*')
                    return
                elif len(columnName) == 1:
                    sheet.rename(columns={columnName[0]: dataName},
                                 inplace=True)
                    if columnName[0] in renameDict.values():
                        columnName[0] = [i for i in renameDict.keys()
                                         if renameDict[i] == columnName[0]][0]
                    rawSheet.loc[0, columnName[0]] = dataName

            # Replace the old raw data sheet with the new one.
            rawSheet.sort_index(inplace=True)
            newData[sheetName] = rawSheet

            # Convert applicable columns to numeric.
            numCols = ['Quantity', 'Ext. Cost', 'Invoiced Dollars',
                       'Paid-On Revenue', 'Actual Comm Paid',
                       'Unit Cost', 'Unit Price']
            for col in numCols:
                try:
                    sheet[col] = pd.to_numeric(sheet[col],
                                               errors='coerce').fillna('')
                except KeyError:
                    pass

            # Fix Commission Rate if it got read in as a decimal.
            pctCols = ['Commission Rate', 'Split Percentage',
                       'Gross Rev Reduction', 'Shared Rev Tier Rate']
            for pctCol in pctCols:
                try:
                    # Remove '%' sign if present.
                    col = sheet[pctCol].astype(str).map(lambda x: x.strip('%'))
                    # Convert to numeric.
                    col = pd.to_numeric(col, errors='coerce')
                    # Identify which entries are not decimal.
                    notDec = col > 1
                    col[notDec] = col[notDec]/100
                    sheet[pctCol] = col.fillna(0)
                except (KeyError, TypeError):
                    pass

            # Do special processing for principal, if applicable.
            tailoredCalc(principal, sheet, sheetName, distMap)
            # Drop entries with emtpy part number or reported customer.
            try:
                sheet.drop(sheet[sheet['Part Number'] == ''].index,
                           inplace=True)
                sheet.reset_index(drop=True, inplace=True)
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
                      + '\n*Program terminated*')
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
                      '*Program terminated*')
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
                                                 ignore_index=True, sort=False)
                else:
                    print('Found no data on this tab. Moving on.\n'
                          '-')

        if totalComm > 0:
            # Show total commissions.
            print('Total commissions for this file: '
                  '${:,.2f}'.format(totalComm))
            # Append filename and total commissions to Files Processed sheet.
            currentDate = datetime.datetime.now().date()
            newFile = pd.DataFrame({'Filename': [filename],
                                    'Total Commissions': [totalComm],
                                    'Date Added': [currentDate],
                                    'Paid Date': ['']})
            filesProcessed = filesProcessed.append(newFile, ignore_index=True,
                                                   sort=False)
            # Save the matched raw data file.
            fname = filename[:-5]
            if filename[-12:] != 'Matched.xlsx':
                fname += ' Matched.xlsx'
            else:
                fname += '.xlsx'
            if saveError(fname):
                print('---\n'
                      'One or more of the raw data files are open in Excel.\n'
                      'Please close these files and try again.\n'
                      '*Program terminated*')
                return
            # Write the raw data file with matches.
            matchDir = 'Z:/Matched Raw Data Files/'
            writer = pd.ExcelWriter(matchDir + fname, engine='xlsxwriter',
                                    datetime_format='mm/dd/yyyy')
            for tab in list(newData):
                newData[tab].to_excel(writer, sheet_name=tab, index=False)
                # Format and fit each column.
                sheet = writer.sheets[tab]
                index = 0
                for col in newData[tab].columns:
                    # Set column width and formatting.
                    try:
                        maxWidth = max(len(str(val)) for val
                                       in newData[tab][col].values)
                    except ValueError:
                        maxWidth = 0
                    maxWidth = max(10, maxWidth)
                    sheet.set_column(index, index, maxWidth+0.8)
                    index += 1
            # Save the file.
            writer.save()
        else:
            print('No new data found in this file.\n'
                  'Moving on without adding file.')

    # %%
    # Fill NaNs left over from appending.
    finalData.fillna('', inplace=True)
    # Create the Entries Need Fixing dataframe (if not loaded in already).
    if not runningCom:
        fixList = pd.DataFrame(columns=list(finalData))
    # Find matches in Lookup Master and extract data from them.
    # Let us know how many rows are being processed.
    numRows = '{:,.0f}'.format(len(finalData) - runComLen)
    if numRows == '0':
        print('---\n'
              'No new valid data provided.\n'
              'Please check the new files for missing '
              'data or column matches.\n'
              '*Program terminated*')
        return
    print('---\n'
          'Beginning processing on ' + numRows + ' rows of data.')
    finalData.reset_index(inplace=True, drop=True)

    # Iterate over each row of the newly appended data.
    for row in range(runComLen, len(finalData)):
        # ------------------------------------------
        # Try to find a match in the Lookup Master.
        # ------------------------------------------
        # First match reported customer.
        repCust = str(finalData.loc[row, 'Reported Customer']).lower()
        POSCust = masterLookup['Reported Customer'].map(
                lambda x: str(x).lower())
        custMatches = masterLookup[repCust == POSCust]
        # Now match part number.
        partNum = str(finalData.loc[row, 'Part Number']).lower()
        PPN = masterLookup['Part Number'].map(lambda x: str(x).lower())
        # Reset index, but keep it around for updating usage below.
        fullMatch = custMatches[partNum == PPN].reset_index()
        # Record number of Lookup Master matches.
        lookMatches = len(fullMatch)
        # If we found one match we're good, so copy it over.
        if lookMatches == 1:
            fullMatch = fullMatch.iloc[0]
            # If there are no salespeople, it means we found a "soft match."
            # These have unknown End Customers and should go to
            # Entries Need Fixing. So, set them to zero matches.
            if fullMatch['CM Sales'] == fullMatch['Design Sales'] == '':
                lookMatches = 0
            # Grab primary and secondary sales people from Lookup Master.
            finalData.loc[row, 'CM Sales'] = fullMatch['CM Sales']
            finalData.loc[row, 'Design Sales'] = fullMatch['Design Sales']
            finalData.loc[row, 'T-Name'] = fullMatch['T-Name']
            finalData.loc[row, 'CM'] = fullMatch['CM']
            finalData.loc[row, 'T-End Cust'] = fullMatch['T-End Cust']
            finalData.loc[row, 'CM Split'] = fullMatch['CM Split']
            # Update usage in lookup Master.
            masterLookup.loc[fullMatch['index'],
                             'Last Used'] = datetime.datetime.now().date()
            # Update OOT city if not already filled in.
            if fullMatch['T-Name'][0:3] == 'OOT' and not fullMatch['City']:
                masterLookup.loc[fullMatch['index'],
                                 'City'] = finalData.loc[row, 'City']
        # If we found multiple matches, then fill in all the options.
        elif lookMatches > 1:
            lookCols = ['CM Sales', 'Design Sales', 'T-Name', 'CM',
                        'T-End Cust', 'CM Split']
            # Write list of all unique entries for each column.
            for col in lookCols:
                finalData.loc[row, col] = ', '.join(
                        fullMatch[col].map(lambda x: str(x)).unique())

        # -----------------------------------------------------------
        # Format the date correctly and fill in the Quarter Shipped.
        # -----------------------------------------------------------
        # Try parsing the date.
        dateError = False
        dateGiven = finalData.loc[row, 'Invoice Date']
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
            print('---'
                  'There is no Invoice Date column in Running Commissions!\n'
                  'Please check to make sure an Invoice Date column exists.\n'
                  'Note: Spelling, whitespace, and capitalization matter.\n'
                  '---')
            dateError = True
        # If no error found in date, fill in the month/year/quarter
        if not dateError:
            date = parse(dateGiven).date()
            # Make sure the date actually makes sense.
            currentYear = int(time.strftime('%Y'))
            if currentYear - date.year not in [0, 1] or date > currentDate:
                dateError = True
            else:
                # Cast date format into mm/dd/yyyy.
                finalData.loc[row, 'Invoice Date'] = date
                # Fill in quarter/year/month data.
                finalData.loc[row, 'Year'] = date.year
                month = calendar.month_name[date.month][0:3]
                finalData.loc[row, 'Month'] = month
                Qtr = str(math.ceil(date.month/3))
                finalData.loc[row, 'Quarter Shipped'] = (str(date.year) + 'Q'
                                                         + Qtr)

        # ---------------------------------------------------
        # Try to correct the distributor to consistent name.
        # ---------------------------------------------------
        # Strip extraneous characters and all spaces, and make lowercase.
        repDist = str(finalData.loc[row, 'Reported Distributor'])
        distName = re.sub('[^a-zA-Z0-9]', '', repDist).lower()

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

        # -----------------------------------------------------------------
        # Go through each column and convert applicable entries to numeric.
        # -----------------------------------------------------------------
        cols = list(finalData)
        # Invoice number sometimes has leading zeros we'd like to keep.
        cols.remove('Invoice Number')
        # The INF gets read in as infinity, so skip the principal column.
        cols.remove('Principal')
        for col in cols:
            finalData.loc[row, col] = pd.to_numeric(finalData.loc[row, col],
                                                    errors='ignore')

        # -----------------------------------------------------------------
        # If any data isn't found/parsed, copy over to Entries Need Fixing.
        # -----------------------------------------------------------------
        if lookMatches != 1 or len(distMatches) != 1 or dateError:
            fixList = fixList.append(finalData.loc[row, :], sort=False)
            fixList.loc[row, 'Running Com Index'] = row
            fixList.loc[row, 'Distributor Matches'] = len(distMatches)
            fixList.loc[row, 'Lookup Master Matches'] = lookMatches
            fixList.loc[row, 'Date Added'] = datetime.datetime.now().date()
            finalData.loc[row, 'TEMP/FINAL'] = 'TEMP'
        else:
            # Everything found, so entry is final.
            finalData.loc[row, 'TEMP/FINAL'] = 'FINAL'

        # Update progress every 1,000 rows.
        if row % 1000 == 0 and row > 0:
            print('Done with row ' '{:,.0f}'.format(row))
    # %% Clean up the finalized data.
    # Reorder columns to match the desired layout in columnNames.
    finalData.fillna('', inplace=True)
    finalData = finalData.loc[:, columnNames]
    columnNames.extend(['Distributor Matches', 'Lookup Master Matches',
                        'Date Added', 'Running Com Index'])
    # Fix up the Entries Need Fixing file.
    fixList = fixList.loc[:, columnNames]
    fixList.reset_index(drop=True, inplace=True)
    fixList.fillna('', inplace=True)
    # Make sure all the dates are formatted correctly.
    finalData['Invoice Date'] = finalData['Invoice Date'].map(
            lambda x: formDate(x))
    fixList['Invoice Date'] = fixList['Invoice Date'].map(
            lambda x: formDate(x))
    fixList['Date Added'] = fixList['Date Added'].map(lambda x: formDate(x))
    masterLookup['Last Used'] = masterLookup['Last Used'].map(
            lambda x: formDate(x))
    masterLookup['Date Added'] = masterLookup['Date Added'].map(
            lambda x: formDate(x))
    # %% Get ready to save files.
    # Check if the files we're going to save are open already.
    currentTime = time.strftime('%Y-%m-%d-%H%M')
    fname1 = outDir + 'Running Commissions ' + currentTime + '.xlsx'
    fname2 = outDir + 'Entries Need Fixing ' + currentTime + '.xlsx'
    fname3 = lookDir + 'Lookup Master - Current.xlsx'
    if saveError(fname1, fname2, fname3):
        print('---\n'
              'One or more of these files are currently open in Excel:\n'
              'Running Commissions, Entries Need Fixing, Lookup Master.\n'
              'Please close these files and try again.\n'
              '*Program terminated*')
        return

    # Write the Running Commissions file.
    writer1 = pd.ExcelWriter(fname1, engine='xlsxwriter',
                             datetime_format='mm/dd/yyyy')
    finalData.to_excel(writer1, sheet_name='Master', index=False)
    filesProcessed.to_excel(writer1, sheet_name='Files Processed',
                            index=False)
    # Format everything in Excel.
    tableFormat(finalData, 'Master', writer1)
    tableFormat(filesProcessed, 'Files Processed', writer1)

    # Write the Needs Fixing file.
    writer2 = pd.ExcelWriter(fname2, engine='xlsxwriter',
                             datetime_format='mm/dd/yyyy')
    fixList.to_excel(writer2, sheet_name='Data', index=False)
    # Format everything in Excel.
    tableFormat(fixList, 'Data', writer2)

    # Write the Lookup Master.
    writer3 = pd.ExcelWriter(fname3, engine='xlsxwriter',
                             datetime_format='mm/dd/yyyy')
    masterLookup.to_excel(writer3, sheet_name='Lookup', index=False)
    # Format everything in Excel.
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
