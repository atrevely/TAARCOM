import pandas as pd
from dateutil.parser import parse
import time
import calendar
import math
import os.path
import re


# The main function.
def main(filepaths, runningCom, fieldMappings, principal):
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
    columnNames[0:0] = ['CM Sales', 'Design Sales', 'Quarter', 'Month',
                        'Year']
    columnNames[7:7] = ['T-End Cust', 'T-Name', 'CM',
                        'Principal', 'Corrected Distributor']
    columnNames.extend(['CM Split', 'TEMP/FINAL', 'Paid Date', 'From File',
                        'Sales Report Date'])

    # Check to see if there's an existing Running Commissions to append to.
    if runningCom:
        finalData = pd.read_excel(runningCom, 'Master').fillna('')
        runComLen = len(finalData)
        filesProcessed = pd.read_excel(runningCom, 'Files Processed')
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

    # Read in distMap. Exit if not found.
    if os.path.exists('distributorLookup.xlsx'):
        distMap = pd.read_excel('distributorLookup.xlsx', 'Distributors')
    else:
        print('---\n'
              'No distributor lookup file found!\n'
              'Please make sure distributorLookup.xlsx is in the directory.\n'
              '***')
        return

    # Read in file of entries that need fixing. Exit if not found.
    if os.path.exists('Entries Need Fixing.xlsx'):
        fixList = pd.read_excel('Entries Need Fixing.xlsx', 'Data').fillna('')
    else:
        print('---\n'
              'No Entries Need Fixing file found!\n'
              'Please make sure Entries Need Fixing.xlsx'
              'is in the directory.\n'
              '***')
        return

    # Read in the Master Lookup. Exit if not found.
    if os.path.exists('Lookup Master 7-16-18.xlsx'):
        masterLookup = pd.read_excel('Lookup Master 7-16-18.xlsx').fillna('')
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
        print('---\n'
              'Working on file: ' + filename)
        # Set total commissions for file back to zero.
        totalComm = 0

        # Iterate over each dataframe in the ordered dictionary.
        # Each sheet in the file is its own dataframe in the dictionary.
        for sheetName in list(newData):
            # Grab next sheet in file.
            # Rework the index just in case it got read in wrong.
            sheet = newData[sheetName].reset_index(drop=True)
            # Clear out unnamed columns.
            sheet = sheet.loc[:, ~sheet.columns.str.contains('^Unnamed')]
            # Make sure index is an integer, not a string.
            sheet.index = sheet.index.map(int)
            totalRows = sheet.shape[0]
            print('Found ' + str(totalRows) + ' entries in the tab: '
                  + sheetName)

            # Iterate over each column of data that we want to append.
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
                    print('No column found for: ' + dataName)
                elif len(columnName) > 1:
                    print('Found multiple matches for: ' + dataName
                          + '\nMatching columns: %s' %
                          ', '.join(map(str, columnName))
                          + '\nPlease fix column names and try again.\n'
                          '***')
                    return
                else:
                    sheet.rename(columns={columnName[0]: dataName},
                                 inplace=True)

            # Now that we've renamed all of the relevant columns,
            # append the new sheet to Running Commissions, where only the
            # properly named columns are appended.
            if sheet.columns.duplicated().any():
                print('Two items are being mapped to the same column!\n'
                      'Please check fieldMappings.xlsx and try again.\n'
                      '***')
                return
            elif 'Actual Comm Paid' not in list(sheet):
                # Tab has no commission data, so it is ignored.
                print('No commission data found on this tab.\n'
                      'Moving on.\n'
                      '-')
            else:
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
                    finalData = finalData.append(sheet[matchingColumns],
                                                 ignore_index=True)
                else:
                    print('Found no data on this tab. Moving on.\n'
                          '-')

        # Show total commissions.
        print('Total commissions for this file: '
              '${:,.2f}'.format(totalComm))
        # Append filename and total commissions to Files Processed sheet.
        newFile = pd.DataFrame({'Filename': [filename],
                                'Total Commissions': [totalComm],
                                'Date Added': [time.strftime('%m/%d/%Y')],
                                'Paid Date': ['']})
        filesProcessed = filesProcessed.append(newFile, ignore_index=True)
        # Fill the NaNs in From File with the filename.
        finalData['From File'].fillna(filename, inplace=True)

    # %%
    # Fill NaNs left over from appending.
    finalData.fillna('', inplace=True)
    # Find matches in Lookup Master and extract data from them.
    # Let us know how many rows are being processed.
    numRows = '{:,.0f}'.format(len(finalData) - runComLen)
    print('---\n'
          'Beginning processing on ' + numRows + ' rows of data.')
    # Check to make sure Actual Comm Paid is all convertible to numeric.
    try:
        pd.to_numeric(finalData['Actual Comm Paid']).fillna(0)
    except ValueError:
        print('---\n'
              'Error parsing commission dollars.\n'
              'Make sure all data going into Actual Comm Paid is numeric.\n'
              'Note: The $ sign should be ok to use in numeric columns.\n'
              '***')
        return
    # Remove entries with no commissions dollars.
    finalData['Actual Comm Paid'] = pd.to_numeric(
            finalData['Actual Comm Paid']).fillna(0)
    finalData = finalData[finalData['Actual Comm Paid'] != 0]
    finalData.reset_index(inplace=True, drop=True)
    # Fill in principal.
    finalData.loc[:, 'Principal'] = principal

    # Iterate over each row of the newly appended data.
    for row in range(runComLen, len(finalData)):
        # First match part number.
        partMatch = finalData.loc[row, 'Part Number'] == masterLookup['PPN']
        partNoMatches = masterLookup[partMatch]
        # Next match reported customer.
        repCust = finalData.loc[row, 'Reported Customer'].lower()
        POSCust = partNoMatches['POSCustomer'].str.lower()
        custMatches = partNoMatches[repCust == POSCust].reset_index()
        # Record number of Lookup Master matches.
        lookMatches = len(custMatches)
        # Make sure we found exactly one match.
        if lookMatches == 1:
            custMatches = custMatches.iloc[0]
            # Grab primary and secondary sales people from Lookup Master.
            finalData.loc[row, 'CM Sales'] = custMatches['CM Sales']
            finalData.loc[row, 'Design Sales'] = custMatches['Design Sales']
            finalData.loc[row, 'T-Name'] = custMatches['Tname']
            finalData.loc[row, 'CM'] = custMatches['CM']
            finalData.loc[row, 'T-End Cust'] = custMatches['EndCustomer']
            finalData.loc[row, 'CM Split'] = custMatches['CM Split']
            # Update usage in lookup Master.
            masterLookup.loc[custMatches['index'],
                             'Last Used'] = time.strftime('%m/%d/%Y')
            # Update OOT city if not already filled in.
            if custMatches['Tname'][0:3] == 'OOT' and not custMatches['City']:
                masterLookup.loc[custMatches['index'],
                                 'City'] = finalData.loc[row, 'City']

        # Try parsing the date.
        dateError = 0
        try:
            parse(finalData.loc[row, 'Invoice Date'])
        except ValueError:
            # The date isn't recognized by the parser.
            dateError = 1
        except TypeError:
            # Check if Pandas read it in as a Timestamp object.
            # If so, turn it back into a string (a bit roundabout, oh well).
            if isinstance(finalData.loc[row, 'Invoice Date'], pd.Timestamp):
                finalData.loc[row, 'Invoice Date'] = str(
                        finalData.loc[row, 'Invoice Date'])
            else:
                dateError = 1
        # If no error found in date, fill in the month/year/quarter
        if not dateError:
            date = parse(finalData.loc[row, 'Invoice Date'])
            # Cast date format into mm/dd/yyyy.
            finalData.loc[row, 'Invoice Date'] = date.strftime('%m/%d/%Y')
            # Fill in quarter/year/month data.
            finalData.loc[row, 'Year'] = date.year
            finalData.loc[row, 'Month'] = calendar.month_name[date.month][0:3]
            finalData.loc[row, 'Quarter'] = (str(date.year)
                                             + 'Q'
                                             + str(math.ceil(date.month/3)))

        # Find a corrected distributor match.
        # Strip extraneous characters and all spaces, and make lowercase.
        distName = re.sub('[^a-zA-Z0-9]', '',
                          finalData.loc[row, 'Distributor']).lower()
        # Reset match count.
        distMatches = 0
        # Find match from distMap.
        for dist in distMap['Search Abbreviation']:
            if dist in distName:
                # Check if it's already been matched.
                if distMatches > 0:
                    # Too many matches, so clear them and check by hand.
                    finalData.loc[row, 'Corrected Distributor'] = ''
                else:
                    # Input corrected distributor name.
                    corrDist = distMap[distMap['Search Abbreviation'] == dist].iloc[0]
                    finalData.loc[row, 'Corrected Distributor'] = corrDist['Corrected Dist']
                    distMatches += 1

        # If any data isn't found/parsed, copy entry to Fix Entries.
        if lookMatches != 1 or distMatches != 1 or dateError:
            fixList = fixList.append(finalData.loc[row, :])
            fixList.loc[row, 'Running Com Index'] = row
            fixList.loc[row, 'Distributor Matches'] = distMatches
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
    finalData = finalData.loc[:, columnNames]
    columnNames.extend(['Distributor Matches', 'Lookup Master Matches',
                        'Date Added', 'Running Com Index'])
    fixList = fixList.loc[:, columnNames]

    # %%
    # Write the Running Commissions file.
    writer1 = pd.ExcelWriter('Running Commissions '
                             + time.strftime('%Y-%m-%d-%H%M')
                             + '.xlsx', engine='xlsxwriter')
    finalData.to_excel(writer1, sheet_name='Master', index=False)
    filesProcessed.to_excel(writer1, sheet_name='Files Processed',
                            index=False)
    sheet1a = writer1.sheets['Master']
    sheet1b = writer1.sheets['Files Processed']
    # Format as table.
    header1a = [{'header': val} for val in finalData.columns.tolist()]
    header1b = [{'header': val} for val in filesProcessed.columns.tolist()]
    set1a = {'header_row': True, 'style': 'TableStyleMedium5',
             'columns': header1a}
    set1b = {'header_row': True, 'style': 'TableStyleMedium5',
             'columns': header1b}
    sheet1a.add_table(0, 0, len(finalData.index),
                      len(finalData.columns)-1, set1a)
    sheet1b.add_table(0, 0, len(filesProcessed.index),
                      len(filesProcessed.columns)-1, set1b)

    # Write the Needs Fixing file.
    writer2 = pd.ExcelWriter('Entries Need Fixing.xlsx', engine='xlsxwriter')
    fixList.to_excel(writer2, sheet_name='Data', index=False)
    sheet2 = writer2.sheets['Data']
    # Format as table.
    header2 = [{'header': val} for val in fixList.columns.tolist()]
    set2 = {'header_row': True, 'style': 'TableStyleMedium5',
            'columns': header2}
    sheet2.add_table(0, 0, len(fixList.index), len(fixList.columns)-1, set2)

    # Write the Lookup Master.
    writer3 = pd.ExcelWriter('Lookup Master ' + time.strftime('%Y-%m-%d-%H%M')
                             + '.xlsx', engine='xlsxwriter')
    masterLookup.to_excel(writer3, sheet_name='Lookup', index=False)
    sheet3 = writer3.sheets['Lookup']
    # Format as table.
    header3 = [{'header': val} for val in masterLookup.columns.tolist()]
    set3 = {'header_row': True, 'style': 'TableStyleMedium5',
            'columns': header3}
    sheet3.add_table(0, 0, len(masterLookup.index),
                     len(masterLookup.columns)-1, set3)

    # Try saving the files, exit with error if any file is currently open.
    try:
        writer1.save()
    except IOError:
        print('---\n'
              'Running Commissions is open in Excel!\n'
              'Please close the file and try again.\n'
              '***')
        return
    try:
        writer2.save()
    except IOError:
        print('---\n'
              'Lookup Master is open in Excel!\n'
              'Please close the file and try again.\n'
              '***')
        return
    try:
        writer3.save()
    except IOError:
        print('---\n'
              'Entries Need Fixing is open in Excel!\n'
              'Please close the file and try again.\n'
              '***')
        return

    # If no errors, save the files.
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
