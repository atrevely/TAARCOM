import pandas as pd
from dateutil.parser import parse
import time
import calendar
import math
import os.path
import re


# The main function.
def main(filepaths, oldMaster, lookupTable):
    """Processes Excel files and appends them to a master list.

    Keyword arguments:
    filepaths -- paths for opening (Excel) files to process.
    oldMaster -- current master list (in Excel) to which we are appending data.
    lookupTable -- dataframe which links master columns to file data columns.
    """

    # Get the master dataframe ready for the new data.
    # %%
    # Grab lookup table data names.
    columnNames = list(lookupTable)
    # Add in non-lookup'd data names.
    columnNames[0:0] = ['CM Sales', 'Design Sales', 'Quarter', 'Month', 'Year']
    columnNames[7:7] = ['T-End Cust', 'T-Name', 'CM',
                        'Principal', 'Corrected Distributor']
    columnNames.append('TEMP/FINAL')
    columnNames.append('Paid Date')
    columnNames.append('From File')

    # Check to see if we've supplied an existing master list to append to.
    if oldMaster:
        finalData = pd.read_excel(oldMaster, 'Master').fillna('')
        oldMastLen = len(finalData)
        filesProcessed = pd.read_excel(oldMaster, 'Files Processed')
        print('Appending files to old master.')
        if list(finalData) != columnNames:
            print('---')
            print('Columns in old master do not match current columns!')
            print('Please check column names and try again.')
            print('***')
            return
    # Start new master.
    else:
        print('No existing master list provided. Starting a new one.')
        oldMastLen = 0
        # These are our names for the data in the master list.
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
        print('---')
        print('The following files are already in the master:')
        for file in list(duplicates):
            print(file)
        print('Duplicate files were removed from processing.')
        # Exit if no new files are left.
        if not filenames:
            print('---')
            print('No new files selected.')
            print('Please try again.')
            print('***')
            return

    # Read in each new file with Pandas and store them as dictionaries.
    # Each dictionary has a dataframe for each sheet in the file.
    inputData = [pd.read_excel(filepath, None) for filepath in filepaths]

    # Read in distMap. Exit if not found.
    if os.path.exists('distributorLookup.xlsx'):
        distMap = pd.read_excel('distributorLookup.xlsx', 'Distributors')
    else:
        print('---')
        print('No distributor lookup found!')
        print('Please make sure distributorLookup.xlsx is in the directory.')
        print('***')
        return

    # Read in file of entries that need fixing. Exit if not found.
    if os.path.exists('Entries Need Fixing.xlsx'):
        fixList = pd.read_excel('Entries Need Fixing.xlsx', 'Data').fillna('')
    else:
        print('---')
        print('No Entries Need Fixing file found!')
        print('Please make sure Entries Need Fixing.xlsx is in the directory.')
        print('***')
        return

    # Read in the Master Lookup. Exit if not found.
    if os.path.exists('LookupMaster052018v2.xlsx'):
        masterLookup = pd.read_excel('LookupMaster052018v2.xlsx').fillna('')
    else:
        print('---')
        print('No Lookup Master found!')
        print('Please make sure LookupMaster*.xlsx is in the directory.')
        print('***')
        return

    # Go through each file, grab the new data, and put it in the master list.
    # %%
    # Iterate through each file that we're appending to the master list.
    fileNum = 0
    for filename in filenames:
        # Grab the next file from the list.
        newData = inputData[fileNum]
        fileNum += 1
        print('---')
        print('Working on file: ' + filename)
        # Set total commissions for file back to zero.
        totalComm = 0

        # Iterate over each dataframe in the ordered dictionary.
        # Each sheet in the file is its own dataframe in the dictionary.
        for sheetName in list(newData):
            # Grab next sheet in file.
            sheet = newData[sheetName]
            # Make sure index is an integer, not a string.
            sheet.index = sheet.index.map(int)
            totalRows = sheet.shape[0]
            print('Found ' + str(totalRows) + ' entries in the tab: '
                  + sheetName)

            # Iterate over each column of data that we want to append.
            for dataName in list(lookupTable):
                # Grab list of names that the data could potentially be under.
                nameList = lookupTable[dataName].dropna().tolist()
                # Look for a match in the sheet column names.
                sheetColumns = list(sheet)
                columnName = [val for val in sheetColumns if val in nameList]

                # Let us know if we didn't find a column that matches,
                # or if we found too many columns that match,
                # then rename the column in the sheet to the master name.
                if not columnName:
                    print('No column found for: ' + dataName)
                elif len(columnName) > 1:
                    print('Found multiple matches for: ' + dataName)
                    print('Please fix column names and try again.')
                    print('***')
                    return
                else:
                    sheet.rename(columns={columnName[0]: dataName},
                                 inplace=True)

            # Now that we've renamed all of the relevant columns,
            # append the new sheet to the master list, where only the properly
            # named columns are appended.
            if sheet.columns.duplicated().any():
                print('Two items are being mapped to the same master column!')
                print('Please check column mappings and try again.')
                print('***')
                return
            elif 'Actual Comm Paid' not in list(sheet):
                # Tab has no comission data, so it is ignored.
                print('No commission data found on this sheet.')
                print('Moving on.')
                print('-')
            else:
                matchingColumns = [val for val in list(sheet) if val in list(lookupTable)]
                if len(matchingColumns) > 0:
                    # Sum commissions paid on sheet.
                    print('Commissions for this sheet: '
                          + '${:,.2f}'.format(sheet['Actual Comm Paid'].sum()))
                    print('-')
                    totalComm += sheet['Actual Comm Paid'].sum()
                    # Strip whitespace from strings.
                    stringCols = [val for val in list(sheet) if sheet[val].dtype == 'object']
                    for col in stringCols:
                        sheet[col] = sheet[col].str.strip()
                    # Append matching data.
                    finalData = finalData.append(sheet[matchingColumns],
                                                 ignore_index=True)
                else:
                    print('Found no data on this sheet. Moving on.')
                    print('-')

        # Show total commissions.
        print('Total commissions for this file: '
              '${:,.2f}'.format(totalComm))
        # Append filename and commissions to Files Processed sheet.
        newFile = pd.DataFrame({'Filename': [filename],
                                'Total Commissions': [totalComm],
                                'Date Added': [time.strftime('%m/%d/%Y')],
                                'Paid Date': ['']})
        filesProcessed = filesProcessed.append(newFile, ignore_index=True)
        # Fill the NaNs in From File with the filename.
        finalData['From File'].fillna(filename, inplace=True)

    # Create and fill columns of derived data.
    # %%
    # Fill NaNs left over from appending.
    finalData.fillna('', inplace=True)
    # Find matches in Lookup Master and extract data from them.
    finalData['Reported Customer'] = finalData['Reported Customer'].astype(str)
    # Iterate over each row of the newly appended data.
    for row in range(oldMastLen, len(finalData)):
        # First match part number.
        partNoMatches = masterLookup[finalData.loc[row, 'Part Number'] == masterLookup['PPN']]
        # Next match End Customer.
        customerMatches = partNoMatches[finalData.loc[row, 'Reported Customer'].lower() == partNoMatches['POSCustomer'].str.lower()]
        customerMatches = customerMatches.reset_index()
        # Make sure we found exactly one match.
        if len(customerMatches) == 1:
            # Grab primary and secondary sales people from Lookup Master.
            finalData.loc[row, 'CM Sales'] = customerMatches['CM Sales'][0]
            finalData.loc[row, 'Design Sales'] = customerMatches['Design Sales'][0]
            finalData.loc[row, 'T-Name'] = customerMatches['Tname'][0]
            finalData.loc[row, 'CM'] = customerMatches['CM'][0]
            finalData.loc[row, 'T-End Cust'] = customerMatches['EndCustomer'][0]
            # Update usage in lookup master.
            masterLookup.loc[customerMatches['index'], 'Last Used'] = time.strftime('%m/%d/%Y')
            # Update OOT city.
            if customerMatches['Tname'][0][0:3] == 'OOT' and not customerMatches['City'][0]:
                 masterLookup.loc[customerMatches['index'], 'City'] = finalData.loc[row, 'City']

        # Try parsing the date.
        dateError = 0
        try:
            parse(finalData.loc[row, 'Invoice Date'])
        except ValueError:
            dateError = 1
        if not dateError:
            dateParsed = parse(finalData.loc[row, 'Invoice Date'])
            # Cast date format into mm/dd/yyyy.
            finalData.loc[row, 'Invoice Date'] = dateParsed.strftime('%m/%d/%Y')
            # Fill in quarter/year/month data.
            finalData.loc[row, 'Year'] = dateParsed.year
            finalData.loc[row, 'Month'] = calendar.month_name[dateParsed.month][0:3]
            finalData.loc[row, 'Quarter'] = str(dateParsed.year) + 'Q' + str(math.ceil(dateParsed.month/3))

        # Find a corrected distributor match.
        # Strip extraneous characters and all spaces, and make lowercase.
        distName = re.sub('[^a-zA-Z0-9]', '',
                          finalData.loc[row, 'Distributor']).lower()
        # Reset match count.
        matches = 0
        # Find match from distMap.
        for dist in distMap['Dist']:
            if dist in distName:
                # Check if it's already been matched.
                if matches > 0:
                    # Too many matches, so clear them and check by hand.
                    finalData.loc[row, 'Corrected Distributor'] = ''
                else:
                    # Input corrected distributor name.
                    finalData.loc[row, 'Corrected Distributor'] = distMap[distMap['Dist'] == dist]['Corrected Dist'].iloc[0]
                    matches += 1

        # If any data isn't found/parsed, copy entry to Fix Entries.
        if len(customerMatches) != 1 or matches != 1 or dateError:
            fixList = fixList.append(finalData.loc[row, :])
            fixList.loc[row, 'Master Index'] = row
            fixList.loc[row, 'Distributor Matches'] = matches
            fixList.loc[row, 'Lookup Master Matches'] = len(customerMatches)
            fixList.loc[row, 'Date Added'] = time.strftime('%m/%d/%Y')
            finalData.loc[row, 'TEMP/FINAL'] = 'TEMP'
        else:
            # Everything found, so entry is final.
            finalData.loc[row, 'TEMP/FINAL'] = 'FINAL'

    # Reorder columns to match the desired layout in columnNames.
    finalData = finalData.loc[:, columnNames]

    # Save the output as a .xlsx file.
    # %%
    # Save the Running Master file.
    writer = pd.ExcelWriter('Running Master ' + time.strftime('%Y-%m-%d-%H%M')
                            + '.xlsx', engine='xlsxwriter')
    finalData.to_excel(writer, sheet_name='Master', index=False)
    filesProcessed.to_excel(writer, sheet_name='Files Processed', index=False)
    try:
        writer.save()
    except IOError:
        print('---')
        print('Running Master is open in Excel!')
        print('Please close the file and try again.')
        print('***')
        return
    writer.save()

    # Save the Needs Fixing file.
    writer = pd.ExcelWriter('Entries Need Fixing.xlsx', engine='xlsxwriter')
    fixList.to_excel(writer, sheet_name='Data', index=False)
    try:
        writer.save()
    except IOError:
        print('---')
        print('Entries Need Fixing is open in Excel!')
        print('Please close the file and try again.')
        print('***')
        return
    writer.save()

    # Save the Lookup Master
    writer = pd.ExcelWriter('Lookup Master ' + time.strftime('%Y-%m-%d-%H%M')
                            + '.xlsx', engine='xlsxwriter')
    masterLookup.to_excel(writer, sheet_name='Lookup', index=False)
    try:
        writer.save()
    except IOError:
        print('---')
        print('Lookup Master is open in Excel!')
        print('Please close the file and try again.')
        print('***')
        return
    writer.save()

    print('---')
    print('Updates completed successfully!')
    print('---')
    print('Running Master updated.')
    print('Lookup Master updated.')
    print('Entries Need Fixing updated.')
    print('***')
