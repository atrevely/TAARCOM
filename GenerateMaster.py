import pandas as pd
from dateutil.parser import parse
import time
import calendar
import math
import os.path
import re


# The main function.
def main(filepaths, runningCom, fieldMappings):
    """Processes commission files and appends them to Running Commissions.

    Arguments:
    filepaths -- paths for opening (Excel) files to process.
    runningCom -- current Running Commissions file (in Excel) to
                  which we are appending data.
    fieldMappings -- dataframe which links Running Commissions columns to
                     file data columns.
    """

    # Get the master dataframe ready for the new data.
    # %%
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
            print('---')
            print('Columns in Running Commissions'
                  'do not match fieldMappings.xlsx!')
            print('Please check column names and try again.')
            print('***')
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
        print('---')
        print('The following files are already in Running Commissions:')
        for file in list(duplicates):
            print(file)
        print('Duplicate files were removed from processing.')
        # Exit if no new files are left.
        if not filenames:
            print('---')
            print('No new commissions files selected.')
            print('Please try selecting files again.')
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
        print('No distributor lookup file found!')
        print('Please make sure distributorLookup.xlsx is in the directory.')
        print('***')
        return

    # Read in file of entries that need fixing. Exit if not found.
    if os.path.exists('Entries Need Fixing.xlsx'):
        fixList = pd.read_excel('Entries Need Fixing.xlsx', 'Data').fillna('')
    else:
        print('---')
        print('No Entries Need Fixing file found!')
        print('Please make sure Entries Need Fixing.xlsx'
              'is in the directory.')
        print('***')
        return

    # Read in the Master Lookup. Exit if not found.
    if os.path.exists('Lookup Master 6-27-18.xlsx'):
        masterLookup = pd.read_excel('Lookup Master 6-27-18.xlsx').fillna('')
    else:
        print('---')
        print('No Lookup Master found!')
        print('Please make sure lookupMaster.xlsx is in the directory.')
        print('***')
        return

    # Go through each file, grab the data, and put it in Running Commissions.
    # %%
    # Iterate through each file that we're appending to Running Commissions.
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
                    print('Found multiple matches for: ' + dataName)
                    print('Please fix column names and try again.')
                    print('***')
                    return
                else:
                    sheet.rename(columns={columnName[0]: dataName},
                                 inplace=True)

            # Now that we've renamed all of the relevant columns,
            # append the new sheet to Running Commissions, where only the
            # properly named columns are appended.
            if sheet.columns.duplicated().any():
                print('Two items are being mapped to the same column!')
                print('Please check fieldMappings.xlsx and try again.')
                print('***')
                return
            elif 'Actual Comm Paid' not in list(sheet):
                # Tab has no commission data, so it is ignored.
                print('No commission data found on this tab.')
                print('Moving on.')
                print('-')
            else:
                matchingColumns = [val for val in list(sheet) if val in list(fieldMappings)]
                if len(matchingColumns) > 0:
                    # Sum commissions paid on sheet.
                    print('Commissions for this tab: '
                          + '${:,.2f}'.format(sheet['Actual Comm Paid'].sum()))
                    print('-')
                    totalComm += sheet['Actual Comm Paid'].sum()
                    # Strip whitespace from all strings in dataframe.
                    stringCols = [val for val in list(sheet) if sheet[val].dtype == 'object']
                    for col in stringCols:
                        sheet[col] = sheet[col].fillna('').astype(str).map(lambda x: x.strip())
                    # Append matching columns of data.
                    finalData = finalData.append(sheet[matchingColumns],
                                                 ignore_index=True)
                else:
                    print('Found no data on this tab. Moving on.')
                    print('-')

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

    # Create and fill columns of derived data.
    # %%
    # Fill NaNs left over from appending.
    finalData.fillna('', inplace=True)
    # Find matches in Lookup Master and extract data from them.
    # Let us know how many rows are being processed.
    numRows = '{:,.0f}'.format(len(finalData) - runComLen)
    print('---')
    print('Beginning processing on ' + numRows + ' rows')
    # Check to make sure Actual Comm Paid is all convertible to numeric.
    try:
        pd.to_numeric(finalData['Actual Comm Paid']).fillna(0)
    except ValueError:
        print('---')
        print('Error parsing commission dollars.')
        print('Make sure all data going into Actual Comm Paid is numeric.')
        print('Note: The $ sign should be ok to use in numeric columns.')
        print('***')
        return
    # Remove entries with no commissions dollars.
    finalData['Actual Comm Paid'] = pd.to_numeric(finalData['Actual Comm Paid']).fillna(0)
    finalData = finalData[finalData['Actual Comm Paid'] != 0]
    finalData.reset_index(inplace=True, drop=True)

    # Iterate over each row of the newly appended data.
    for row in range(runComLen, len(finalData)):
        # First match part number.
        partMatch = finalData.loc[row, 'Part Number'] == masterLookup['PPN']
        partNoMatches = masterLookup[partMatch]
        # Next match Reported Customer.
        custMatch = finalData.loc[row, 'Reported Customer'].lower() == partNoMatches['POSCustomer'].str.lower()
        custMatches = partNoMatches[custMatch].reset_index()
        # Make sure we found exactly one match.
        if len(custMatches) == 1:
            # Grab primary and secondary sales people from Lookup Master.
            finalData.loc[row, 'CM Sales'] = custMatches['CM Sales'][0]
            finalData.loc[row, 'Design Sales'] = custMatches['Design Sales'][0]
            finalData.loc[row, 'T-Name'] = custMatches['Tname'][0]
            finalData.loc[row, 'CM'] = custMatches['CM'][0]
            finalData.loc[row, 'T-End Cust'] = custMatches['EndCustomer'][0]
            finalData.loc[row, 'CM Split'] = custMatches['CM Split'][0]
            # Update usage in lookup Master.
            masterLookup.loc[custMatches['index'],
                             'Last Used'] = time.strftime('%m/%d/%Y')
            # Update OOT city if not already filled in.
            if custMatches['Tname'][0][0:3] == 'OOT' and not custMatches['City'][0]:
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
            # If so, turn it back into a string.
            if isinstance(finalData.loc[row, 'Invoice Date'], pd.Timestamp):
                finalData.loc[row,'Invoice Date'] = str(finalData.loc[row, 'Invoice Date'])
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
        matches = 0
        # Find match from distMap.
        for dist in distMap['Search Abbreviation']:
            if dist in distName:
                # Check if it's already been matched.
                if matches > 0:
                    # Too many matches, so clear them and check by hand.
                    finalData.loc[row, 'Corrected Distributor'] = ''
                else:
                    # Input corrected distributor name.
                    finalData.loc[row, 'Corrected Distributor'] = distMap[distMap['Search Abbreviation'] == dist]['Corrected Dist'].iloc[0]
                    matches += 1

        # If any data isn't found/parsed, copy entry to Fix Entries.
        if len(custMatches) != 1 or matches != 1 or dateError:
            fixList = fixList.append(finalData.loc[row, :])
            fixList.loc[row, 'Running Com Index'] = row
            fixList.loc[row, 'Distributor Matches'] = matches
            fixList.loc[row, 'Lookup Master Matches'] = len(custMatches)
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

    # Save the output as a .xlsx file.
    # %%
    # Write the Running Commissions file.
    writer1 = pd.ExcelWriter('Running Commissions '
                             + time.strftime('%Y-%m-%d-%H%M')
                             + '.xlsx', engine='xlsxwriter')
    finalData.to_excel(writer1, sheet_name='Master', index=False)
    filesProcessed.to_excel(writer1, sheet_name='Files Processed',
                            index=False)
    # Write the Needs Fixing file.
    writer2 = pd.ExcelWriter('Entries Need Fixing.xlsx', engine='xlsxwriter')
    fixList.to_excel(writer2, sheet_name='Data', index=False)
    # Write the Lookup Master.
    writer3 = pd.ExcelWriter('Lookup Master ' + time.strftime('%Y-%m-%d-%H%M')
                             + '.xlsx', engine='xlsxwriter')
    masterLookup.to_excel(writer3, sheet_name='Lookup', index=False)

    try:
        writer1.save()
    except IOError:
        print('---')
        print('Running Commissions is open in Excel!')
        print('Please close the file and try again.')
        print('***')
        return
    try:
        writer2.save()
    except IOError:
        print('---')
        print('Lookup Master is open in Excel!')
        print('Please close the file and try again.')
        print('***')
        return
    try:
        writer3.save()
    except IOError:
        print('---')
        print('Entries Need Fixing is open in Excel!')
        print('Please close the file and try again.')
        print('***')
        return

    # If no errors, save the files.
    writer1.save()
    writer2.save()
    writer3.save()

    print('---')
    print('Updates completed successfully!')
    print('---')
    print('Running Commissions updated.')
    print('Lookup Master updated.')
    print('Entries Need Fixing updated.')
    print('+++')
