import pandas as pd
import time
import os.path


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
    # Add in derived data names.
    columnNames[0:0] = ['CM Sales', 'Design Sales']
    columnNames[3:3] = ['T-Name', 'CM', 'T-End Cust', 'Principal']
    columnNames[7:7] = ['Corrected Distributor']
    
    # LOAD UP PRINCIPAL HERE

    # Check to see if we've supplied an existing master list to append to,
    # otherwise start a new one.
    if oldMaster:
        finalData = pd.read_excel(oldMaster, 'Master')
        oldMastLen = len(finalData)
        filesProcessed = pd.read_excel(oldMaster, 'Files Processed')
        print('Appending files to old master.')
        if list(finalData) != columnNames:
            print('---')
            print('Columns in old master do not match current columns!')
            print('Please check column names and try again.')
            print('***')
            return
    else:
        print('No existing master list provided. Starting a new one.')
        # These are our names for the data in the master list.
        finalData = pd.DataFrame(columns=columnNames)
        filesProcessed = pd.DataFrame(columns=['Filenames',
                                               'Total Commissions'])

    # Strip the root off of the filepaths and leave just the filenames.
    filenames = [os.path.basename(val) for val in filepaths]
    # Check if we've duplicated any files.
    duplicates = list(set(filenames).intersection(filesProcessed['Filenames']))
    # Don't let duplicate files get processed.
    filenames = [val for val in filenames if val not in duplicates]
    if duplicates:
        # Let us know we found duplictes and removed them.
        print('---')
        print('The following files are already in the master:')
        for file in list(duplicates):
            print(file)
        print('Duplicate files were removed from processing.')
        if not filenames:
            print('---')
            print('Files were all duplicates.')
            print('Please try again.')
            print('***')
            return

    # Read in each new file with Pandas and store them as dictionaries.
    # Each dictionary has a dataframe for each sheet in the file.
    inputData = [pd.read_excel(filepath, None) for filepath in filepaths]

    # Read in the Master Lookup.
    masterLookup = pd.read_excel('LookupMaster052018.xlsx')
    # Fill NaNs with blank entries.
    masterLookup = masterLookup.fillna('')

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
            totalRows = sheet.shape[0]
            print('Found ' + str(totalRows) + ' entries in the tab: '
                  + sheetName)

            # Iterate over each column of data that we want to append.
            for dataName in list(lookupTable):
                # Grab list of names that the data could potentially be under.
                nameList = lookupTable[dataName].tolist()

                # Look for a match in the sheet column names.
                sheetColumns = list(sheet)
                columnName = [val for val in sheetColumns if val in nameList]

                # Let us know if we didn't find a column that matches,
                # or if we found too many columns that match,
                # then rename the column in the sheet to the master name.
                if columnName == []:
                    print('No column found for: ' + dataName)
                elif len(columnName) > 1:
                    print('Found multiple matches for: ' + dataName)
                    print('Please fix column names and try again.')
                    print('***')
                    return
                else:
                    sheet = sheet.rename(index=str,
                                         columns={columnName[0]: dataName})

            # Replace NaNs in the sheet with empty field.
            sheet = sheet.fillna('')

            # Now that we've renamed all of the relevant columns,
            # append the new sheet to the master list, where only the properly
            # named columns are appended.
            if sheet.columns.duplicated().any():
                print('Two items are being mapped to the same master column!')
                print('Please check column mappings and try again.')
                print('***')
                return
            elif 'Actual Comm Paid' not in list(sheet):
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
                    # Append matching data.
                    finalData = finalData.append(sheet[matchingColumns],
                                                 ignore_index=True)
                else:
                    print('Found no data on this sheet. Moving on.')
                    print('-')

        # Show total commissions.
        print('-')
        print('Total commissions for this file: '
              '${:,.2f}'.format(totalComm))
        # Append filename and commissions to Files Processed sheet.
        newFile = pd.DataFrame({'Filenames': [filename],
                                'Total Commissions': [totalComm]})
        filesProcessed = filesProcessed.append(newFile, ignore_index=True)

    # Create and fill columns of derived data.
    # %%
    # Find matches in Lookup Master and extract data from them.
    finalData['Billing Customer'] = finalData['Billing Customer'].astype(str)
    for row in range(oldMastLen, len(finalData)):
        # First match part number.
        partNoMatches = masterLookup.loc[finalData.loc[row, 'Part Number'] == masterLookup['PPN']]
        # Next match End Customer.
        customerMatches = partNoMatches.loc[finalData.loc[row, 'Billing Customer'].lower() == partNoMatches['POSCustomer'].str.lower()]
        # Make sure we found exactly one match.
        if len(customerMatches) == 1:
            # Grab primary and secondary sales people from Lookup Master.
            finalData.loc[row, 'CM Sales'] = customerMatches['CM Sales'][0]
            finalData.loc[row, 'Design Sales'] = customerMatches['Design Sales'][0]
            finalData.loc[row, 'T-Name'] = customerMatches['Tname'][0]
            finalData.loc[row, 'CM'] = customerMatches['CM'][0]
            finalData.loc[row, 'T-End Cust'] = customerMatches['EndCustomer'][0]

    # Clean up the master list before we save it.
    # %%
    # Reorder columns to match the lookup table.
    finalData = finalData.loc[:, columnNames]

    # Save the output as a .xlsx file.
    # %%
    writer = pd.ExcelWriter('CurrentMaster' + time.strftime('%Y-%m-%d-%H%M')
                            + '.xlsx', engine='xlsxwriter')
    finalData.to_excel(writer, sheet_name='Master', index=False)
    filesProcessed.to_excel(writer, sheet_name='Files Processed', index=False)
    writer.save()
    print('---')
    print('New master list generated.')
    print('***')
