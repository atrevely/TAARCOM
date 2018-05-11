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
    # Check to see if we've supplied an existing master list to append to,
    # otherwise start a new one.
    if oldMaster:
        finalData = pd.read_excel(oldMaster, 'Master')
        filesProcessed = pd.read_excel(oldMaster, 'Files Processed')
        print('Appending files to old master.')
        if list(finalData) != list(lookupTable):
            print('---')
            print('Columns in old master do not match current columns!')
            print('Please check column names and try again.')
            return
    else:
        print('No existing master list provided. Starting a new one.')
        # These are our names for the data in the master list.
        columnNames = list(lookupTable)
        finalData = pd.DataFrame(columns=columnNames)
        filesProcessed = pd.DataFrame(columns=['Filenames'])

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
            print('---')
            return

    # Read in each new file with Pandas and store them as dictionaries.
    # Each dictionary has a dataframe for each sheet in the file.
    inputData = [pd.read_excel(filepath, None) for filepath in filepaths]

    # Read in the Master Lookup.
    masterLookup = pd.read_excel('LookupMaster052018.xlsx')
    masterLookup = masterLookup.fillna('')

    # Decide which columns we want formatted as dollar amounts.
    dollarCols = ['Cust Revenue YTD', 'Invoiced Dollars', 'Actual Comm Paid',
                  'Unit Price', 'Paid-On Revenue', 'Gross Comm Earned']

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

        # Iterate over each dataframe in the ordered dictionary.
        # Each sheet in the file is its own dataframe in the dictionary.
        for sheetName in list(newData):
            sheet = newData[sheetName]
            totalRows = sheet.shape[0]
            print('Found ' + str(totalRows) + ' entries in the tab: '
                  + sheetName)

            # Iterate over each column of data that we want to append.
            for dataName in list(finalData):
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
                    print('---')
                    return
                else:
                    sheet = sheet.rename(index=str,
                                         columns={columnName[0]: dataName})
                    # Format to dollar amount (with commas as thousands).
                    if dataName in dollarCols:
                        sheet[dataName] = sheet[dataName].apply(lambda x: '${:,.2f}'.format(x))

            # Now that we've renamed all of the relevant columns,
            # append the new sheet to the master list, where only the properly
            # named columns are appended.
            if sheet.columns.duplicated().any():
                print('Two items are being mapped to the same master column!')
                print('Please check column mappings and try again.')
                print('---')
                return
            else:
                matchingColumns = [val for val in list(sheet) if val in list(finalData)]
                if len(matchingColumns) > 0:
                    finalData = finalData.append(sheet[matchingColumns],
                                                 ignore_index=True)
                else:
                    print('Found no data on this sheet. Moving on.')

    # Create and fill columns of derived data.
    # %%
    # Add new columns to final data.
    finalData['Sales Primary'] = ''
    finalData['Sales Secondary'] = ''

    # Find match for each row in Lookup Master and tag it.
    for row in range(len(finalData)):
        # First match part number.
        partNoMatches = masterLookup.loc[finalData.loc[row, 'Part Number'] == masterLookup['PPN']]
        # Next match End Customer.
        finalMatch = partNoMatches.loc[finalData.loc[row, 'POS Customer'] == partNoMatches['POSCustomer']]
        # Make sure we found exactly one match.
        if len(finalMatch) == 1:
            # Grab primary and secondary sales people from Lookup Master.
            finalData.loc[row, 'Sales Primary'] = finalMatch['Sales'].str[0:2].tolist()[0]
            finalData.loc[row, 'Sales Secondary'] = finalMatch['Sales'].str[2:4].tolist()[0]
        elif len(finalMatch) > 1:
            print('Found multiple matches for row '
                  + str(row + 1) + ' in Lookup Master')
            print('---')
        else:
            print('Found no matches for row '
                  + str(row + 1) + ' in Lookup Master')
            print('---')


    # Clean up the master list before we save it.
    # %%
    # Replace NaNs with empty cells.
    finalData = finalData.fillna('')

    # Reorder columns to match the lookup table.
    finalData = finalData.loc[:, list(lookupTable)]

    # Add the new files we processed to the filepath list.
    filenamesFrame = pd.DataFrame({'Filenames': filenames})
    filesProcessed = filesProcessed.append(filenamesFrame,
                                           ignore_index=True)

    # Save the output as a .xlsx file.
    # %%
    writer = pd.ExcelWriter('CurrentMaster' + time.strftime('%Y-%m-%d-%H%M')
                            + '.xlsx', engine='xlsxwriter')
    finalData.to_excel(writer, sheet_name='Master')
    filesProcessed.to_excel(writer, sheet_name='Files Processed')
    writer.save()
    print('---')
    print('New master list generated.')
    print('---')
