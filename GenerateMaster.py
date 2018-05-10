import pandas as pd
import time


# The main function.
def main(filenames, oldMaster, lookupTable):
    """Processes Excel files and appends them to a master list.

    Keyword arguments:
    filenames -- filepaths for opening (Excel) files to process.
    oldMaster -- current master list (in Excel) to which we are appending data.
    lookupTable -- dataframe which links master columns to file data.
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
        filesProcessed = pd.DataFrame(columns=['Filepaths'])

    # Check if we've duplicated any files.
    duplicates = list(set(filenames).intersection(filesProcessed['Filepaths']))
    filenames = [val for val in filenames if val not in duplicates]
    if duplicates:
        # Don't let duplicate files get processed.
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

    # Add the new files we processed to the list.
    filenamesFrame = pd.DataFrame({'Filepaths': filenames})
    filesProcessed = filesProcessed.append(filenamesFrame,
                                           ignore_index=True)

    # Read in each new file with Pandas and store them as dictionaries.
    # Each dictionary has a dataframe for each sheet in the file.
    inputData = [pd.read_excel(filename, None) for filename in filenames]

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

    # Replace NaNs with empty cells.
    finalData = finalData.fillna('')

    # Reorder columns to match the lookup table.
    finalData = finalData.loc[:, list(lookupTable)]

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
