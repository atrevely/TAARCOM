import pandas as pd
import time


# The main function.
def main(filenames, master):

    # Get the master dataframe ready for the new data.
    # %%
    # Check for an existing master list.
    oldMaster = master

    # Check to see if we've supplied an existing master list to append to,
    # otherwise start a new one.
    if oldMaster:
        finalData = pd.read_excel(oldMaster)
        print('Appending files to old master.')
    else:
        print('No existing master list provided. Starting a new one.')
        # These are our names for the data in the master list.
        finalData = pd.DataFrame(columns=['Invoice Number',
                                          'Invoice Date',
                                          'POS Customer',
                                          'Distributor',
                                          'PPN',
                                          'Invoice Amount',
                                          'Commission Rate',
                                          'Commission Paid'])

    # Read in each new file with Pandas and store them as dictionaries.
    # Each dictionary has a dataframe for each sheet in the file.
    inputData = [pd.read_excel(filename, None) for filename in filenames]

    # Create the lookup tables for each column.
    # %%
    # Start with an empty lookup table for each column of data.
    lookupTable = pd.DataFrame(columns=list(finalData))

    # Add the lists of keywords for each data column that we want to be able to
    # find in the new sheets.
    lookupTable.at[0, 'Invoice Number'] = ['Invoice Number',
                                           'Invoice']

    lookupTable.at[0, 'Invoice Date'] = ['Date',
                                         'Billing Date',
                                         'Trans Date',
                                         'Invoice Date',
                                         'POS_ShpDate']

    lookupTable.at[0, 'POS Customer'] = ['Ship To Name',
                                         'Customer Name',
                                         'Cust Name',
                                         'Customer',
                                         'Source Customer Name']

    lookupTable.at[0, 'Distributor'] = ['Sold To Name',
                                        'Distri',
                                        'Disti',
                                        'Distributor']

    lookupTable.at[0, 'PPN'] = ['Material Number',
                                'Product',
                                'Item',
                                'PtNo',
                                'Databook Product']

    lookupTable.at[0, 'Invoice Amount'] = ['Split Dollars',
                                           'AdjPOS',
                                           'Adj Inv Amt',
                                           'Post Split Amt']

    lookupTable.at[0, 'Commission Rate'] = ['Prod Class Rate',
                                            'Rate',
                                            'Comm Rate',
                                            'Commission Rate']

    lookupTable.at[0, 'Commission Paid'] = ['Commissions',
                                            'Act Comm Due',
                                            'Comm Due',
                                            'Post Split Amt']

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
            print('Found ' + str(totalRows) + ' entries in the tab: ' + sheetName)

            # Iterate over each column of data that we want to append.
            for dataName in list(finalData):
                # Grab list of names that the data could potentially be under.
                nameList = lookupTable.at[0, dataName]

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
                    print('Shutting down. Please fix column names.')
                    print('---')
                    return
                else:
                    sheet = sheet.rename(index=str, columns={columnName[0]: dataName})

            # Now that we've renamed all of the relevant columns,
            # append the new sheet to the master list, where only the properly
            # named columns are appended.
            if sheet.columns.duplicated().any():
                print('Found duplicate column names. Please fix.')
                print('Two items are being mapped to the same master column.')
                print('Shutting down.')
                print('---')
                return
            else:
                matchingColumns = [val for val in list(sheet) if val in list(finalData)]
                if len(matchingColumns) > 0:
                    finalData = finalData.append(sheet[matchingColumns], ignore_index=True)
                else:
                    print('Found no data on this sheet. Moving on.')

    # Replace NaNs with empty cells.
    finalData = finalData.fillna('')

    # Save the output as a .xlsx file.
    # %%
    writer = pd.ExcelWriter('CurrentMaster' + time.strftime('%Y-%m-%d') + '.xlsx', engine='xlsxwriter')
    finalData.to_excel(writer, sheet_name='Master')
    writer.save()
    print('---')
    print('New master list generated.')
    print('---')
