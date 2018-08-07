import pandas as pd
import os
import time


def tableFormat(sheetData, sheetName, wbook):
    """Formats the Excel output as a table with correct column formatting."""
    # Create the table.
    sheet = wbook.sheets[sheetName]
    header = [{'header': val} for val in sheetData.columns.tolist()]
    setStyle = {'header_row': True, 'style': 'TableStyleMedium5',
                'columns': header}
    sheet.add_table(0, 0, len(sheetData.index),
                    len(sheetData.columns)-1, setStyle)
    # Set document formatting.
    docFormat = wbook.book.add_format({'font': 'Century Gothic',
                                       'font_size': 8})
    # Format and fit each column.
    i = 0
    for col in sheetData.columns:
        # Set column width and formatting.
        maxWidth = max(len(str(val)) for val in sheetData[col].values)
        sheet.set_column(i, i, maxWidth+0.8, docFormat)
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


# The main function.
def main(filepaths):
    """Appends new Digikey Insight file to the Digikey Insight Master.

    Arguments:
    filepaths -- The filepaths to the files that will be appended.
    """
    # Load the Digikey Insights Master file.
    if os.path.exists('Digikey Insight Master.xlsx'):
        insMast = pd.read_excel('Digikey Insight Master.xlsx',
                                'Master').fillna('')
        filesProcessed = pd.read_excel('Digikey Insight Master.xlsx',
                                       'Files Processed').fillna('')
    else:
        print('---\n'
              'No Insight Master file found!\n'
              'Please make sure Digikey Insight Master is in the directory.\n'
              '***')
        return

    # Load the Root Customer Mappings file.
    if os.path.exists('rootCustomerMappings.xlsx'):
        rootCustMap = pd.read_excel('rootCustomerMappings.xlsx',
                                    'Sales Lookup').fillna('')
    else:
        print('---\n'
              'No Root Customer Mappings file found!\n'
              'Please make sure rootCustomerMappings.xlsx'
              'is in the directory.\n'
              '***')
        return

    # Get column name layout, prepare combined insight file.
    colNames = list(insMast)
    newDatComb = pd.DataFrame(columns=colNames)

    # Strip the root off of the filepaths and leave just the filenames.
    filenames = [os.path.basename(val) for val in filepaths]
    # Check if we've duplicated any files.
    duplicates = list(set(filenames).intersection(filesProcessed['Filename']))
    # Don't let duplicate files get processed.
    filenames = [val for val in filenames if val not in duplicates]
    if duplicates:
        # Let us know we found duplictes and removed them.
        print('---\n'
              'The following files are already in Digikey Master:')
        for file in list(duplicates):
            print(file)
        print('Duplicate files were removed from processing.')
        # Exit if no new files are left.
        if not filenames:
            print('---\n'
                  'No new insight files selected.\n'
                  'Please try selecting files again.\n'
                  '***')
            return

    newFiles = pd.DataFrame({'Filename': filenames})
    filesProcessed = filesProcessed.append(newFiles, ignore_index=True)
    # Load the Insight files.
    inputData = [pd.read_excel(filepath, None) for filepath in filepaths]

    # Iterate through each file that we're appending to Digikey Master.
    fileNum = 0
    for filename in filenames:
        # Grab the next file from the list.
        newData = inputData[fileNum]
        fileNum += 1
        print('---\n'
              'Working on file: ' + filename)

        # Iterate over each dataframe in the ordered dictionary.
        # Each sheet in the file is its own dataframe in the dictionary.
        for sheetName in list(newData):
            # Grab next sheet in file.
            # Rework the index just in case it got read in wrong.
            sheet = newData[sheetName].reset_index(drop=True).fillna('')

            # Check to see if column names match.
            noMatch = [val for val in list(insMast) if val not in list(sheet)]
            if noMatch:
                print('The following Digikey Master columns were not found:')
                for colName in noMatch:
                    print(colName)
                print('***')
                return

            # Append new salespeople mappings to rootCustMappings.
            for row in range(len(sheet)):
                # Get root customer and salesperson.
                cust = sheet.loc[row, 'Root Customer..']
                salesperson = sheet.loc[row, 'Sales']
                if not salesperson:
                    print('Not all entries in the Sales column are filled in!'
                          '\nPlease check Sales column for each file.'
                          '\n***')
                if cust:
                    # Find match in rootCustomerMappings.
                    custMatch = rootCustMap['Root Customer'] == cust
                    if sum(custMatch) == 1:
                        matchID = rootCustMap[custMatch].index
                        # Input (possibly new) salesperson.
                        rootCustMap.loc[matchID, 'Salesperson'] = salesperson
                    elif not custMatch.any():
                        # New customer (no match), so append to mappings.
                        newCust = pd.DataFrame({'Root Customer': [cust],
                                                'Salesperson': [salesperson]})
                        rootCustMap = rootCustMap.append(newCust,
                                                         ignore_index=True)
                    else:
                        print('There appears to be a duplicate customer in'
                              ' rootCustomerMappings:\n'
                              + str(cust)
                              + '\n***')
                        return

            # Append the sheet to the combined dataframe.
            newDatComb = newDatComb.append(sheet, ignore_index=True)

    # Go through the combined insights and prepare sales reports.
    salespeople = newDatComb['Sales'].unique()
    salespeople = [val for val in salespeople if len(val) == 2]
    for sales in salespeople:
        repDat = newDatComb[newDatComb['Sales'] == sales]
        repDat = repDat.loc[:, colNames]

        # Try saving.
        fname = ('Digikey Insights Report'
                 + time.strftime(' %m-%d-%Y - ') + sales + '.xlsx')
        if saveError(fname):
            print('---\n'
                  'One of the report files is currently open in Excel!\n'
                  'Please close the file and try again.\n'
                  '***')
            return

        # Write report to file.
        writer = pd.ExcelWriter(fname, engine='xlsxwriter')
        repDat.to_excel(writer, sheet_name='Report Data', index=False)
        # Format as table in Excel.
        tableFormat(repDat, 'Report Data', writer)
        writer.save()

    # Append the new data to the Insight Master.
    insMast = insMast.append(newDatComb, ignore_index=True)
    insMast = insMast.loc[:, colNames]

    # Try saving the files, exit with error if any file is currently open.
    fname1 = 'Digikey Insight Master.xlsx'
    fname2 = 'rootCustomerMappings.xlsx'
    fname3 = ('Digikey Insights Report ' + time.strftime('%m-%d-%Y') +
              ' - Full Report.xlsx')
    if saveError(fname1):
        print('---\n'
              'Insight Master is currently open in Excel!\n'
              'Please close the file and try again.\n'
              '***')
        return

    # Write the Insight Master file.
    writer1 = pd.ExcelWriter(fname1, engine='xlsxwriter')
    insMast.to_excel(writer1, sheet_name='Master', index=False)
    filesProcessed.to_excel(writer1, sheet_name='Files Processed',
                            index=False)
    # Format as table in Excel.
    tableFormat(insMast, 'Master', writer1)
    tableFormat(filesProcessed, 'Files Processed', writer1)

    # Write the rootCustomerMappings file.
    writer2 = pd.ExcelWriter(fname2, engine='xlsxwriter')
    rootCustMap.to_excel(writer2, sheet_name='Sales Lookup', index=False)
    # Format as table in Excel.
    tableFormat(rootCustMap, 'Sales Lookup', writer2)

    # Write the full salespeople file.
    writer3 = pd.ExcelWriter(fname3, engine='xlsxwriter')
    newDatComb.to_excel(writer3, sheet_name='Data', index=False)
    # Format as table in Excel.
    tableFormat(newDatComb, 'Data', writer3)

    # Save the file.
    writer1.save()
    writer2.save()
    writer3.save()

    print('---\n'
          'Updates completed successfully!\n'
          '---\n'
          'Digikey Master updated.\n'
          'New Root Customers updated.\n'
          '+++')
