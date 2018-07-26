import pandas as pd
import os


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
        maxWidth = max([len(str(val)) for val in sheetData[col].values])
        sheet.set_column(i, i, maxWidth+0.8, docFormat)
        i += 1


def saveError(*excelFiles):
    """Try saving Excel files and return True if any file is open."""
    for file in excelFiles:
        try:
            file.save()
        except IOError:
            return True
    return False


# The main function.
def main(filepath):
    """Appends new Digikey Insight file to the Digikey Insight Master.

    Arguments:
    filepath -- The filepath to the new Digikey Insight file.
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
    # Get column name layout.
    colNames = list(insMast)

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

    # Load the New Root Customers file.
    if os.path.exists('New Root Customers.xlsx'):
        newRootCusts = pd.read_excel('New Root Customers.xlsx',
                                     'Data').fillna('')
    else:
        print('---\n'
              'No New Root Customers file found!\n'
              'Please make sure New Root Customers'
              'is in the directory.\n'
              '***')
        return

    # Strip the root off of the filepath and leave just the filename.
    filename = os.path.basename(filepath)
    if filename in filesProcessed['Filename']:
        # Let us know the file is a duplicte.
        print('---\n'
              'The selected Insight file is already in the Insight Master!\n'
              '***')
        return
    newFile = pd.DataFrame({'Filename': [filename]})
    filesProcessed = filesProcessed.append(newFile, ignore_index=True)
    # Load the Insight file.
    insFile = pd.read_excel(filepath, None)
    insFile = insFile[list(insFile)[0]].fillna('')

    # Check to see if column names match.
    noMatch = [val for val in list(insMast) if val not in list(insFile)]
    if noMatch:
        print('The following Digikey Master columns were not found:')
        for colName in noMatch:
            print(colName)
        print('***')
        return

    # Go through each entry in the Insight file and look for a sales match.
    for row in range(len(insFile)):
        # Check for individuals and CMs and note them in comments.
        if 'contract' in insFile.loc[row, 'Root Customer Class'].lower():
            insFile.loc[row, 'TAARCOM Comments'] = 'Contract Manufacturer'
        if 'individual' in insFile.loc[row, 'Root Customer Class'].lower():
            insFile.loc[row, 'TAARCOM Comments'] = 'Individual'
        salesMatch = insFile.loc[row, 'Root Customer..'] == rootCustMap['Root Customer']
        match = rootCustMap[salesMatch]
        if len(match) == 1:
            # Match to salesperson if exactly one match is found.
            insFile.loc[row, 'Sales'] = match['Salesperson'].iloc[0]
        else:
            # Append to the New Root Customers file.
            newRootCusts = newRootCusts.append(insFile.loc[row, :],
                                               ignore_index=True)

    # Append the new data to the Insight Master.
    insMast = insMast.append(insFile, ignore_index=True)
    insMast = insMast.loc[:, colNames]
    newRootCusts = newRootCusts.loc[:, colNames]

    # Write the Insight Master file.
    writer1 = pd.ExcelWriter('Digikey Insight Master.xlsx',
                             engine='xlsxwriter')
    insMast.to_excel(writer1, sheet_name='Master', index=False)
    filesProcessed.to_excel(writer1, sheet_name='Files Processed',
                            index=False)
    # Format as table in Excel.
    tableFormat(insMast, 'Master', writer1)
    tableFormat(filesProcessed, 'Files Processed', writer1)

    # Write the New Root Customers file.
    writer2 = pd.ExcelWriter('New Root Customers.xlsx', engine='xlsxwriter')
    newRootCusts.to_excel(writer2, sheet_name='Data', index=False)
    # Format as table in Excel.
    tableFormat(newRootCusts, 'Data', writer2)

    # Try saving the files, exit with error if any file is currently open.
    if saveError(writer1, writer2):
        print('---\n'
              'One or more files are currently open in Excel!\n'
              'Please close the files and try again.\n'
              '***')
        return

    # No errors, so save the files.
    writer1.save()
    writer2.save()

    print('---\n'
          'Updates completed successfully!\n'
          '---\n'
          'Digikey Master updated.\n'
          'New Root Customers updated.\n'
          '+++')
