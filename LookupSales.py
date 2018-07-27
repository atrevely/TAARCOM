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
    commaFormat = wbook.book.add_format({'font': 'Century Gothic',
                                         'font_size': 8,
                                         'num_format': 3})
    # Format and fit each column.
    i = 0
    for col in sheetData.columns:
        # Set column width and formatting.
        if col == 'Qty Shipped':
            formatting = commaFormat
        else:
            formatting = docFormat
        maxWidth = max([len(str(val)) for val in sheetData[col].values])
        sheet.set_column(i, i, maxWidth+0.8, formatting)
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

    # Strip the root off of the filepath and leave just the filename.
    filename = os.path.basename(filepath)

    # Load the Insight file.
    insFile = pd.read_excel(filepath, None)
    insFile = insFile[list(insFile)[0]].fillna('')

    # Get the output files ready.
    newInsFile = pd.DataFrame(columns=list(insFile))
    newRootCusts = pd.DataFrame(columns=list(insFile))

    # Go through each entry in the Insight file and look for a sales match.
    for row in range(len(insFile)):
        # Check for individuals and CMs and note them in comments.
        if 'contract' in insFile.loc[row, 'Root Customer Class'].lower():
            insFile.loc[row, 'TAARCOM Comments'] = 'Contract Manufacturer'
        if 'individual' in insFile.loc[row, 'Root Customer Class'].lower():
            insFile.loc[row, 'TAARCOM Comments'] = 'Individual'
        salesMatch = insFile.loc[row, 'Root Customer..'] == rootCustMap['Root Customer']
        match = rootCustMap[salesMatch]
        # Convert applicable entries to numeric.
        for col in list(insFile):
            insFile.loc[row, col] = pd.to_numeric(insFile.loc[row, col],
                                                  errors='ignore')
        if len(match) == 1:
            # Match to salesperson if exactly one match is found.
            insFile.loc[row, 'Sales'] = match['Salesperson'].iloc[0]
            newInsFile = newInsFile.append(insFile.loc[row, :],
                                           ignore_index=True)
        else:
            # Append to the New Root Customers file.
            newRootCusts = newRootCusts.append(insFile.loc[row, :],
                                               ignore_index=True)

    # Write the Insight file, which now contains salespeople.
    writer1 = pd.ExcelWriter(filename[:-5] + ' With Salespeople.xlsx',
                             engine='xlsxwriter')
    newInsFile.to_excel(writer1, sheet_name='Data', index=False)
    # Format as table in Excel.
    tableFormat(newInsFile, 'Data', writer1)

    # Write the New Root Customers file.
    writer2 = pd.ExcelWriter(filename[:-5] + ' New Root Customers.xlsx',
                             engine='xlsxwriter')
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
