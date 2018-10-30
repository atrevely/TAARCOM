import pandas as pd
import os


def tableFormat(sheetData, sheetName, wbook):
    """Formats the Excel output as a table with correct column formatting."""
    # Create the table.
    sheet = wbook.sheets[sheetName]
    header = [{'header': val} for val in sheetData.columns.tolist()]
    setStyle = {'header_row': True, 'style': 'TableStyleLight1',
                'columns': header}
    sheet.add_table(0, 0, len(sheetData.index),
                    len(sheetData.columns)-1, setStyle)
    # Set document formatting.
    docFormat = wbook.book.add_format({'font': 'Calibri',
                                       'font_size': 11})
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
    """Writes comments from file into the Digikey Master.

    Arguments:
    filepaths -- The filepaths to the files with new comments.
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

    # Strip the root off of the filepaths and leave just the filenames.
    filenames = [os.path.basename(i) for i in filepaths]

    # Load the Insight files.
    inputData = [pd.read_excel(i, None) for i in filepaths]

    # Iterate through each file that we're appending to Digikey Master.
    fileNum = 0
    for filename in filenames:
        # Grab the next file from the list.
        newData = inputData[fileNum]
        fileNum += 1
        print('---\n'
              'Copying comments from file: ' + filename)

        # Iterate over each dataframe in the ordered dictionary.
        # Each sheet in the file is its own dataframe in the dictionary.
        for sheetName in list(newData):
            # Grab next sheet in file.
            # Rework the index just in case it got read in wrong.
            sheet = newData[sheetName].reset_index(drop=True).fillna('')

            # Go through and fill in comments for matching entries.
            for row in range(len(sheet)):
                matchMatrix = insMast == sheet.loc[row, :]
                # Remove comments from matching criteria.
                matchMatrix.drop(labels=['TAARCOM Comments', 'New Customer?'],
                                 axis=1, inplace=True)
                # Find matching index and copy comments.
                match = [i for i in range(len(matchMatrix))
                         if matchMatrix.loc[i, :].all()]
                if len(match) == 1:
                    comments = sheet.loc[row, 'TAARCOM Comments']
                    insMast.loc[max(match), 'TAARCOM Comments'] = comments
                elif len(match) > 1:
                    print('Multiple matches to Digikey Master found for row '
                          + str(row))
                else:
                    print('Match to Digikey Master not found for row '
                          + str(row))

    # Try saving the files, exit with error if any file is currently open.
    fname1 = 'Digikey Insight Master.xlsx'
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

    # Save the file.
    writer1.save()

    print('---\n'
          'Updates completed successfully!\n'
          '---\n'
          'Digikey Master updated.\n'
          '+++')
