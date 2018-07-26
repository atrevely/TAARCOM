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
def main(filepaths):
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
                  'No new commissions files selected.\n'
                  'Please try selecting files again.\n'
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




    # Append the new data to the Insight Master.
    insMast = insMast.append(insFile, ignore_index=True)
    insMast = insMast.loc[:, colNames]

    # Write the Insight Master file.
    writer1 = pd.ExcelWriter('Digikey Insight Master.xlsx',
                             engine='xlsxwriter')
    insMast.to_excel(writer1, sheet_name='Master', index=False)
    filesProcessed.to_excel(writer1, sheet_name='Files Processed',
                            index=False)
    # Format as table in Excel.
    tableFormat(insMast, 'Master', writer1)
    tableFormat(filesProcessed, 'Files Processed', writer1)

    # Try saving the files, exit with error if any file is currently open.
    if saveError(writer1, writer2):
        print('---\n'
              'One or more files are currently open in Excel!\n'
              'Please close the files and try again.\n'
              '***')
        return

    # No errors, so save the files.
    writer1.save()

    print('---\n'
          'Updates completed successfully!\n'
          '---\n'
          'Digikey Master updated.\n'
          'New Root Customers updated.\n'
          '+++')
