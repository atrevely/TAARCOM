import pandas as pd
import os
import time
from xlrd import XLRDError


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
    """Combine files into one finalized monthly Digikey file.

    Arguments:
    filepaths -- The filepaths to the files with new comments.
    """
    # Strip the root off of the filepaths and leave just the filenames.
    filenames = [os.path.basename(i) for i in filepaths]

    # Load the Insight files.
    try:
        inputData = [pd.read_excel(i, 'Full Data') for i in filepaths]
    except XLRDError:
        print('---\n'
              'Error reading sheet name(s) for Digikey Reports!\n'
              'Please make sure the report tabs are named Full Data in '
              'each file.\n'
              '***')

    # Make sure each filename has a salesperson initials.
    salespeople = ['CM', 'CR', 'DC', 'HS', 'IT', 'JC', 'JW', 'KC', 'LK',
                   'MG', 'MM', 'VD']
    for filename in filenames:
        inits = filename[0:2]
        if inits not in salespeople:
            print('Salesperson initials ' + inits + ' not recognized!\n'
                  'Make sure the first two letters of each filename are '
                  'salesperson initials (capitalized).\n'
                  '***')
            return

    # Create the master dataframe to append to.
    finalData = pd.DataFrame(columns=list(inputData[0]))

    fileNum = 0
    for sheet in inputData:
        print('---\n'
              'Copying comments from file: ' + filenames[fileNum])

        # Grab only the salesperson's data.
        sales = filenames[fileNum][0:2]
        sheetData = sheet[sheet['Sales'] == sales]
        # Append data to the output dataframe.
        finalData = finalData.append(sheetData, ignore_index=True)
        # Next file.
        fileNum += 1

    # Drop any unnamed columns that got processed.
    try:
        finalData = finalData.loc[:, ~sheet.columns.str.contains('^Unnamed')]
        finalData = finalData.loc[:, list(inputData[0])]
    except AttributeError:
        pass

    # Try saving the files, exit with error if any file is currently open.
    currentTime = time.strftime('%Y-%m-%d')
    fname1 = 'Digikey Insight Final ' + currentTime + '.xlsx'
    if saveError(fname1):
        print('---\n'
              'Insight Master is currently open in Excel!\n'
              'Please close the file and try again.\n'
              '***')
        return

    # Write the Insight Master file.
    writer1 = pd.ExcelWriter(fname1, engine='xlsxwriter',
                             datetime_format='mm/dd/yyyy')
    finalData.to_excel(writer1, sheet_name='Master', index=False)
    # Format as table in Excel.
    tableFormat(finalData, 'Master', writer1)

    # Save the file.
    writer1.save()

    print('---\n'
          'Updates completed successfully!\n'
          '---\n'
          'Digikey Master updated.\n'
          '+++')
