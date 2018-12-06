import pandas as pd
import time


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
    acctFormat = wbook.book.add_format({'font': 'Calibri',
                                        'font_size': 11,
                                        'num_format': 44})
    commaFormat = wbook.book.add_format({'font': 'Calibri',
                                         'font_size': 11,
                                         'num_format': 3})
    # Format and fit each column.
    i = 0
    for col in sheetData.columns:
        # Match the correct formatting to each column.
        acctCols = ['Unit Price', 'Invoiced Dollars', 'Paid-On Revenue',
                    'Actual Comm Paid', 'Total NDS', 'Post-Split NDS',
                    'Customer Revenue YTD']
        if col in acctCols:
            formatting = acctFormat
        elif col == 'Quantity':
            formatting = commaFormat
        else:
            formatting = docFormat
        # Set column width and formatting.
        maxWidth = max(len(str(val)) for val in sheetData[col].values)
        sheet.set_column(i, i, maxWidth+0.8, formatting)
        i += 1


# The main function.
def main(runCom):
    """Generates sales reports for each salesperson.

    Finds entries in Running Commissions that are marked as currently
    unreported and filters them into reports for each salesperson. Entries are
    marked as reported in Running Commissions after being assigned to a report
    by this function.
    """
    # Load up the Running Master.
    runningCom = pd.read_excel('Running Commissions 2018-06-29-1347.xlsx',
                               'Master').fillna('')
    filesProcessed = pd.read_excel('Running Commissions 2018-06-29-1347.xlsx',
                                   'Files Processed').fillna('')

    # Grab all of the salespeople initials.
    salespeople = list(set().union(runningCom['CM Sales'].unique(),
                                   runningCom['Design Sales'].unique()))
    del salespeople[salespeople == '']

    # Select data that has not been reported yet.
    unrepComms = runningCom[runningCom['Sales Report Date'] == '']

    # %%
    # Go through each salesperson and pull their data.
    for person in salespeople:
        # Find sales entries for the salesperson.
        CM = unrepComms['CM Sales'] == person
        Design = unrepComms['Design Sales'] == person

        # Grab entries that are CM Sales for this salesperson.
        CMSales = unrepComms[[x and not y for x, y in zip(CM, Design)]]
        # Determine share of sales.
        CMOnly = CMSales[CMSales['Design Sales'] == '']
        CMOnly['Sales Percent'] = 100
        CMWithDesign = CMSales[CMSales['Design Sales'] != '']
        CMWithDesign['Sales Percent'] = CMWithDesign['CM Split']

        # Grab entries that are Design Sales only.
        designSales = unrepComms[[not x and y for x, y in zip(CM, Design)]]
        # Determine share of sales.
        designOnly = designSales[designSales['CM Sales'] == '']
        designOnly['Sales Percent'] = 100
        designWithCM = designSales[designSales['CM Sales'] != '']
        designWithCM['Sales Percent'] = 100 - designWithCM['CM Split']

        # Grab CM + Design Sales entries.
        dualSales = unrepComms[[x and y for x, y in zip(CM, Design)]]
        dualSales['Sales Percent'] = 100

        # Set report columns.
        reportCols = ['Salesperson', 'Sales Percent', 'T-End Cust', 'T-Name',
                      'CM', 'Reported Customer', 'Reported End Customer',
                      'Principal', 'Corrected Distributor', 'Invoice Number',
                      'Part Number', 'Quantity', 'Unit Price',
                      'Invoiced Dollars', 'Paid-On Revenue',
                      'Split Percentage', 'Commission Rate',
                      'Actual Comm Paid', 'Invoice Date', 'On/Offshore',
                      'City', 'Cust Part Number']

        # Start creating report.
        finalReport = pd.DataFrame(columns=reportCols)
        # Columns appended from Running Commissions.
        colAppend = [val for val in reportCols if val in list(dualSales)]
        # Append the data.
        finalReport = finalReport.append([CMOnly[colAppend],
                                          CMWithDesign[colAppend],
                                          designOnly[colAppend],
                                          designWithCM[colAppend],
                                          dualSales[colAppend]],
                                         ignore_index=True)

        # Fill in salesperson initials.
        finalReport['Salesperson'] = person

        # Reorder columns.
        finalReport = finalReport.loc[:, reportCols]

        # Write report to file.
        writer = pd.ExcelWriter('Sales Report - ' + person
                                + time.strftime(' %Y-%m-%d')
                                + '.xlsx', engine='xlsxwriter')
        finalReport.to_excel(writer, sheet_name='Report Data', index=False)
        # Format as table in Excel.
        tableFormat(finalReport, 'Report Data', writer)

        # Try saving the file, exit with error if file is currently open.
        try:
            writer.save()
        except IOError:
            print('---\n'
                  'File is open in Excel!\n'
                  'Please close the file and try again.\n'
                  '***')
            return
        # No errors, so save the file.
        writer.save()

    # %%
    # Fill in the Sales Report Date in Running Commissions.
    runningCom.loc[runningCom['Sales Report Date'] == '',
                   'Sales Report Date'] = time.strftime('%m/%d/%Y')
    # Save the Running Commissions with entered report date.
    writer1 = pd.ExcelWriter('Running Commissions '
                             + time.strftime('%Y-%m-%d-%H%M')
                             + '.xlsx', engine='xlsxwriter')
    runningCom.to_excel(writer1, sheet_name='Master', index=False)
    filesProcessed.to_excel(writer1, sheet_name='Files Processed',
                            index=False)
    # Format as table in Excel.
    tableFormat(runningCom, 'Master', writer1)
    tableFormat(filesProcessed, 'Files Processed', writer1)

    # Try saving the file, exit with error if file is currently open.
    try:
        writer1.save()
    except IOError:
        print('---\n'
              'File is open in Excel!\n'
              'Please close the file and try again.\n'
              '***')
        return
    # No errors, so save the file.
    writer1.save()
