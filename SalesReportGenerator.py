import pandas as pd
import time


# The main function.
def main():
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
        sheet = writer.sheets['Lookup']
        # Format as table.
        header = [{'header': val} for val in finalReport.columns.tolist()]
        set = {'header_row': True, 'style': 'TableStyleMedium5',
               'columns': header}
        sheet.add_table(0, 0, len(finalReport.index),
                        len(finalReport.columns)-1, set)
        # Try saving the file, exit with error if file is currently open.
        try:
            writer.save()
        except IOError:
            print('---\n'
                  'File is open in Excel!\n'
                  'Please close the file and try again.\n'
                  '***')
            return
        writer.save()

    # Fill in the Sales Report Date.
    runningCom.loc[runningCom['Sales Report Date'] == '',
                   'Sales Report Date'] = time.strftime('%m/%d/%Y')
    # Save the Running Commissions with entered report date.
    writer1 = pd.ExcelWriter('Running Commissions '
                             + time.strftime('%Y-%m-%d-%H%M')
                             + '.xlsx', engine='xlsxwriter')
    runningCom.to_excel(writer1, sheet_name='Master', index=False)
    filesProcessed.to_excel(writer1, sheet_name='Files Processed',
                            index=False)
    sheet1a = writer1.sheets['Master']
    sheet1b = writer1.sheets['Files Processed']
    # Format as table.
    header1a = [{'header': val} for val in runningCom.columns.tolist()]
    header1b = [{'header': val} for val in filesProcessed.columns.tolist()]
    set1a = {'header_row': True, 'style': 'TableStyleMedium5',
             'columns': header1a}
    set1b = {'header_row': True, 'style': 'TableStyleMedium5',
             'columns': header1b}
    sheet1a.add_table(0, 0, len(runningCom.index),
                      len(runningCom.columns)-1, set1a)
    sheet1b.add_table(0, 0, len(filesProcessed.index),
                      len(filesProcessed.columns)-1, set1b)
