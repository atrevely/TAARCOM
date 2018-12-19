import pandas as pd
import time


def tableFormat(sheetData, sheetName, wbook):
    """Formats the Excel output as a table with correct column formatting."""
    # Nothing to format (emtpy table), so return.
    if sheetData.shape[0] == 0:
        return
    sheet = wbook.sheets[sheetName]
    # Set document formatting.
    docFormat = wbook.book.add_format({'font': 'Calibri',
                                       'font_size': 11})
    acctFormat = wbook.book.add_format({'font': 'Calibri',
                                        'font_size': 11,
                                        'num_format': 44})
    commaFormat = wbook.book.add_format({'font': 'Calibri',
                                         'font_size': 11,
                                         'num_format': 3})
    pctFormat = wbook.book.add_format({'font': 'Calibri',
                                       'font_size': 11,
                                       'num_format': '0.0%'})
    dateFormat = wbook.book.add_format({'font': 'Calibri',
                                        'font_size': 11,
                                        'num_format': 14})
    # Format and fit each column.
    i = 0
    for col in sheetData.columns:
        # Match the correct formatting to each column.
        acctCols = ['Unit Price', 'Paid-On Revenue', 'Actual Comm Paid',
                    'Total NDS', 'Post-Split NDS', 'Cust Revenue YTD',
                    'Ext. Cost', 'Unit Cost', 'Total Commissions',
                    'Sales Commission', 'Invoiced Dollars']
        pctCols = ['Split Percentage', 'Commission Rate',
                   'Gross Rev Reduction', 'Shared Rev Tier Rate']
        coreCols = ['CM Sales', 'Design Sales', 'T-End Cust', 'T-Name',
                    'CM', 'Invoice Date']
        dateCols = ['Invoice Date', 'Paid Date', 'Sales Report Date',
                    'Date Added']
        if col in acctCols:
            formatting = acctFormat
        elif col in pctCols:
            formatting = pctFormat
        elif col in dateCols:
            formatting = dateFormat
        elif col == 'Quantity':
            formatting = commaFormat
        else:
            formatting = docFormat
        # Set column width and formatting.
        try:
            maxWidth = max(len(str(val)) for val in sheetData[col].values)
        except ValueError:
            maxWidth = 0
        # If column is one that always gets filled in, then keep it expanded.
        if col in coreCols:
            maxWidth = max(maxWidth, len(col), 10)
        # Don't let the columns get too wide.
        maxWidth = min(maxWidth, 50)
        # Extra space for '$' in accounting format.
        if col in acctCols or col == 'Invoiced Dollars':
            maxWidth += 2
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
    # Load up the current Running Commissions file.
    runningCom = pd.read_excel(runCom, 'Master').fillna('')
    filesProcessed = pd.read_excel(runCom, 'Files Processed').fillna('')

    # Grab all of the salespeople initials.
    salespeople = list(set().union(runningCom['CM Sales'].unique(),
                                   runningCom['Design Sales'].unique()))
    del salespeople[salespeople == '']

    # Select data that has not been reported yet.
    unrepComms = runningCom[runningCom['Sales Report Date'] == '']

    # Create the dataframe with the Sales Totals information.
    salesTot = pd.DataFrame(columns=['Salesperson', 'Invoiced Dollars',
                                     'Sales Commission'])
    # %%
    # Go through each salesperson and pull their data.
    print('Running reports...')
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
        split = CMWithDesign['CM Split']/100
        CMWithDesign['Sales Percent'] = split*100
        CMWithDesign['Sales Commission'] = split*CMWithDesign['Sales Commission']

        # Grab entries that are Design Sales only.
        designSales = unrepComms[[not x and y for x, y in zip(CM, Design)]]
        # Determine share of sales.
        designOnly = designSales[designSales['CM Sales'] == '']
        designOnly['Sales Percent'] = 100
        designWithCM = designSales[designSales['CM Sales'] != '']
        split = (100 - designWithCM['CM Split'])/100
        designWithCM['Sales Percent'] = split*100
        designWithCM['Sales Commission'] = split*designWithCM['Sales Commission']

        # Grab CM + Design Sales entries.
        dualSales = unrepComms[[x and y for x, y in zip(CM, Design)]]
        dualSales['Sales Percent'] = 100

        # Set report columns.
        reportCols = ['Salesperson', 'Sales Percent', 'Reported Customer',
                      'T-Name', 'CM', 'T-End Cust', 'Reported End Customer',
                      'Principal', 'Corrected Distributor', 'Invoice Number',
                      'Part Number', 'Quantity', 'Unit Price',
                      'Invoiced Dollars', 'Paid-On Revenue',
                      'Split Percentage', 'Commission Rate',
                      'Actual Comm Paid', 'Sales Commission',
                      'Quarter Shipped', 'Invoice Date', 'On/Offshore', 'City']

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
        # Make sure columns are numeric.
        finalReport['Invoiced Dollars'] = pd.to_numeric(finalReport['Invoiced Dollars'],
                                                        errors='coerce').fillna(0)
        finalReport['Sales Commission'] = pd.to_numeric(finalReport['Sales Commission'],
                                                        errors='coerce').fillna(0)
        # Total up the Invoiced Dollars and Sales Commission.
        reportTot = pd.DataFrame(columns=['Salesperson', 'Invoiced Dollars',
                                          'Sales Commission'], index=[0])
        reportTot['Salesperson'] = person
        reportTot['Invoiced Dollars'] = sum(finalReport['Invoiced Dollars'])
        reportTot['Sales Commission'] = sum(finalReport['Sales Commission'])
        # Append to Sales Totals.
        salesTot = salesTot.append(reportTot, ignore_index=True)

        # Build table of sales by principal.
        princTab = pd.DataFrame(columns=['Principal', 'Invoiced Dollars',
                                         'Sales Commission'])
        row = 0
        totInv = 0
        totComm = 0
        # Tally up Invoiced Dollars and Sales Commission for each principal.
        for principal in finalReport['Principal'].unique():
            princSales = finalReport[finalReport['Principal'] == principal]
            princInv = sum(princSales['Invoiced Dollars'])
            princComm = sum(princSales['Sales Commission'])
            totInv += princInv
            totComm += princComm
            # Fill in table with principal's totals.
            princTab.loc[row, 'Principal'] = principal
            princTab.loc[row, 'Invoiced Dollars'] = princInv
            princTab.loc[row, 'Sales Commission'] = princComm
            row += 1
        # Fill in overall totals.
        princTab.loc[row, 'Principal'] = 'Grand Total'
        princTab.loc[row, 'Invoiced Dollars'] = totInv
        princTab.loc[row, 'Sales Commission'] = totComm

        # Build table of sales by customer.
        custTab = pd.DataFrame(columns=['Customer', 'Principal',
                                        'Invoiced Dollars',
                                        'Sales Commission'])
        row = 0
        for customer in finalReport['T-End Cust'].unique():
            custSales = finalReport[finalReport['T-End Cust'] == customer]
            custInv = sum(custSales['Invoiced Dollars'])
            custComm = sum(custSales['Sales Commission'])
            custTab.loc[row, 'Customer'] = customer
            custTab.loc[row, 'Invoiced Dollars'] = custInv
            custTab.loc[row, 'Sales Commission'] = custComm
            # Filter for each principal within a customer.
            row += 1
            for principal in custSales['Principal'].unique():
                custSub = custSales[custSales['Principal'] == principal]
                subInv = sum(custSub['Invoiced Dollars'])
                subComm = sum(custSub['Sales Commission'])
                custTab.loc[row, 'Principal'] = principal
                custTab.loc[row, 'Invoiced Dollars'] = subInv
                custTab.loc[row, 'Sales Commission'] = subComm
                row += 1

        # Write report to file.
        writer = pd.ExcelWriter(person + ' Sales Report - '
                                + time.strftime(' %Y-%m-%d')
                                + '.xlsx', engine='xlsxwriter',
                                datetime_format='mm/dd/yyyy')
        princTab.to_excel(writer, sheet_name='Principals', index=False)
        custTab.to_excel(writer, sheet_name='Customers', index=False)
        finalReport.to_excel(writer, sheet_name='Report Data', index=False)
        # Format as table in Excel.
        tableFormat(princTab, 'Principals', writer)
        tableFormat(custTab, 'Customers', writer)
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
    # Sum the sales totals into a grand total.
    grandTot = pd.DataFrame(columns=['Salesperson', 'Invoiced Dollars',
                                     'Sales Commission'], index=[0])
    grandTot['Salesperson'] = 'Grand total'
    grandTot['Invoiced Dollars'] = sum(salesTot['Invoiced Dollars'])
    grandTot['Sales Commission'] = sum(salesTot['Sales Commission'])
    # Append the grand totals to Sales Totals.
    salesTot = salesTot.append(grandTot, ignore_index=True)
    # Save the Running Commissions with entered report date.
    writer1 = pd.ExcelWriter('Running Commissions '
                             + time.strftime('%Y-%m-%d') + ' Reported'
                             + '.xlsx', engine='xlsxwriter',
                             datetime_format='mm/dd/yyyy')
    runningCom.to_excel(writer1, sheet_name='Master', index=False)
    filesProcessed.to_excel(writer1, sheet_name='Files Processed',
                            index=False)
    salesTot.to_excel(writer1, sheet_name='Sales Totals',
                      index=False)
    # Format as table in Excel.
    tableFormat(runningCom, 'Master', writer1)
    tableFormat(filesProcessed, 'Files Processed', writer1)
    tableFormat(salesTot, 'Sales Totals', writer1)

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
    print('---\n'
          'Reports completed successfully!\n'
          '+++')
