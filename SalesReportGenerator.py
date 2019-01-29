import pandas as pd
import time
from dateutil.parser import parse


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
    index = 0
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
        elif col == 'Invoice Number':
            # We're going to do some work in order to format the Invoice
            # Number as a number, yet keep leading zeros.
            for row in sheetData.index:
                invLen = len(sheetData.loc[row, 'Invoice Number'])
                # Figure out how many places the number goes to.
                numPadding = '0'*invLen
                invNum = pd.to_numeric(sheetData.loc[row, 'Invoice Number'],
                                       errors='ignore')
                invFormat = wbook.book.add_format({'font': 'Calibri',
                                                   'font_size': 11,
                                                   'num_format': numPadding})
                try:
                    sheet.write_number(row+1, index, invNum, invFormat)
                except TypeError:
                    pass
            # Move to the next column, as we're now done formatting
            # the Invoice Numbers.
            index += 1
            continue
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
        if col in acctCols:
            maxWidth += 2
        sheet.set_column(index, index, maxWidth+0.8, formatting)
        index += 1


def formDate(inputDate):
    """Attemps to format a string as a date, otherwise ignores it."""
    try:
        outputDate = parse(str(inputDate)).date()
        return outputDate
    except ValueError:
        return inputDate


# The main function.
def main(runCom):
    """Generates sales reports for each salesperson.

    Finds entries in Running Commissions that are marked as currently
    unreported and filters them into reports for each salesperson. Entries are
    marked as reported in Running Commissions after being assigned to a report
    by this function.
    """
    # Load up the current Running Commissions file.
    runningCom = pd.read_excel(runCom, 'Master', dtype=str).fillna('')
    filesProcessed = pd.read_excel(runCom, 'Files Processed').fillna('')

    # Convert applicable columns to numeric.
    numCols = ['Quantity', 'Ext. Cost', 'Invoiced Dollars', 'Paid-On Revenue',
               'Actual Comm Paid', 'Unit Cost', 'Unit Price', 'CM Split',
               'Year', 'Sales Commission', 'Split Percentage',
               'Commission Rate', 'Gross Rev Reduction',
               'Shared Rev Tier Rate']
    for col in numCols:
        try:
            runningCom[col] = pd.to_numeric(runningCom[col],
                                            errors='coerce').fillna('')
        except KeyError:
            pass
    # Convert individual numbers to numeric in rest of columns.
    mixedCols = [col for col in list(runningCom) if col not in numCols]
    # Invoice number sometimes has leading zeros we'd like to keep.
    mixedCols.remove('Invoice Number')
    # The INF gets read in as infinity, so skip the principal column.
    mixedCols.remove('Principal')
    for col in mixedCols:
        runningCom[col] = runningCom[col].map(
                lambda x: pd.to_numeric(x, errors='ignore'))
    # Now remove the nans.
    runningCom.replace('nan', '', inplace=True)

    # Make sure all the dates are formatted correctly.
    runningCom['Invoice Date'] = runningCom['Invoice Date'].map(
            lambda x: formDate(x))

    # Grab all of the salespeople initials.
    salespeople = list(set().union(runningCom['CM Sales'].unique(),
                                   runningCom['Design Sales'].unique()))
    del salespeople[salespeople == '']

    # Select data that has not been reported yet.
    unrepComms = runningCom[runningCom['Sales Report Date'] == '']

    # Create the dataframe with the Sales Totals information.
    salesTot = pd.DataFrame(columns=['Salesperson', 'Paid-On Revenue',
                                     'Sales Commission'])

    # Set report columns.
    reportCols = ['Salesperson', 'Sales Percent', 'Reported Customer',
                  'T-Name', 'CM', 'T-End Cust', 'Reported End Customer',
                  'Principal', 'Corrected Distributor', 'Invoice Number',
                  'Part Number', 'Quantity', 'Unit Price',
                  'Invoiced Dollars', 'Paid-On Revenue',
                  'Split Percentage', 'Commission Rate',
                  'Actual Comm Paid', 'Sales Commission', 'Comm Source',
                  'Quarter Shipped', 'Invoice Date', 'On/Offshore', 'City']
    # Columns appended from Running Commissions.
    colAppend = [val for val in reportCols if val in list(unrepComms)]
    # %%
    # Go through each salesperson and pull their data.
    print('Running reports...')
    for person in salespeople:
        # Find sales entries for the salesperson.
        CM = unrepComms['CM Sales'] == person
        Design = unrepComms['Design Sales'] == person

        # Grab entries that are CM Sales for this salesperson.
        CMSales = unrepComms[[x and not y for x, y in zip(CM, Design)]]
        if CMSales.shape[0]:
            # Determine share of sales.
            CMOnly = CMSales[CMSales['Design Sales'] == '']
            CMOnly['Sales Percent'] = 100
            CMWithDesign = CMSales[CMSales['Design Sales'] != '']
            try:
                split = CMWithDesign['CM Split']/100
            except TypeError:
                split = 0.8
            CMWithDesign['Sales Percent'] = split*100
            CMWithDesign['Sales Commission'] = split*CMWithDesign['Sales Commission']
        else:
            CMOnly = pd.DataFrame(columns=colAppend)
            CMWithDesign = pd.DataFrame(columns=colAppend)

        # Grab entries that are Design Sales only.
        designSales = unrepComms[[not x and y for x, y in zip(CM, Design)]]
        if designSales.shape[0]:
            # Determine share of sales.
            designOnly = designSales[designSales['CM Sales'] == '']
            designOnly['Sales Percent'] = 100
            designWithCM = designSales[designSales['CM Sales'] != '']
            try:
                split = (100 - designWithCM['CM Split'])/100
            except TypeError:
                split = 0.2
            designWithCM['Sales Percent'] = split*100
            designWithCM['Sales Commission'] = split*designWithCM['Sales Commission']
        else:
            designOnly = pd.DataFrame(columns=colAppend)
            designWithCM = pd.DataFrame(columns=colAppend)

        # Grab CM + Design Sales entries.
        dualSales = unrepComms[[x and y for x, y in zip(CM, Design)]]
        if dualSales.shape[0]:
            dualSales['Sales Percent'] = 100
        else:
            dualSales = pd.DataFrame(columns=colAppend)

        # Start creating report.
        finalReport = pd.DataFrame(columns=reportCols)
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
        finalReport['Paid-On Revenue'] = pd.to_numeric(
                finalReport['Paid-On Revenue'], errors='coerce').fillna(0)
        finalReport['Sales Commission'] = pd.to_numeric(
                finalReport['Sales Commission'], errors='coerce').fillna(0)
        # Total up the Paid-On Revenue and Sales Commission.
        reportTot = pd.DataFrame(columns=['Salesperson', 'Paid-On Revenue',
                                          'Sales Commission'], index=[0])
        reportTot['Salesperson'] = person
        reportTot['Paid-On Revenue'] = sum(finalReport['Paid-On Revenue'])
        reportTot['Sales Commission'] = sum(finalReport['Sales Commission'])
        # Append to Sales Totals.
        salesTot = salesTot.append(reportTot, ignore_index=True)

        # Build table of sales by principal.
        princTab = pd.DataFrame(columns=['Principal', 'Paid-On Revenue',
                                         'Sales Commission'])
        row = 0
        totInv = 0
        totComm = 0
        # Tally up Paid-On Revenue and Sales Commission for each principal.
        for principal in finalReport['Principal'].unique():
            princSales = finalReport[finalReport['Principal'] == principal]
            princInv = sum(princSales['Paid-On Revenue'])
            princComm = sum(princSales['Sales Commission'])
            totInv += princInv
            totComm += princComm
            # Fill in table with principal's totals.
            princTab.loc[row, 'Principal'] = principal
            princTab.loc[row, 'Paid-On Revenue'] = princInv
            princTab.loc[row, 'Sales Commission'] = princComm
            row += 1
        # Fill in overall totals.
        princTab.loc[row, 'Principal'] = 'Grand Total'
        princTab.loc[row, 'Paid-On Revenue'] = totInv
        princTab.loc[row, 'Sales Commission'] = totComm

        # Build table of sales by customer.
        custTab = pd.DataFrame(columns=['Customer', 'Principal',
                                        'Paid-On Revenue',
                                        'Sales Commission'])
        row = 0
        for customer in finalReport['T-End Cust'].unique():
            custSales = finalReport[finalReport['T-End Cust'] == customer]
            custInv = sum(custSales['Paid-On Revenue'])
            custComm = sum(custSales['Sales Commission'])
            custTab.loc[row, 'Customer'] = customer
            custTab.loc[row, 'Paid-On Revenue'] = custInv
            custTab.loc[row, 'Sales Commission'] = custComm
            # Filter for each principal within a customer.
            row += 1
            for principal in custSales['Principal'].unique():
                custSub = custSales[custSales['Principal'] == principal]
                subInv = sum(custSub['Paid-On Revenue'])
                subComm = sum(custSub['Sales Commission'])
                custTab.loc[row, 'Principal'] = principal
                custTab.loc[row, 'Paid-On Revenue'] = subInv
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

    # Generate the table for sales numbers by principal.
    princTab = pd.DataFrame(columns=['Principal', 'Paid-On Revenue',
                            'Sales Commission'])
    row = 0
    for principal in unrepComms['Principal'].unique():
        princSales = unrepComms[unrepComms['Principal'] == principal]
        revenue = pd.to_numeric(princSales['Paid-On Revenue'],
                                errors='coerce').fillna(0)
        comm = pd.to_numeric(princSales['Sales Commission'],
                             errors='coerce').fillna(0)
        princInv = sum(revenue)
        princComm = sum(comm)
        # Fill in table with principal's totals.
        princTab.loc[row, 'Principal'] = principal
        princTab.loc[row, 'Paid-On Revenue'] = princInv
        princTab.loc[row, 'Sales Commission'] = princComm
        row += 1
    # Fill in overall totals.
    totRev = sum(princTab['Paid-On Revenue'])
    totComm = sum(princTab['Sales Commission'])
    princTab.loc[row, 'Paid-On Revenue'] = totRev
    princTab.loc[row, 'Sales Commission'] = totComm
    princTab.loc[row, 'Principal'] = 'Grand Total'

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
    princTab.to_excel(writer1, sheet_name='Principal Totals',
                      index=False)
    # Format as table in Excel.
    tableFormat(runningCom, 'Master', writer1)
    tableFormat(filesProcessed, 'Files Processed', writer1)
    tableFormat(salesTot, 'Sales Totals', writer1)
    tableFormat(princTab, 'Principal Totals', writer1)

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
