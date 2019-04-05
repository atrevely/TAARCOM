import pandas as pd
import time
from RCExcelTools import tableFormat, formDate
from xlrd import XLRDError
import win32com.client
import os


# The main function.
def main(runCom):
    """Generates sales reports.

    Finds entries in Running Commissions filters them into reports for each
    salesperson, as well as an overall report.
    """
    # ----------------------------------------------
    # Load and prepare the Running Commissions file.
    # ----------------------------------------------
    # Load up the current Running Commissions file.
    try:
        runningCom = pd.read_excel(runCom, 'Master', dtype=str)
        filesProcessed = pd.read_excel(runCom, 'Files Processed').fillna('')
    except FileNotFoundError:
        print('No Running Commissions file found!\n'
              '***')
        return
    except XLRDError:
        print('Running Commissions tab names incorrect!\n'
              'Make sure the tabs are named Master and Files Processed.\n'
              '***')
        return

    # Convert applicable columns to numeric.
    numCols = ['Quantity', 'Ext. Cost', 'Invoiced Dollars', 'Paid-On Revenue',
               'Actual Comm Paid', 'Unit Cost', 'Unit Price', 'CM Split',
               'Year', 'Sales Commission', 'Split Percentage',
               'Commission Rate', 'Gross Rev Reduction',
               'Shared Rev Tier Rate']
    for col in numCols:
        try:
            runningCom[col] = pd.to_numeric(runningCom[col],
                                            errors='coerce').fillna(0)
        except KeyError:
            pass
    # Convert individual numbers to numeric in rest of columns.
    mixedCols = [col for col in list(runningCom) if col not in numCols]
    # Invoice/part numbers sometimes has leading zeros we'd like to keep.
    mixedCols.remove('Invoice Number')
    mixedCols.remove('Part Number')
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

    # ---------------------------------------------
    # Load and prepare the Commissions Master file.
    # ---------------------------------------------
    # Load up the current Commissions Master file.
    try:
        comMast = pd.read_excel('Commissions Master.xlsx', 'Master', dtype=str)
    except FileNotFoundError:
        print('No Commissions Master file found!\n'
              '***')
        return
    except XLRDError:
        print('Commissions Master tab names incorrect!\n'
              'Make sure the tabs are named Master and Files Processed.\n'
              '***')
        return
    # Convert applicable columns to numeric.
    for col in numCols:
        try:
            comMast[col] = pd.to_numeric(comMast[col],
                                         errors='coerce').fillna(0)
        except KeyError:
            pass
    for col in mixedCols:
        comMast[col] = comMast[col].map(
                lambda x: pd.to_numeric(x, errors='ignore'))
    # Now remove the nans.
    comMast.replace('nan', '', inplace=True)

    # Make sure all the dates are formatted correctly.
    comMast['Invoice Date'] = comMast['Invoice Date'].map(
            lambda x: formDate(x))

    # ---------------------------------------
    # Load and prepare the Account List file.
    # ---------------------------------------
    # Load up the Account List.
    try:
        acctList = pd.read_excel('Master Account List.xlsx', 'Allacct')
    except FileNotFoundError:
        print('No Account List file found!\n'
              '***')
        return
    except XLRDError:
        print('Account List tab names incorrect!\n'
              'Make sure the main tab is named Allacct.\n'
              '***')
        return

    print('Preparing report data...')
    # --------------------------------------
    # Get the salespeople information ready.
    # --------------------------------------
    # Grab all of the salespeople initials.
    salespeople = list(set().union(runningCom['CM Sales'].unique(),
                                   runningCom['Design Sales'].unique()))
    del salespeople[salespeople == '']
    salespeople.sort()

    # Create the dataframe with the commission information by salesperson.
    salesTot = pd.DataFrame(columns=['Salesperson', 'Principal',
                                     'Paid-On Revenue', 'Sales Commission'])

    # Columns appended from Running Commissions.
    colAppend = list(runningCom)

    # --------------------------------------------------------------------
    # Combine and tag data for the quarters that we're going to report on.
    # --------------------------------------------------------------------
    # Grab the quarters in Commission Master and Running Commissions.
    comMastQuarters = comMast['Quarter Shipped'].unique()
    runComQuarters = runningCom['Quarter Shipped'].unique()
    quarters = list(set().union(comMastQuarters, runComQuarters))
    quarters.sort()
    # Use the most recent five quarters of data.
    quarters = quarters[-5:]
    # Get the revenue report data ready.
    revDat = comMast[[i in quarters for i in comMast['Quarter Shipped']]]
    revDat.reset_index(drop=True, inplace=True)
    revDat = revDat.append(runningCom, ignore_index=True, sort=False)
    # Tag the data by current Design Sales.
    for cust in revDat['T-End Cust'].unique():
        # Check for a single match in Account List.
        if sum(acctList['ProperName'] == cust) == 1:
            sales = acctList[acctList['ProperName'] == cust]['SLS'].iloc[0]
            custID = revDat[revDat['T-End Cust'] == cust].index
            revDat.loc[custID, 'CDS'] = sales
    # Fill in the CDS (current design sales) for missing entries as simply the
    # Design Sales for that line.
    for row in revDat[pd.isna(revDat['CDS'])].index:
        revDat.loc[row, 'CDS'] = revDat.loc[row, 'Design Sales']

    # ----------------------------------
    # Open Excel using the win32c tools.
    # ----------------------------------
    Excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    win32c = win32com.client.constants

    # %%
    # Go through each salesperson and prepare their reports.
    print('Running reports...')
    for person in salespeople:
        # -----------------------------------------------------------
        # Create the revenue reports for each salesperson, using only
        # design data.
        # -----------------------------------------------------------
        # Grab the raw data for this salesperson's design sales.
        designDat = revDat[revDat['CDS'] == person]
        # Get rid of empty Quarter Shipped lines.
        designDat = designDat[designDat['Quarter Shipped'] != '']
        designDat.reset_index(drop=True, inplace=True)
        # Write the raw data to a file.
        filename = (person + ' Revenue Report - ' + time.strftime('%Y-%m-%d')
                    + '.xlsx')
        writer = pd.ExcelWriter(filename, engine='xlsxwriter',
                                datetime_format='mm/dd/yyyy')
        designDat.to_excel(writer, sheet_name='Raw Data', index=False)
        tableFormat(designDat, 'Raw Data', writer)
        # Try saving the report.
        try:
            writer.save()
        except IOError:
            print('---\n'
                  'A salesperson report file is open in Excel!\n'
                  'Please close the file(s) and try again.\n'
                  '***')
            return
        # Create the workbook and add the report sheet.
        wb = Excel.Workbooks.Open(os.getcwd() + '\\' + filename)
        wb.Sheets.Add()
        pivotSheet = wb.Worksheets(1)
        pivotSheet.Name = 'Revenue Report'
        dataSheet = wb.Worksheets('Raw Data')
        # Grab the report data by selecting the current region.
        dataRange = dataSheet.Range('A1').CurrentRegion
        pivotRange = pivotSheet.Range('A1')
        # Create the pivot table and deploy it on the sheet.
        pivCache = wb.PivotCaches().Create(SourceType=win32c.xlDatabase,
                                           SourceData=dataRange,
                                           Version=win32c.xlPivotTableVersion14)
        pivTable = pivCache.CreatePivotTable(TableDestination=pivotRange,
                                             TableName='Revenue Data',
                                             DefaultVersion=win32c.xlPivotTableVersion14)
        # Drop the data fields into the pivot table.
        pivTable.PivotFields('T-End Cust').Orientation = win32c.xlRowField
        pivTable.PivotFields('T-End Cust').Position = 1
        pivTable.PivotFields('Part Number').Orientation = win32c.xlRowField
        pivTable.PivotFields('Part Number').Position = 2        
        pivTable.PivotFields('CM').Orientation = win32c.xlRowField
        pivTable.PivotFields('CM').Position = 3
        pivTable.PivotFields('Quarter Shipped').Orientation = win32c.xlColumnField
        pivTable.PivotFields('Principal').Orientation = win32c.xlPageField
        # Add the sum of Paid-On Revenue as the data field.
        dataField = pivTable.AddDataField(pivTable.PivotFields('Paid-On Revenue'),
                                            'Revenue', win32c.xlSum)
        dataField.NumberFormat = '$#,##0'
        wb.Close(SaveChanges=1)

        # --------------------------------------------------------------
        # Create the sales reports for each salesperson, using all data.
        # --------------------------------------------------------------
        # Find sales entries for the salesperson.
        CM = runningCom['CM Sales'] == person
        Design = runningCom['Design Sales'] == person
        # Grab entries that are CM Sales for this salesperson.
        CMSales = runningCom[[x and not y for x, y in zip(CM, Design)]]
        if CMSales.shape[0]:
            # Determine share of sales.
            CMOnly = CMSales[CMSales['Design Sales'] == '']
            CMOnly['Sales Percent'] = 100
            CMWithDesign = CMSales[CMSales['Design Sales'] != '']
            try:
                split = CMWithDesign['CM Split']/100
            except TypeError:
                split = 0.2
            CMWithDesign['Sales Percent'] = split*100
            CMWithDesign['Sales Commission'] *= split
        else:
            CMOnly = pd.DataFrame(columns=colAppend)
            CMWithDesign = pd.DataFrame(columns=colAppend)

        # Grab entries that are Design Sales only.
        designSales = runningCom[[not x and y for x, y in zip(CM, Design)]]
        if designSales.shape[0]:
            # Determine share of sales.
            designOnly = designSales[designSales['CM Sales'] == '']
            designOnly['Sales Percent'] = 100
            designWithCM = designSales[designSales['CM Sales'] != '']
            try:
                split = (100 - designWithCM['CM Split'])/100
            except TypeError:
                split = 0.8
            designWithCM['Sales Percent'] = split*100
            designWithCM['Sales Commission'] *= split
        else:
            designOnly = pd.DataFrame(columns=colAppend)
            designWithCM = pd.DataFrame(columns=colAppend)

        # Grab CM + Design Sales entries.
        dualSales = runningCom[[x and y for x, y in zip(CM, Design)]]
        if dualSales.shape[0]:
            dualSales['Sales Percent'] = 100
        else:
            dualSales = pd.DataFrame(columns=colAppend)

        # Start creating report.
        finalReport = pd.DataFrame(columns=colAppend)
        # Append the data.
        finalReport = finalReport.append([CMOnly[colAppend],
                                          CMWithDesign[colAppend],
                                          designOnly[colAppend],
                                          designWithCM[colAppend],
                                          dualSales[colAppend]],
                                         ignore_index=True, sort=False)
        # Make sure columns are numeric.
        finalReport['Paid-On Revenue'] = pd.to_numeric(
                finalReport['Paid-On Revenue'], errors='coerce').fillna(0)
        finalReport['Sales Commission'] = pd.to_numeric(
                finalReport['Sales Commission'], errors='coerce').fillna(0)
        # Total up the Paid-On Revenue and Sales Commission.
        reportTot = pd.DataFrame(columns=['Salesperson', 'Paid-On Revenue',
                                          'Sales Commission'], index=[0])
        reportTot['Salesperson'] = person
        reportTot['Principal'] = ''
        reportTot['Paid-On Revenue'] = sum(finalReport['Paid-On Revenue'])
        reportTot['Sales Commission'] = sum(finalReport['Sales Commission'])
        # Append to Sales Totals.
        salesTot = salesTot.append(reportTot, ignore_index=True, sort=False)

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

        # Sort principals in descending order alphabetically.
        princTab.sort_values(by=['Principal'], inplace=True)
        princTab.reset_index(drop=True, inplace=True)
        # Append to Salesperson Totals tab.
        salesTot = salesTot.append(princTab, ignore_index=True, sort=False)

        # Sort principals in descending order by Sales Commission.
        princTab.sort_values(by=['Sales Commission'], ascending=False,
                             inplace=True)
        princTab.reset_index(drop=True, inplace=True)
        # Fill in overall totals.
        princTab.loc[row, 'Principal'] = 'Grand Total'
        princTab.loc[row, 'Paid-On Revenue'] = totInv
        princTab.loc[row, 'Sales Commission'] = totComm

        # Build table of sales by customer.
        custTab = pd.DataFrame(columns=['Customer', 'Principal',
                                        'Paid-On Revenue',
                                        'Sales Commission'])
        finalCusts = pd.DataFrame(columns=['Customer', 'Principal',
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
            row += 1
        # Sort customers in descending order by Sales Commission.
        custTab.sort_values(by=['Sales Commission'], ascending=False,
                            inplace=True)
        # Add in subtotals by principal for each customer.
        row = 0
        for customer in custTab['Customer'].unique():
            finalCusts.loc[row, :] = custTab[custTab['Customer'] == customer].iloc[0]
            custSales = finalReport[finalReport['T-End Cust'] == customer]
            row += 1
            for principal in custSales['Principal'].unique():
                custSub = custSales[custSales['Principal'] == principal]
                subInv = sum(custSub['Paid-On Revenue'])
                subComm = sum(custSub['Sales Commission'])
                finalCusts.loc[row, 'Principal'] = principal
                finalCusts.loc[row, 'Paid-On Revenue'] = subInv
                finalCusts.loc[row, 'Sales Commission'] = subComm
                row += 1

        # Write report to file.
        writer = pd.ExcelWriter(person + ' Sales Report - '
                                + time.strftime('%Y-%m-%d')
                                + '.xlsx', engine='xlsxwriter',
                                datetime_format='mm/dd/yyyy')
        princTab.to_excel(writer, sheet_name='Principals', index=False)
        finalCusts.to_excel(writer, sheet_name='Customers', index=False)
        finalReport.to_excel(writer, sheet_name='Report Data', index=False)
        # Format as table in Excel.
        tableFormat(princTab, 'Principals', writer)
        tableFormat(finalCusts, 'Customers', writer)
        tableFormat(finalReport, 'Report Data', writer)

        # Try saving the file, exit with error if file is currently open.
        try:
            writer.save()
        except IOError:
            print('---\n'
                  'A salesperson report file is open in Excel!\n'
                  'Please close the file(s) and try again.\n'
                  '***')
            return

    # %%
    # Fill in the Sales Report Date in Running Commissions.
    runningCom.loc[runningCom['Sales Report Date'] == '',
                   'Sales Report Date'] = time.strftime('%m/%d/%Y')

    # -----------------------------------------------------
    # Create the tabs for the reported Running Commissions.
    # -----------------------------------------------------
    # Generate the table for sales numbers by principal.
    princTab = pd.DataFrame(columns=['Principal', 'Paid-On Revenue',
                                     'Actual Comm Paid', 'Sales Commission'])
    row = 0
    for principal in runningCom['Principal'].unique():
        princSales = runningCom[runningCom['Principal'] == principal]
        revenue = pd.to_numeric(princSales['Paid-On Revenue'],
                                errors='coerce').fillna(0)
        comm = pd.to_numeric(princSales['Actual Comm Paid'],
                             errors='coerce').fillna(0)
        salesComm = pd.to_numeric(princSales['Sales Commission'],
                                  errors='coerce').fillna(0)
        princInv = sum(revenue)
        actComm = sum(comm)
        salesComm = sum(salesComm)
        try:
            trueCommPct = actComm/princInv
        except ZeroDivisionError:
            trueCommPct = ''
        # Fill in table with principal's totals.
        princTab.loc[row, 'Principal'] = principal
        princTab.loc[row, 'Paid-On Revenue'] = princInv
        princTab.loc[row, 'Actual Comm Paid'] = actComm
        princTab.loc[row, 'True Comm %'] = trueCommPct
        princTab.loc[row, 'Sales Commission'] = salesComm
        row += 1
    # Sort principals in descending order alphabetically.
    princTab.sort_values(by=['Principal'], inplace=True)
    princTab.reset_index(drop=True, inplace=True)
    # Fill in overall totals.
    totRev = sum(princTab['Paid-On Revenue'])
    totSalesComm = sum(princTab['Sales Commission'])
    totComm = sum(princTab['Actual Comm Paid'])
    princTab.loc[row, 'Paid-On Revenue'] = totRev
    princTab.loc[row, 'Sales Commission'] = totSalesComm
    princTab.loc[row, 'Actual Comm Paid'] = totComm
    princTab.loc[row, 'Principal'] = 'Grand Total'

    # -----------------------------------
    # Create the overall Revenue Report.
    # -----------------------------------
    # Write the raw data to a file.
    filename = ('Revenue Report - ' + time.strftime('%Y-%m-%d') + '.xlsx')
    writer = pd.ExcelWriter(filename, engine='xlsxwriter',
                            datetime_format='mm/dd/yyyy')
    revDat.to_excel(writer, sheet_name='Raw Data', index=False)
    tableFormat(designDat, 'Raw Data', writer)
    # Try saving the report.
    try:
        writer.save()
    except IOError:
        print('---\n'
              'Revenue report file is open in Excel!\n'
              'Please close the file(s) and try again.\n'
              '***')
        return
    # Create the workbook and add the report sheet.
    wb = Excel.Workbooks.Open(os.getcwd() + '\\' + filename)
    wb.Sheets.Add()
    pivotSheet = wb.Worksheets(1)
    pivotSheet.Name = 'Revenue Report'
    dataSheet = wb.Worksheets('Raw Data')
    # Grab the report data by selecting the current region.
    dataRange = dataSheet.Range('A1').CurrentRegion
    pivotRange = pivotSheet.Range('A1')
    # Create the pivot table and deploy it on the sheet.
    pivCache = wb.PivotCaches().Create(SourceType=win32c.xlDatabase,
                                       SourceData=dataRange,
                                       Version=win32c.xlPivotTableVersion14)
    pivTable = pivCache.CreatePivotTable(TableDestination=pivotRange,
                                         TableName='Revenue Data',
                                         DefaultVersion=win32c.xlPivotTableVersion14)
    # Drop the data fields into the pivot table.
    pivTable.PivotFields('T-End Cust').Orientation = win32c.xlRowField
    pivTable.PivotFields('T-End Cust').Position = 1
    pivTable.PivotFields('CM').Orientation = win32c.xlRowField
    pivTable.PivotFields('CM').Position = 2
    pivTable.PivotFields('Part Number').Orientation = win32c.xlRowField
    pivTable.PivotFields('Part Number').Position = 3
    pivTable.PivotFields('Quarter Shipped').Orientation = win32c.xlColumnField
    pivTable.PivotFields('Principal').Orientation = win32c.xlPageField
    # Add the sum of Paid-On Revenue as the data field.
    dataField = pivTable.AddDataField(pivTable.PivotFields('Paid-On Revenue'),
                                      'Revenue', win32c.xlSum)
    dataField.NumberFormat = '$#,##0'
    wb.Close(SaveChanges=1)

    # ------------------------------------------------------
    # Save the Running Commissions with entered report date.
    # ------------------------------------------------------
    writer1 = pd.ExcelWriter('Running Commissions '
                             + time.strftime('%Y-%m-%d') + ' Reported'
                             + '.xlsx', engine='xlsxwriter',
                             datetime_format='mm/dd/yyyy')
    runningCom.to_excel(writer1, sheet_name='Master', index=False)
    filesProcessed.to_excel(writer1, sheet_name='Files Processed',
                            index=False)
    salesTot.to_excel(writer1, sheet_name='Salesperson Totals',
                      index=False)
    princTab.to_excel(writer1, sheet_name='Principal Totals',
                      index=False)
    # Format as table in Excel.
    tableFormat(runningCom, 'Master', writer1)
    tableFormat(filesProcessed, 'Files Processed', writer1)
    tableFormat(salesTot, 'Salesperson Totals', writer1)
    tableFormat(princTab, 'Principal Totals', writer1)

    # Try saving the file, exit with error if file is currently open.
    try:
        writer1.save()
    except IOError:
        print('---\n'
              'Final report file is open in Excel!\n'
              'Please close the file and try again.\n'
              '***')
        return
    print('---\n'
          'Reports completed successfully!\n'
          '+++')
    # Close the Excel instance.
    Excel.Application.Quit()
