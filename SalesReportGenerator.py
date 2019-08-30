import pandas as pd
import numpy as np
import time
import pythoncom
import datetime
from dateutil.parser import parse
from RCExcelTools import tableFormat, formDate, saveError
from xlrd import XLRDError
import win32com.client
import os
import re
import sys
import shutil


# The main function.
def main(runCom):
    """Generates sales reports, then appends the Running Commissions data
    to the Commissions Master.

    If runCom is not supplied, then no new data is read and appended;
    reports are run instead on the data for the most recent quarter
    in Commissions Master.
    """
    # Set the directory for the data input/output.
    dataDir = 'Z:/MK Working Commissions/'
    lookDir = 'Z:/Commissions Lookup/'

    # Call this for multithreading using win32com, for some reason.
    pythoncom.CoInitialize()

    print('Loading the data from Commissions Master...')

    # -----------------------------------------------------------------------
    # Read in Salespeople Info. Terminate if not found or if errors in file.
    # -----------------------------------------------------------------------
    if os.path.exists(lookDir + 'Salespeople Info.xlsx'):
        try:
            salesInfo = pd.read_excel(lookDir + 'Salespeople Info.xlsx',
                                      'Info')
        except XLRDError:
            print('---\n'
                  'Error reading sheet name for Salespeople Info.xlsx!\n'
                  'Please make sure the main tab is named Info.\n'
                  '*Program terminated*')
            return
    else:
        print('---\n'
              'No Salespeople Info file found!\n'
              'Please make sure Salespeople Info.xlsx is in the directory.\n'
              '*Program terminated*')
        return

    # ----------------------------------------------
    # Load and prepare the Commissions Master file.
    # ----------------------------------------------
    try:
        comMast = pd.read_excel(dataDir + 'Commissions Master.xlsx',
                                'Master Data', dtype=str)
        masterFiles = pd.read_excel(dataDir + 'Commissions Master.xlsx',
                                    'Files Processed').fillna('')
    except FileNotFoundError:
        print('No Commissions Master file found!\n'
              '*Program Terminated*')
        return
    except XLRDError:
        print('Commissions Master tab names incorrect!\n'
              'Make sure the tabs are named Master Data and Files Processed.\n'
              '*Program Terminated*')
        return
    # Convert applicable columns to numeric.
    numCols = ['Quantity', 'Ext. Cost', 'Invoiced Dollars',
               'Paid-On Revenue', 'Actual Comm Paid', 'Unit Cost',
               'Unit Price', 'CM Split', 'Year', 'Sales Commission',
               'Split Percentage', 'Commission Rate',
               'Gross Rev Reduction', 'Shared Rev Tier Rate']
    for col in numCols:
        try:
            comMast[col] = pd.to_numeric(comMast[col],
                                         errors='coerce').fillna(0)
        except KeyError:
            pass
    # Convert individual numbers to numeric in rest of columns.
    mixedCols = [col for col in list(comMast) if col not in numCols]
    # Invoice/part numbers sometimes has leading zeros we'd like to keep.
    mixedCols.remove('Invoice Number')
    mixedCols.remove('Part Number')
    # The INF gets read in as infinity, so skip the principal column.
    mixedCols.remove('Principal')
    for col in mixedCols:
        comMast[col] = comMast[col].map(
                lambda x: pd.to_numeric(x, errors='ignore'))
    # Now remove the nans.
    comMast.replace(['nan', np.nan], '', inplace=True)
    # Make sure all the dates are formatted correctly.
    for col in ['Invoice Date', 'Paid Date', 'Sales Report Date']:
        comMast[col] = comMast[col].map(lambda x: formDate(x))
    # Make sure that the CM Splits aren't blank or zero.
    comMast['CM Split'] = comMast['CM Split'].replace(['', '0', 0], 20)
    # Column list.
    colAppend = list(comMast)

    # ------------------------------------------------------------
    # Load and prepare the Running Commissions file, if supplied.
    # ------------------------------------------------------------
    if runCom:
        try:
            runningCom = pd.read_excel(runCom, 'Master', dtype=str)
            filesProcessed = pd.read_excel(runCom,
                                           'Files Processed').fillna('')
        except FileNotFoundError:
            print('No Running Commissions file found!\n'
                  '*Program Terminated*')
            return
        except XLRDError:
            print('Running Commissions tab names incorrect!\n'
                  'Make sure the tabs are named Master and Files Processed.\n'
                  '*Program Terminated*')
            return
        for col in numCols:
            try:
                runningCom[col] = pd.to_numeric(runningCom[col],
                                                errors='coerce').fillna('')
            except KeyError:
                pass
        for col in mixedCols:
            runningCom[col] = runningCom[col].map(
                    lambda x: pd.to_numeric(x, errors='ignore'))
        # Now remove the nans.
        runningCom.replace(['nan', np.nan], '', inplace=True)
        # Make sure all the dates are formatted correctly.
        runningCom['Invoice Date'] = runningCom['Invoice Date'].map(
                lambda x: formDate(x))
        # Make sure that the CM Splits aren't blank or zero.
        runningCom['CM Split'] = runningCom['CM Split'].replace(['', '0', 0],
                                                                20)
        # Strip any extra spaces that made their way into salespeople columns.
        for col in ['CM Sales', 'Design Sales']:
            runningCom[col] = runningCom[col].map(lambda x: x.strip())

        # ---------------------------------------------
        # Fill in the Sales Commission in the RC file.
        # ---------------------------------------------
        for row in runningCom.index:
            # Get the CM and Design salespeople percentages.
            CMSales = runningCom.loc[row, 'CM Sales']
            DesignSales = runningCom.loc[row, 'Design Sales']
            # Deal with the QQ lines.
            if 'QQ' in (CMSales, DesignSales):
                salesComm = 0.45*runningCom.loc[row, 'Actual Comm Paid']
                runningCom.loc[row, 'Sales Commission'] = salesComm
                continue
            CM = salesInfo[salesInfo['Sales Initials'] == CMSales]
            design = salesInfo[salesInfo['Sales Initials'] == DesignSales]
            CMpct = CM['Sales Percentage']/100
            designPct = design['Sales Percentage']/100
            # Calculate the total sales commission
            if CMSales and DesignSales:
                CMpct *= runningCom.loc[row, 'CM Split']
                designPct *= 100 - runningCom.loc[row, 'CM Split']
                totPct = (CMpct.iloc[0] + designPct.iloc[0])/100
            else:
                totPct = [i.iloc[0] for i in (CMpct, designPct) if any(i)][0]
            salesComm = totPct*runningCom.loc[row, 'Actual Comm Paid']
            runningCom.loc[row, 'Sales Commission'] = salesComm
    else:
            print('Running reports without new Running Commissions...')

    # ------------------------------------------------------------------
    # Determine the commission months that are currently in the Master.
    # ------------------------------------------------------------------
    commMonths = comMast['Comm Month'].unique()
    try:
        commMonths = [parse(str(i).strip()) for i in commMonths if i != '']
    except ValueError:
        print('Error parsing dates in Comm Month column of Commissions Master!'
              '\nPlease check that all dates are in standard formatting and '
              'try again.\n*Program Terminated*')
        return
    # Grab the most recent month in Commissions Master.
    lastMonth = max(commMonths)
    if runCom:
        # Increment the month.
        currentMonth = lastMonth.month + 1
        currentYear = lastMonth.year
        # If current month is over 12, then it's time to go to January.
        if currentMonth > 12:
            currentMonth = 1
            currentYear += 1
        # Tag the new data as the current month/year.
        currentYrMo = str(currentYear) + '-' + str(currentMonth)
        runningCom['Comm Month'] = currentYrMo
    else:
        currentMonth = lastMonth.month
        currentYear = lastMonth.year
        currentYrMo = str(currentYear) + '-' + str(currentMonth)

    # ----------------------------------------
    # Load and prepare the Account List file.
    # ----------------------------------------
    try:
        acctList = pd.read_excel(lookDir + 'Master Account List.xlsx',
                                 'Allacct')
    except FileNotFoundError:
        print('No Account List file found!\n'
              '*Program Terminated*')
        return
    except XLRDError:
        print('Account List tab names incorrect!\n'
              'Make sure the main tab is named Allacct.\n'
              '*Program Terminated*')
        return

    if runCom:
        # ------------------------------------
        # Load and prepare the Master Lookup.
        # ------------------------------------
        if os.path.exists(lookDir + 'Lookup Master - Current.xlsx'):
            masterLookup = pd.read_excel(lookDir + 'Lookup Master - '
                                         'Current.xlsx').fillna('')
            # Check the column names.
            lookupCols = ['CM Sales', 'Design Sales', 'CM Split',
                          'Reported Customer', 'CM', 'Part Number', 'T-Name',
                          'T-End Cust', 'Last Used', 'Principal', 'City',
                          'Date Added']
            missCols = [i for i in lookupCols if i not in list(masterLookup)]
            if missCols:
                print('The following columns were not detected in '
                      'Lookup Master.xlsx:\n%s' %
                      ', '.join(map(str, missCols))
                      + '\n*Program Terminated*')
                return
        else:
            print('---\n'
                  'No Lookup Master found!\n'
                  'Please make sure Lookup Master - Current.xlsx is '
                  'in the directory.\n'
                  '*Program Terminated*')
            return

    print('Preparing report data...')
    if runCom:
        # -------------------------------------------------------------------
        # Check to make sure new files aren't already in Commissions Master.
        # -------------------------------------------------------------------
        # Check if we've duplicated any files.
        filenames = masterFiles['Filename']
        duplicates = list(set(filenames).intersection(
                filesProcessed['Filename']))
        # Don't let duplicate files get processed.
        if duplicates:
            # Let us know we found duplictes and removed them.
            print('---\n'
                  'The following files are already in '
                  'Commissions Master:\n%s' %
                  ', '.join(map(str, duplicates)) + '\nPlease check '
                  'the files and try again.\n*Program Terminated*')
            return

    # -----------------------------------------------------------------------
    # Combine and tag revenue data for the quarters that we're reporting on.
    # -----------------------------------------------------------------------
    # Grab the quarters in Commissions Master and Running Commissions.
    quarters = comMast['Quarter Shipped'].unique()
    if runCom:
        runComQuarters = runningCom['Quarter Shipped'].unique()
        quarters = list(set().union(quarters, runComQuarters))
    quarters.sort()
    # Use the most recent five quarters of data.
    quarters = quarters[-5:]
    # Get the revenue report data ready.
    revDat = comMast[[i in quarters for i in comMast['Quarter Shipped']]]
    revDat.reset_index(drop=True, inplace=True)
    if runCom:
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
        # If no design sales, use CM sales.
        if not revDat.loc[row, 'CDS']:
            revDat.loc[row, 'CDS'] = revDat.loc[row, 'CM Sales'].strip()
    # Also grab the section of the data that aren't 80/20 splits.
    splitDat = revDat[revDat['CM Split'] > 20]

    # --------------------------------------------------------
    # Combine and tag commission data for the current quarter.
    # --------------------------------------------------------
    # Figure out what slice of commissions data is in the current quarter.
    comMastTracked = comMast[comMast['Comm Month'] != '']
    try:
        commDates = comMastTracked['Comm Month'].map(lambda x: parse(str(x)))
    except (TypeError, ValueError):
        print('Error reading month in Comm Month column!\n'
              'Please make sure all months are in YYYY-MM format.\n'
              '*Program Terminated*')
        return
    # Filter to data in this year.
    yearData = comMastTracked[commDates.map(lambda x: x.year) == currentYear]
    # Determine how many months back we need to go.
    numPrevMos = (currentMonth - 1) % 3
    months = range(currentMonth, currentMonth - numPrevMos - 1, -1)
    dataMos = yearData['Comm Month'].map(lambda x: parse(str(x)).month)
    qtrData = yearData[dataMos.isin(list(months))]
    # Compile the commissions data.
    if runCom:
        commData = qtrData.append(runningCom, ignore_index=True, sort=False)
    else:
        commData = qtrData

    # ---------------------------------------
    # Get the salespeople information ready.
    # ---------------------------------------
    # Grab all of the salespeople initials.
    if runCom:
        salespeople = list(set().union(runningCom['CM Sales'].unique(),
                                       runningCom['Design Sales'].unique()))
    else:
        salespeople = list(set().union(revDat['CM Sales'].unique(),
                                       revDat['Design Sales'].unique()))
    salespeople = [i for i in salespeople if i not in ['QQ', '']]
    salespeople.sort()
    # Create the dataframe with the commission information by salesperson.
    salesTot = pd.DataFrame(columns=['Salesperson', 'Principal',
                                     'Paid-On Revenue', 'Actual Comm Paid',
                                     'Sales Commission', 'Comm Pct'],
                            index=[0])

    # ----------------------------------
    # Open Excel using the win32c tools.
    # ----------------------------------
    try:
        excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    except AttributeError:
        # Need to shut down the old module, or something. Windows sucks.
        module_list = [m.__name__ for m in sys.modules.values()]
        for module in module_list:
            if re.match(r'win32com\.gen_py\..+', module):
                del sys.modules[module]
        shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp',
                                   'gen_py'))
        # Now try again.
        excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
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
        # Also grab any nonstandard splits.
        cmDat = splitDat[splitDat['CM Sales'] == person]
        cmDat = cmDat[cmDat['CDS'] != person]
        designDat = designDat.append(cmDat, ignore_index=True, sort=False)
        # Get rid of empty Quarter Shipped lines.
        designDat = designDat[designDat['Quarter Shipped'] != '']
        designDat.reset_index(drop=True, inplace=True)
        # Replace zeros with blanks.
        for col in numCols:
            try:
                designDat[col].replace(0, '', inplace=True)
            except KeyError:
                pass
        # Write the raw data to a file.
        filename = (person + ' Revenue Report - ' + currentYrMo + '.xlsx')
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
                  '*Program Terminated*')
            return
        # Create the workbook and add the report sheet.
        wb = excel.Workbooks.Open(os.getcwd() + '\\' + filename)
        wb.Sheets.Add()
        pivotSheet = wb.Worksheets(1)
        pivotSheet.Name = 'Revenue Report'
        dataSheet = wb.Worksheets('Raw Data')
        # Grab the report data by selecting the current region.
        dataRange = dataSheet.Range('A1').CurrentRegion
        pivotRange = pivotSheet.Range('A1')
        # Create the pivot table and deploy it on the sheet.
        pivCache = wb.PivotCaches().Create(
                SourceType=win32c.xlDatabase,
                SourceData=dataRange,
                Version=win32c.xlPivotTableVersion14)
        pivTable = pivCache.CreatePivotTable(
                TableDestination=pivotRange,
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
        dataField = pivTable.AddDataField(
                pivTable.PivotFields('Paid-On Revenue'),
                'Revenue', win32c.xlSum)
        dataField.NumberFormat = '$#,##0'
        wb.Close(SaveChanges=1)

        # ---------------------------------------------------------------------
        # Create the commissions reports for each salesperson, using all data.
        # ---------------------------------------------------------------------
        # Determine the salesperson's commission percentage.
        sales = salesInfo[salesInfo['Sales Initials'] == person]
        commPct = sales['Sales Percentage'].iloc[0]/100
        # Find sales entries for the salesperson.
        CM = commData['CM Sales'] == person
        Design = commData['Design Sales'] == person
        # Grab entries that are CM Sales for this salesperson.
        CMSales = commData[[x and not y for x, y in zip(CM, Design)]]
        if not CMSales.empty:
            # Determine share of sales.
            CMOnly = CMSales[CMSales['Design Sales'] == '']
            CMOnly['Sales Commission'] = commPct*CMOnly['Actual Comm Paid']
            CMWithDesign = CMSales[CMSales['Design Sales'] != '']
            if not CMWithDesign.empty:
                try:
                    split = CMWithDesign['CM Split']/100
                except TypeError:
                    split = 0.2
                # Need to calculate sales commission from start for these.
                actComm = split*CMWithDesign['Actual Comm Paid']
                CMWithDesign['Actual Comm Paid'] = actComm
                salesComm = commPct*actComm
                CMWithDesign['Sales Commission'] = salesComm
        else:
            CMOnly = pd.DataFrame(columns=colAppend)
            CMWithDesign = pd.DataFrame(columns=colAppend)
        # Grab entries that are Design Sales for this salesperson.
        designSales = commData[[not x and y for x, y in zip(CM, Design)]]
        if not designSales.empty:
            # Determine share of sales.
            designOnly = designSales[designSales['CM Sales'] == '']
            desSalesComm = commPct*designOnly['Actual Comm Paid']
            designOnly['Sales Commission'] = desSalesComm
            designWithCM = designSales[designSales['CM Sales'] != '']
            if not designWithCM.empty:
                try:
                    split = (100 - designWithCM['CM Split'])/100
                except TypeError:
                    split = 0.8
                # Need to calculate sales commission from start for these.
                actComm = split*designWithCM['Actual Comm Paid']
                designWithCM['Actual Comm Paid'] = actComm
                salesComm = commPct*actComm
                designWithCM['Sales Commission'] = salesComm
        else:
            designOnly = pd.DataFrame(columns=colAppend)
            designWithCM = pd.DataFrame(columns=colAppend)
        # Grab CM + Design Sales entries.
        dualSales = commData[[x and y for x, y in zip(CM, Design)]]
        dualSalesComm = commPct*dualSales['Actual Comm Paid']
        dualSales['Sales Commission'] = dualSalesComm
        if dualSales.empty:
            dualSales = pd.DataFrame(columns=colAppend)

        # -----------------------------------------------
        # Grab the QQ entries and combine into one line.
        # -----------------------------------------------
        qqDat = commData[commData['Design Sales'] == 'QQ']
        qqCondensed = pd.DataFrame(columns=colAppend)
        qqCondensed.loc[0, 'T-End Cust'] = 'MISC POOL'
        qqCondensed.loc[0, 'Sales Commission'] = sum(qqDat['Sales Commission'])
        qqCondensed.loc[0, 'Design Sales'] = 'QQ'
        qqCondensed.loc[0, 'Principal'] = 'VARIOUS (MISC POOL)'
        qqCondensed.loc[0, 'Comm Month'] = currentYrMo
        qqCondensed.loc[0, 'Actual Comm Paid'] = sum(qqDat['Actual Comm Paid'])
        # Scale down the QQ entries based on the salesperson's share.
        QQperson = salesInfo[salesInfo['Sales Initials'] == person]
        try:
            QQscale = QQperson['QQ Split'].iloc[0]
            qqCondensed.loc[0, 'Sales Commission'] *= QQscale/100
        except IndexError:
            # No salesperson QQ split found, so empty it out.
            qqCondensed = pd.DataFrame(columns=colAppend)

        # -----------------------
        # Start creating report.
        # -----------------------
        finalReport = pd.DataFrame(columns=colAppend)
        # Append the data.
        finalReport = finalReport.append([CMOnly[colAppend],
                                          CMWithDesign[colAppend],
                                          designOnly[colAppend],
                                          designWithCM[colAppend],
                                          dualSales[colAppend],
                                          qqCondensed[colAppend]],
                                         ignore_index=True, sort=False)
        # Make sure columns are numeric.
        finalReport['Paid-On Revenue'] = pd.to_numeric(
                finalReport['Paid-On Revenue'], errors='coerce').fillna(0)
        finalReport['Actual Comm Paid'] = pd.to_numeric(
                finalReport['Actual Comm Paid'], errors='coerce').fillna(0)
        finalReport['Sales Commission'] = pd.to_numeric(
                finalReport['Sales Commission'], errors='coerce').fillna(0)
        # Total up the Paid-On Revenue and Actual/Sales Commission.
        reportTot = pd.DataFrame(columns=['Salesperson', 'Paid-On Revenue',
                                          'Actual Comm Paid',
                                          'Sales Commission', 'Comm Pct'],
                                 index=[0])
        reportTot['Salesperson'] = person
        reportTot['Principal'] = ''
        actComm = sum(finalReport['Actual Comm Paid'])
        salesComm = sum(finalReport['Sales Commission'])
        reportTot['Paid-On Revenue'] = sum(finalReport['Paid-On Revenue'])
        reportTot['Actual Comm Paid'] = actComm
        reportTot['Sales Commission'] = salesComm
        reportTot['Comm Pct'] = salesComm/actComm
        # Append to Sales Totals.
        salesTot = salesTot.append(reportTot, ignore_index=True, sort=False)
        # Build table of sales by principal.
        princTab = pd.DataFrame(columns=['Principal', 'Paid-On Revenue',
                                         'Sales Commission'])
        row = 0
        totInv = 0
        totAct = 0
        totComm = 0
        # Tally up Paid-On Revenue and Sales Commission for each principal.
        for principal in finalReport['Principal'].unique():
            princSales = finalReport[finalReport['Principal'] == principal]
            princInv = sum(princSales['Paid-On Revenue'])
            princAct = sum(princSales['Actual Comm Paid'])
            princComm = sum(princSales['Sales Commission'])
            totInv += princInv
            totAct += princAct
            totComm += princComm
            # Fill in table with principal's totals.
            princTab.loc[row, 'Principal'] = principal
            princTab.loc[row, 'Paid-On Revenue'] = princInv
            princTab.loc[row, 'Actual Comm Paid'] = princAct
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
        princTab.loc[row, 'Actual Comm Paid'] = totAct
        princTab.loc[row, 'Sales Commission'] = totComm
        princTab.loc[row, 'Comm Pct'] = totComm/totAct
        # Replace zeros with blanks.
        for col in numCols:
            try:
                finalReport[col].replace(0, '', inplace=True)
            except KeyError:
                pass
        # Write report to file.
        filename = (person + ' Commission Report - ' + currentYrMo + '.xlsx')
        writer = pd.ExcelWriter(filename, engine='xlsxwriter',
                                datetime_format='mm/dd/yyyy')
        princTab.to_excel(writer, sheet_name='Principals', index=False)
        finalReport.to_excel(writer, sheet_name='Raw Data', index=False)
        # Format as table in Excel.
        tableFormat(princTab, 'Principals', writer)
        tableFormat(finalReport, 'Raw Data', writer)
        # Try saving the file, exit with error if file is currently open.
        try:
            writer.save()
        except IOError:
            print('---\n'
                  'A salesperson report file is open in Excel!\n'
                  'Please close the file(s) and try again.\n'
                  '*Program Terminated*')
            return
        # Create the workbook and add the report sheet.
        wb = excel.Workbooks.Open(os.getcwd() + '\\' + filename)
        Principals = wb.Worksheets(1)
        wb.Sheets.Add(After=Principals)
        pivotSheet = wb.Worksheets(2)
        pivotSheet.Name = 'Commission Report'
        dataSheet = wb.Worksheets('Raw Data')
        # Grab the report data by selecting the current region.
        dataRange = dataSheet.Range('A1').CurrentRegion
        pivotRange = pivotSheet.Range('A1')
        # Create the pivot table and deploy it on the sheet.
        pivCache = wb.PivotCaches().Create(
                SourceType=win32c.xlDatabase,
                SourceData=dataRange,
                Version=win32c.xlPivotTableVersion14)
        pivTable = pivCache.CreatePivotTable(
                TableDestination=pivotRange,
                TableName='Commission Data',
                DefaultVersion=win32c.xlPivotTableVersion14)
        # Drop the data fields into the pivot table.
        pivTable.PivotFields('T-End Cust').Orientation = win32c.xlRowField
        pivTable.PivotFields('T-End Cust').Position = 1
        pivTable.PivotFields('Principal').Orientation = win32c.xlRowField
        pivTable.PivotFields('Principal').Position = 2
        pivTable.PivotFields('Comm Month').Orientation = win32c.xlColumnField
        # Add the sum of Sales Commissions as the data field.
        dataField = pivTable.AddDataField(
                pivTable.PivotFields('Sales Commission'),
                'Sales Comm', win32c.xlSum)
        dataField.NumberFormat = '$#,##0'
        wb.Close(SaveChanges=1)
    # %%
    # Fill in the Sales Report Date in Running Commissions.
    if runCom:
        runningCom.loc[runningCom['Sales Report Date'] == '',
                       'Sales Report Date'] = time.strftime('%m/%d/%Y')
        # ------------------------------------------------------
        # Create the tabs for the reported Running Commissions.
        # ------------------------------------------------------
        # Generate the table for sales numbers by principal.
        princTab = pd.DataFrame(columns=['Principal', 'Paid-On Revenue',
                                         'Actual Comm Paid',
                                         'Sales Commission'])
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
    filename = ('Revenue Report - ' + currentYrMo + '.xlsx')
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
              '*Program Terminated*')
        return
    # -------------------------------------------
    # Add the pivot table for revenue by quarter.
    # -------------------------------------------
    # Create the workbook and add the report sheet.
    wb = excel.Workbooks.Open(os.getcwd() + '\\' + filename)
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
    pivTable = pivCache.CreatePivotTable(
            TableDestination=pivotRange,
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

    # ------------------------------------------------------------------------
    # Go through each line of the finished Running Commissions and use them to
    # update the Lookup Master.
    # ------------------------------------------------------------------------
    if runCom:
        # Don't copy over INDIVIDUAL, MISC, or ALLOWANCE.
        noCopy = ['INDIVIDUAL', 'UNKNOWN', 'ALLOWANCE']
        paredID = [i for i in runningCom.index
                   if not any(j in runningCom.loc[i, 'T-End Cust'].upper()
                              for j in noCopy)]
        for row in paredID:
            # First match reported customer.
            repCust = str(runningCom.loc[row, 'Reported Customer']).lower()
            POSCust = masterLookup['Reported Customer'].map(
                    lambda x: str(x).lower())
            custMatches = masterLookup[repCust == POSCust]
            # Now match part number.
            partNum = str(runningCom.loc[row, 'Part Number']).lower()
            PPN = masterLookup['Part Number'].map(lambda x: str(x).lower())
            fullMatches = custMatches[PPN == partNum]
            # Figure out if this entry is a duplicate of any existing entry.
            duplicate = False
            for matchID in fullMatches.index:
                matchCols = ['CM Sales', 'Design Sales', 'CM', 'T-Name',
                             'T-End Cust']
                duplicate = all(
                        fullMatches.loc[matchID, i] == runningCom.loc[row, i]
                        for i in matchCols)
                if duplicate:
                    break
            # If it's not an exact duplicate, add it to the Lookup Master.
            if not duplicate:
                lookupCols = ['CM Sales', 'Design Sales', 'CM Split', 'CM',
                              'T-Name', 'T-End Cust', 'Reported Customer',
                              'Principal', 'Part Number', 'City']
                newLookup = runningCom.loc[row, lookupCols]
                newLookup['Date Added'] = datetime.datetime.now().date()
                newLookup['Last Used'] = datetime.datetime.now().date()
                masterLookup = masterLookup.append(newLookup,
                                                   ignore_index=True)

        # --------------------------------------------------------------
        # Append the new Running Commissions to the Commissions Master.
        # --------------------------------------------------------------
        comMast = comMast.append(runningCom, ignore_index=True, sort=False)
        masterFiles = masterFiles.append(filesProcessed, ignore_index=True,
                                         sort=False)
        # Make sure all the dates are formatted correctly.
        comMast['Invoice Date'] = comMast['Invoice Date'].map(
                lambda x: formDate(x))
        masterFiles['Date Added'] = masterFiles['Date Added'].map(
                lambda x: formDate(x))
        masterFiles['Paid Date'] = masterFiles['Paid Date'].map(
                lambda x: formDate(x))
        # Convert commission dollars to numeric.
        masterFiles['Total Commissions'] = pd.to_numeric(
                masterFiles['Total Commissions'], errors='coerce').fillna('')
        # Convert applicable columns to numeric.
        numCols = ['Quantity', 'Ext. Cost', 'Invoiced Dollars',
                   'Paid-On Revenue', 'Actual Comm Paid', 'Unit Cost',
                   'Unit Price', 'Year', 'Sales Commission',
                   'Split Percentage', 'Commission Rate',
                   'Gross Rev Reduction', 'Shared Rev Tier Rate']
        for col in numCols:
            try:
                comMast[col] = pd.to_numeric(comMast[col],
                                             errors='coerce').fillna('')
            except KeyError:
                pass
        # Convert individual numbers to numeric in rest of columns.
        mixedCols = [col for col in list(comMast) if col not in numCols]
        # Invoice/part numbers sometimes has leading zeros we'd like to keep.
        mixedCols.remove('Invoice Number')
        mixedCols.remove('Part Number')
        # The INF gets read in as infinity, so skip the principal column.
        mixedCols.remove('Principal')
        for col in mixedCols:
            comMast[col] = comMast[col].map(
                    lambda x: pd.to_numeric(x, errors='ignore'))

        # %%
        # Save the files.
        fname1 = dataDir + 'Commissions Master.xlsx'
        fname2 = (dataDir + 'Running Commissions ' + currentYrMo
                  + ' Reported.xlsx')
        fname3 = lookDir + 'Lookup Master - Current.xlsx'

        if saveError(fname1, fname2, fname3):
            print('---\n'
                  'One or more of these files are currently open in Excel:\n'
                  'Running Commissions, Entries Need Fixing, Lookup Master.\n'
                  'Please close these files and try again.\n'
                  '*Program Terminated*')
            return
        # Write the Commissions Master file.
        writer = pd.ExcelWriter(fname1, engine='xlsxwriter',
                                datetime_format='mm/dd/yyyy')
        comMast.to_excel(writer, sheet_name='Master', index=False)
        masterFiles.to_excel(writer, sheet_name='Files Processed', index=False)
        # Format everything in Excel.
        tableFormat(comMast, 'Master Data', writer)
        tableFormat(masterFiles, 'Files Processed', writer)

        # Write the Running Commissions report.
        writer1 = pd.ExcelWriter(fname2, engine='xlsxwriter',
                                 datetime_format='mm/dd/yyyy')
        runningCom.to_excel(writer1, sheet_name='Master', index=False)
        filesProcessed.to_excel(writer1, sheet_name='Files Processed',
                                index=False)
        salesTot.to_excel(writer1, sheet_name='Salesperson Totals',
                          index=False)
        princTab.to_excel(writer1, sheet_name='Principal Totals',
                          index=False)
        # Format as table in Excel.
        tableFormat(runningCom, 'Master Data', writer1)
        tableFormat(filesProcessed, 'Files Processed', writer1)
        tableFormat(salesTot, 'Salesperson Totals', writer1)
        tableFormat(princTab, 'Principal Totals', writer1)

        # Write the Lookup Master.
        writer2 = pd.ExcelWriter(fname3, engine='xlsxwriter',
                                 datetime_format='mm/dd/yyyy')
        masterLookup.to_excel(writer2, sheet_name='Lookup', index=False)
        # Format everything in Excel.
        tableFormat(masterLookup, 'Lookup', writer2)

        # Save the files.
        writer.save()
        writer1.save()
        writer2.save()
        print('---\n'
              'Sales reports finished successfully!\n'
              '---\n'
              'Commissions Master updated.\n'
              'Lookup Master updated.\n'
              '+++')
    else:
        print('Reports finished successfully!')
    # Close the Excel instance.
    excel.Application.Quit()
