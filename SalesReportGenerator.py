import pandas as pd
import numpy as np
import time
import datetime
import os
import re
import sys
import shutil
from dateutil.parser import parse
from RCExcelTools import table_format, form_date, save_error, PivotTables
from FileLoader import load_salespeople_info, load_com_master, load_run_com, \
                       load_acct_list
from xlrd import XLRDError
# from PDFReportGenerator import pdfReport


# Set the numerical columns.
num_cols = ['Quantity', 'Ext. Cost', 'Invoiced Dollars',
            'Paid-On Revenue', 'Actual Comm Paid', 'Unit Cost',
            'Unit Price', 'CM Split', 'Year', 'Sales Commission',
            'Split Percentage', 'Commission Rate',
            'Gross Rev Reduction', 'Shared Rev Tier Rate']

# Set the directory for the data input/output.
if os.path.exists('Z:\\'):
    data_dir = 'Z:\\MK Working Commissions'
    look_dir = 'Z:\\Commissions Lookup'
    reports_dir = 'Z:\\MK Working Commissions\\Reports'
    if not os.path.exists(reports_dir):
        try:
            os.mkdir(reports_dir)
        except OSError:
            print('Error creating Reports folder in MK Working Commissions.')
else:
    data_dir = os.getcwd()
    look_dir = os.getcwd()
    reports_dir = os.getcwd()


def get_sales_comm_data(salesperson, input_data, sales_info):
    """Returns all the data for a particular salesperson, with sales
    commission scaled down by split percentage."""
    # Determine the salesperson's commission percentage.
    sales = sales_info[sales_info['Sales Initials'] == salesperson]
    comm_pct = sales['Sales Percentage'].iloc[0] / 100
    # Grab the data that has either CM or design sales for this person.
    CM = input_data['CM Sales'] == salesperson
    design = input_data['Design Sales'] == salesperson
    sales_data = input_data[np.logical_or(CM, design)]
    # Scale the commission data by split percentage.
    shared_sales = np.logical_and(sales_data['CM Sales'] != '',
                                  sales_data['Design Sales'] != '')
    for row in shared_sales.index:
        if CM.loc[row] and not design.loc[row]:
            split = sales_data.loc[row, 'CM Split'] / 100
        elif design.loc[row] and not CM.loc[row]:
            split = 1 - sales_data.loc[row, 'CM Split'] / 100
        else:
            split = 1
        sales_data.loc[row, 'Actual Comm Paid'] *= split
        sales_data.loc[row, 'Sales Commission'] *= split
    sales_data.reset_index(drop=True, inplace=True)
    return sales_data


def data_by_princ_tab(input_data):
    """Builds an Excel tab of the provided data broken down by principal."""
    princ_tab = pd.DataFrame(columns=['Principal', 'Paid-On Revenue',
                                      'Sales Commission'])
    # Tally up Paid-On Revenue and Sales Commission for each principal.
    for row, principal in enumerate(input_data['Principal'].unique()):
        princ_sales = input_data[input_data['Principal'] == principal]
        princ_inv = sum(princ_sales['Paid-On Revenue'])
        princ_act = sum(princ_sales['Actual Comm Paid'])
        princ_comm = sum(princ_sales['Sales Commission'])
        # Fill in the table with this principal's totals.
        princ_tab.loc[row, 'Principal'] = principal
        princ_tab.loc[row, 'Paid-On Revenue'] = princ_inv
        princ_tab.loc[row, 'Actual Comm Paid'] = princ_act
        princ_tab.loc[row, 'Sales Commission'] = princ_comm
    # Sort principals in descending order alphabetically.
    princ_tab.sort_values(by=['Principal'], inplace=True)
    princ_tab.reset_index(drop=True, inplace=True)
    # Fill in overall totals.
    princ_tab.loc[row + 1, 'Principal'] = 'Grand Total'
    princ_tab.loc[row + 1, 'Paid-On Revenue'] = np.sum(princ_tab['Paid-On Revenue'])
    princ_tab.loc[row + 1, 'Actual Comm Paid'] = np.sum(princ_tab['Actual Comm Paid'])
    princ_tab.loc[row + 1, 'Sales Commission'] = np.sum(princ_tab['Sales Commission'])
    return princ_tab


def create_quarterly_report(comm_data, comm_qtr, salespeople, sales_info):
    """Builds the report that runs at the end of each quarter."""
    print('Creating end-of-quarter report.')
    # Build the commission by principal tab.
    princ_tab = data_by_princ_tab(input_data=comm_data)
    for princ_row in princ_tab.index:
        try:
            actual_comm = princ_tab.loc[princ_row, 'Actual Comm Paid']
            revenue = princ_tab.loc[princ_row, 'Paid-On Revenue']
            true_comm_pct = actual_comm / revenue
        except ZeroDivisionError:
            true_comm_pct = ''
        princ_tab.loc[princ_row, 'True Comm %'] = true_comm_pct
    # Build the commission by salesperson and principal tab.
    sales_tab = pd.DataFrame(columns=['Salesperson', 'Principal', 'Actual Comm Paid',
                                      'Sales Commission'])
    row = 0
    for person in salespeople:
        sales_data = get_sales_comm_data(salesperson=person, input_data=comm_data,
                                         sales_info=sales_info)
        sales_tab.loc[row, 'Salesperson'] = person
        sales_tab.loc[row, 'Actual Comm Paid'] = np.sum(sales_data['Actual Comm Paid'])
        sales_tab.loc[row, 'Sales Commission'] = np.sum(sales_data['Sales Commission'])
        row += 1
        for princ in sorted(sales_data['Principal'].unique()):
            sales_tab.loc[row, 'Principal'] = princ
            princ_data = sales_data[sales_data['Principal'] == princ]
            sales_tab.loc[row, 'Actual Comm Paid'] = np.sum(princ_data['Actual Comm Paid'])
            sales_tab.loc[row, 'Sales Commission'] = np.sum(princ_data['Sales Commission'])
            row += 1
    # Save the report.
    filename = (reports_dir + '\\Quarterly Commission Report '
                + comm_qtr + '.xlsx')
    writer = pd.ExcelWriter(filename, engine='xlsxwriter',
                            datetime_format='mm/dd/yyyy')
    comm_data.to_excel(writer, sheet_name='Comm Data', index=False)
    princ_tab.to_excel(writer, sheet_name='Principals', index=False)
    sales_tab.to_excel(writer, sheet_name='Salespeople', index=False)
    table_format(comm_data, 'Comm Data', writer)
    table_format(princ_tab, 'Principals', writer)
    table_format(sales_tab, 'Salespeople', writer)
    if not save_error(filename):
        writer.save()
    else:
        print('Error saving quarter commission report.')


# The main function.
def main(run_com):
    """Generates sales reports, then appends the Running Commissions data
    to the Commissions Master.

    If run_com is not supplied, then no new data is read/appended;
    reports are run instead on the data for the most recent month
    in Commissions Master.
    """
    # Create the pivot tables class instance.
    pivots = PivotTables()
    print('Loading the data from Commissions Master...')
    # --------------------------------------------------------
    # Load in the supporting files, exit if any aren't found.
    # --------------------------------------------------------
    sales_info = load_salespeople_info(file_dir=look_dir)
    acctList = load_acct_list(file_dir=look_dir)
    com_mast, master_files = load_com_master(file_dir=data_dir)
    if any([acctList.empty, sales_info.empty, com_mast.empty, master_files.empty]):
        print('*Program Terminated*')
        return
    # Grab the column list for use later.
    master_cols = list(com_mast)

    # ------------------------------------------------------------------
    # Determine the commission months that are currently in the Master.
    # ------------------------------------------------------------------
    comm_months = com_mast['Comm Month'].unique()
    try:
        comm_months = [parse(str(i).strip()) for i in comm_months if i != '']
    except ValueError:
        print('Error parsing dates in Comm Month column of Commissions Master!'
              '\nPlease check that all dates are in standard formatting and '
              'try again.\n*Program Terminated*')
        return
    # Grab the most recent month in Commissions Master.
    last_month = max(comm_months)

    if run_com:
        look_mast = load_lookup_master(file_path=look_dir)
        running_com, files_processed = load_run_com(file_path=run_com)
        if any([look_mast.empty, running_com.empty, files_processed.empty]):
            print('*Program Terminated*')
            return
        # Fill in the Sales Report Date in Running Commissions.
        running_com.loc[running_com['Sales Report Date'] == '',
                        'Sales Report Date'] = time.strftime('%m/%d/%Y')
        # -------------------------------------------------------------------
        # Check to make sure new files aren't already in Commissions Master.
        # -------------------------------------------------------------------
        # Check if we've duplicated any files.
        filenames = master_files['Filename']
        duplicates = list(set(filenames).intersection(
            files_processed['Filename']))
        # Don't let duplicate files get processed.
        if duplicates:
            # Let us know we found duplicates and removed them.
            print('---\n'
                  'The following files are already in '
                  'Commissions Master:\n%s' %
                  ', '.join(map(str, duplicates)) + '\nPlease check '
                  'the files and try again.\n*Program Terminated*')
            return

        # ---------------------------------------------
        # Fill in the Sales Commission in the RC file.
        # ---------------------------------------------
        for row in running_com.index:
            tot_pct = 0
            # Get the CM and Design salespeople percentages.
            cm_sales = running_com.loc[row, 'CM Sales']
            design_sales = running_com.loc[row, 'Design Sales']
            # Deal with the QQ lines.
            if 'QQ' in (cm_sales, design_sales):
                sales_comm = 0.45 * running_com.loc[row, 'Actual Comm Paid']
                running_com.loc[row, 'Sales Commission'] = sales_comm
                continue
            CM = sales_info[sales_info['Sales Initials'] == cm_sales]
            design = sales_info[sales_info['Sales Initials'] == design_sales]
            cm_pct = CM['Sales Percentage'] / 100
            design_pct = design['Sales Percentage'] / 100
            # Calculate the total sales commission.
            if cm_sales and design_sales:
                try:
                    cm_pct *= running_com.loc[row, 'CM Split']
                    design_pct *= 100 - running_com.loc[row, 'CM Split']
                    tot_pct = (cm_pct.iloc[0] + design_pct.iloc[0]) / 100
                except IndexError:
                    print('Error finding sales percentages on line '
                          + str(row + 2) + ' in Running Commissions.')
            else:
                try:
                    tot_pct = [i.iloc[0] for i in (cm_pct, design_pct) if any(i)][0]
                except IndexError:
                    print('No salesperson found on line ' + str(row + 2)
                          + ' in Running Commissions.')
            if tot_pct:
                sales_comm = tot_pct * running_com.loc[row, 'Actual Comm Paid']
                running_com.loc[row, 'Sales Commission'] = sales_comm

        # ---------------------------------------------------------
        # Calculate the new commission month we're adding from RC.
        # ---------------------------------------------------------
        current_month = last_month.month + 1
        current_year = last_month.year
        # If current month is over 12, then it's time to go to January.
        if current_month > 12:
            current_month = 1
            current_year += 1
        # Tag the new data as the current month/year.
        currentYrMo = str(current_year) + '-' + str(current_month)
        running_com['Comm Month'] = currentYrMo
        # No add-on since this is the first RC run.
        RC_addon = ''
    else:
        # -----------------------------------------------------------------
        # Use the most recent commission month as the RC from here on out.
        # -----------------------------------------------------------------
        current_month = last_month.month
        current_year = last_month.year
        currentYrMo = str(current_year) + '-' + str(current_month)
        running_com = com_mast[com_mast['Comm Month'] == currentYrMo]
        # Indicate that this is a rerun.
        RC_addon = ' (Rerun)'
        print('No new RC supplied. Reporting on latest quarter in '
              'the Commissions Master.')

    print('Preparing report data...')
    # ------------------------------------------------------------------------
    # Combine and tag revenue data for the quarters that we're reporting on.
    # We report on the most recent 5 quarters of data for the Revenue Report.
    # ------------------------------------------------------------------------
    quarters = com_mast['Quarter Shipped'].unique()
    if run_com:
        run_com_quarters = running_com['Quarter Shipped'].unique()
        quarters = list(set().union(quarters, run_com_quarters))
    # Use the most recent five quarters of data.
    quarters = sorted(quarters)[-5:]
    # Get the revenue report data ready.
    revenue_data = com_mast[com_mast['Quarter Shipped'].isin(quarters)]
    revenue_data.reset_index(drop=True, inplace=True)
    if run_com:
        revenue_data = revenue_data.append(running_com, ignore_index=True, sort=False)
    # Tag the data by current Design Sales in the Account List.
    for cust in revenue_data['T-End Cust'].unique():
        # Check for a single match in Account List.
        try:
            if sum(acctList['ProperName'] == cust) == 1:
                sales = acctList[acctList['ProperName'] == cust]['SLS'].iloc[0]
                custID = revenue_data[revenue_data['T-End Cust'] == cust].index
                revenue_data.loc[custID, 'CDS'] = sales
        except KeyError:
            print('Error reading column names in Account List!\n'
                  'Please make sure the columns ProperName and SLS are in '
                  'the Account List.\n'
                  '*Program Terminated*')
            return
    # Fill in the CDS (current design sales) for missing entries as simply the
    # Design Sales for that line.
    for row in revenue_data[pd.isna(revenue_data['CDS'])].index:
        revenue_data.loc[row, 'CDS'] = revenue_data.loc[row, 'Design Sales']
        # If no design sales, use CM sales.
        if not revenue_data.loc[row, 'CDS']:
            revenue_data.loc[row, 'CDS'] = revenue_data.loc[row, 'CM Sales']
    # Also grab the section of the data that aren't 80/20 splits.
    splitDat = revenue_data[revenue_data['CM Split'] > 20]

    # ---------------------------------------------------------
    # Combine and tag commission data for the current quarter.
    # ---------------------------------------------------------
    # Figure out what slice of commissions data is in the current quarter.
    com_mast_tracked = com_mast[com_mast['Comm Month'] != '']
    try:
        com_mast_tracked['Comm Month'].map(lambda x: parse(str(x)))
    except (TypeError, ValueError):
        print('Error reading month in Comm Month column!\n'
              'Please make sure all months are in YYYY-MM format.\n'
              '*Program Terminated*')
        return
    # Determine how many months back we need to go.
    num_prev_mos = (current_month - 1) % 3
    months = range(current_month, current_month - num_prev_mos - 1, -1)
    qtr_mos = [str(current_year) + '-' + str(i) for i in months]
    qtr_data = com_mast_tracked[com_mast_tracked['Comm Month'].isin(qtr_mos)]
    # Compile the quarter data.
    if run_com:
        comm_data = qtr_data.append(running_com, ignore_index=True, sort=False)
    else:
        comm_data = qtr_data
    del qtr_data, com_mast_tracked

    # ---------------------------------------
    # Get the salespeople information ready.
    # ---------------------------------------
    # Grab all of the salespeople initials.
    salespeople = sorted(sales_info['Sales Initials'].values)
    print('Found the following sales initials in the Salespeople Info file: '
          + ', '.join(salespeople))
    # Create the dataframe with the commission information by salesperson.
    sales_tot = pd.DataFrame(columns=['Salesperson', 'Principal',
                                      'Paid-On Revenue', 'Actual Comm Paid',
                                      'Sales Commission'],
                             index=[0])

    # Go through each salesperson and prepare their reports.
    print('Running reports...')
    for person in salespeople:
        # ------------------------------------------------------------
        # Create the revenue reports for each salesperson, using only
        # design data.
        # ------------------------------------------------------------
        # Grab the raw data for this salesperson's design sales.
        designDat = revenue_data[revenue_data['CDS'] == person]
        # Also grab any nonstandard splits.
        cmDat = splitDat[splitDat['CM Sales'] == person]
        cmDat = cmDat[cmDat['CDS'] != person]
        designDat = designDat.append(cmDat, ignore_index=True, sort=False)
        # Get rid of empty Quarter Shipped lines.
        designDat = designDat[designDat['Quarter Shipped'] != '']
        designDat.reset_index(drop=True, inplace=True)
        # Replace zeros with blanks, except in commission columns.
        for col in set(num_cols).difference({'Actual Comm Paid', 'Sales Commission'}):
            try:
                designDat[col].replace(0, '', inplace=True)
            except KeyError:
                pass
        # Write the raw data to a file.
        filename = (reports_dir + '\\' + person + ' Revenue Report - '
                    + currentYrMo + '.xlsx')
        writer = pd.ExcelWriter(filename, engine='xlsxwriter',
                                datetime_format='mm/dd/yyyy')
        designDat.to_excel(writer, sheet_name='Raw Data', index=False)
        table_format(designDat, 'Raw Data', writer)
        # Try saving the report.
        try:
            writer.save()
        except IOError:
            print('---\n'
                  'A salesperson report file is open in Excel!\n'
                  'Please close the file(s) and try again.\n'
                  '*Program Terminated*')
            return
        # Create the revenue pivot table.
        pivots.create_pivot_table(excel_file=filename,
                                  data_sheet_name='Raw Data',
                                  pivot_sheet_name='Revenue Table',
                                  row_fields=['T-End Cust', 'Part Number', 'CM'],
                                  col_field='Quarter Shipped',
                                  data_field='Paid-On Revenue',
                                  page_field='Principal')

        # --------------------------------------------------------------------
        # Create the commission reports for each salesperson, using all data.
        # --------------------------------------------------------------------
        # Determine the salesperson's commission percentage.
        sales = sales_info[sales_info['Sales Initials'] == person]
        commPct = sales['Sales Percentage'].iloc[0] / 100
        # Find sales entries for the salesperson.
        CM = comm_data['CM Sales'] == person
        Design = comm_data['Design Sales'] == person
        # Grab entries that are CM Sales for this salesperson.
        cm_sales = comm_data[[x and not y for x, y in zip(CM, Design)]]
        if not cm_sales.empty:
            # Determine share of sales.
            CMOnly = cm_sales[cm_sales['Design Sales'] == '']
            CMOnly['Sales Commission'] = commPct * CMOnly['Actual Comm Paid']
            CMWithDesign = cm_sales[cm_sales['Design Sales'] != '']
            if not CMWithDesign.empty:
                split = CMWithDesign['CM Split'] / 100
                # Need to calculate sales commission from start for these.
                actComm = split * CMWithDesign['Actual Comm Paid']
                CMWithDesign['Actual Comm Paid'] = actComm
                sales_comm = commPct * actComm
                CMWithDesign['Sales Commission'] = sales_comm
        else:
            CMOnly = pd.DataFrame(columns=master_cols)
            CMWithDesign = pd.DataFrame(columns=master_cols)
        # Grab entries that are Design Sales for this salesperson.
        designSales = comm_data[[not x and y for x, y in zip(CM, Design)]]
        if not designSales.empty:
            # Determine share of sales.
            designOnly = designSales[designSales['CM Sales'] == '']
            desSalesComm = commPct * designOnly['Actual Comm Paid']
            designOnly['Sales Commission'] = desSalesComm
            designWithCM = designSales[designSales['CM Sales'] != '']
            if not designWithCM.empty:
                split = (100 - designWithCM['CM Split']) / 100
                # Need to calculate sales commission from start for these.
                actComm = split * designWithCM['Actual Comm Paid']
                designWithCM['Actual Comm Paid'] = actComm
                sales_comm = commPct * actComm
                designWithCM['Sales Commission'] = sales_comm
        else:
            designOnly = pd.DataFrame(columns=master_cols)
            designWithCM = pd.DataFrame(columns=master_cols)
        # Grab CM + Design Sales entries.
        dualSales = comm_data[[x and y for x, y in zip(CM, Design)]]
        dualSalesComm = commPct * dualSales['Actual Comm Paid']
        dualSales['Sales Commission'] = dualSalesComm
        if dualSales.empty:
            dualSales = pd.DataFrame(columns=master_cols)

        # -----------------------------------------------
        # Grab the QQ entries and combine into one line.
        # -----------------------------------------------
        qqDat = comm_data[comm_data['Design Sales'] == 'QQ']
        qqCondensed = pd.DataFrame(columns=master_cols)
        qqCondensed.loc[0, 'T-End Cust'] = 'MISC POOL'
        qqCondensed.loc[0, 'Sales Commission'] = sum(qqDat['Sales Commission'])
        qqCondensed.loc[0, 'Design Sales'] = 'QQ'
        qqCondensed.loc[0, 'Principal'] = 'VARIOUS (MISC POOL)'
        qqCondensed.loc[0, 'Comm Month'] = currentYrMo
        # Scale down the QQ entries based on the salesperson's share.
        QQperson = sales_info[sales_info['Sales Initials'] == person]
        try:
            QQscale = QQperson['QQ Split'].iloc[0]
            qqCondensed.loc[0, 'Sales Commission'] *= QQscale / 100
        except IndexError:
            # No salesperson QQ split found, so empty it out.
            qqCondensed = pd.DataFrame(columns=master_cols)

        # -----------------------
        # Start creating report.
        # -----------------------
        final_report = pd.DataFrame(columns=master_cols)
        # Append the data.
        final_report = final_report.append([CMOnly[master_cols],
                                            CMWithDesign[master_cols],
                                            designOnly[master_cols],
                                            designWithCM[master_cols],
                                            dualSales[master_cols],
                                            qqCondensed[master_cols]],
                                           ignore_index=True, sort=False)
        # Make sure columns are numeric.
        for col in ['Paid-On Revenue', 'Actual Comm Paid', 'Sales Commission']:
            final_report[col] = pd.to_numeric(final_report[col], errors='coerce').fillna(0)
        # Total up the Paid-On Revenue and Actual/Sales Commission.
        reportTot = pd.DataFrame(columns=['Salesperson', 'Paid-On Revenue',
                                          'Actual Comm Paid',
                                          'Sales Commission'],
                                 index=[0])
        reportTot['Salesperson'] = person
        reportTot['Principal'] = ''
        actComm = sum(final_report['Actual Comm Paid'])
        sales_comm = sum(final_report['Sales Commission'])
        reportTot['Paid-On Revenue'] = sum(final_report['Paid-On Revenue'])
        reportTot['Actual Comm Paid'] = actComm
        reportTot['Sales Commission'] = sales_comm
        # Append to Sales Totals.
        sales_tot = sales_tot.append(reportTot, ignore_index=True, sort=False)
        # Build table of sales by principal.
        princ_tab = data_by_princ_tab(input_data=final_report)
        # Replace zeros with blanks in columns that don't have commission data.
        for col in set(num_cols).difference({'Actual Comm Paid', 'Sales Commission'}):
            try:
                final_report[col].replace(0, '', inplace=True)
            except KeyError:
                pass
        # Write report to file.
        filename = (reports_dir + '\\' + person +
                    ' Commission Report - ' + currentYrMo + '.xlsx')
        writer = pd.ExcelWriter(filename, engine='xlsxwriter',
                                datetime_format='mm/dd/yyyy')
        princ_tab.to_excel(writer, sheet_name='Principals', index=False)
        final_report.to_excel(writer, sheet_name='Raw Data', index=False)
        # Format as table in Excel.
        table_format(princ_tab, 'Principals', writer)
        table_format(final_report, 'Raw Data', writer)
        # Try saving the file, exit with error if file is currently open.
        try:
            writer.save()
        except IOError:
            print('---\n'
                  'A salesperson report file is open in Excel!\n'
                  'Please close the file(s) and try again.\n'
                  '*Program Terminated*')
            return
        # Create the commission pivot table.
        pivots.create_pivot_table(excel_file=filename,
                                  data_sheet_name='Raw Data',
                                  pivot_sheet_name='Comm Table',
                                  row_fields=['T-End Cust', 'Principal'],
                                  col_field='Comm Month',
                                  data_field='Sales Commission')

    # -------------------------------------------------------------------
    # If we're at the end of a quarter, create the quarterly/PDF report.
    # -------------------------------------------------------------------
    if num_prev_mos == 2:
        current_qtr = str(current_year) + 'Q' + str(int(current_month / 3))
        create_quarterly_report(comm_data,  current_qtr, salespeople, sales_info)
        # fullName = sales['Salesperson'].iloc[0]
        # priorComm = sales['Prior Qtr Commission'].iloc[0]
        # priorDue = sales['Prior Qtr Due'].iloc[0]
        # salesDraw = sales['Sales Draw'].iloc[0]
        # pdfReport(fullName, finalReport, priorComm, priorDue, salesDraw)

    # ------------------------------------------------------
    # Create the tabs for the reported Running Commissions.
    # ------------------------------------------------------
    # Generate the table for sales numbers by principal.
    princ_tab = data_by_princ_tab(input_data=running_com)
    for princ_row in princ_tab.index:
        try:
            actual_comm = princ_tab.loc[princ_row, 'Actual Comm Paid']
            revenue = princ_tab.loc[princ_row, 'Paid-On Revenue']
            true_comm_pct = actual_comm / revenue
        except ZeroDivisionError:
            true_comm_pct = ''
        princ_tab.loc[princ_row, 'True Comm %'] = true_comm_pct

    # -----------------------------------
    # Create the overall Revenue Report.
    # -----------------------------------
    # Write the raw data to a file.
    filename = (reports_dir + '\\' + 'Revenue Report - '
                + currentYrMo + RC_addon + '.xlsx')
    writer = pd.ExcelWriter(filename, engine='xlsxwriter',
                            datetime_format='mm/dd/yyyy')
    revenue_data.to_excel(writer, sheet_name='Raw Data', index=False)
    table_format(revenue_data, 'Raw Data', writer)
    # Try saving the report.
    try:
        writer.save()
    except IOError:
        print('---\n'
              'Revenue report file is open in Excel!\n'
              'Please close the file(s) and try again.\n'
              '*Program Terminated*')
        return
    # Add the pivot table for revenue by quarter.
    pivots.create_pivot_table(excel_file=filename,
                              data_sheet_name='Raw Data',
                              pivot_sheet_name='Revenue Table',
                              row_fields=['T-End Cust', 'CM', 'Part Number'],
                              col_field='Quarter Shipped',
                              data_field='Paid-On Revenue',
                              page_field='Principal')

    # -------------------------------------------------------------------------
    # Go through each line of the finished Running Commissions and use them to
    # update the Lookup Master.
    # -------------------------------------------------------------------------
    if run_com:
        # Don't copy over INDIVIDUAL, MISC, or ALLOWANCE.
        noCopy = ['INDIVIDUAL', 'UNKNOWN', 'ALLOWANCE']
        paredID = [i for i in running_com.index
                   if not any(j in running_com.loc[i, 'T-End Cust'].upper()
                              for j in noCopy)]
        for row in paredID:
            # First match reported customer.
            repCust = str(running_com.loc[row, 'Reported Customer']).lower()
            POSCust = look_mast['Reported Customer'].map(
                lambda x: str(x).lower())
            custMatches = look_mast[repCust == POSCust]
            # Now match part number.
            partNum = str(running_com.loc[row, 'Part Number']).lower()
            PPN = look_mast['Part Number'].map(lambda x: str(x).lower())
            fullMatches = custMatches[PPN == partNum]
            # Figure out if this entry is a duplicate of any existing entry.
            duplicate = False
            for matchID in fullMatches.index:
                matchCols = ['CM Sales', 'Design Sales', 'CM', 'T-Name',
                             'T-End Cust']
                duplicate = all(
                    fullMatches.loc[matchID, i] == running_com.loc[row, i]
                    for i in matchCols)
                if duplicate:
                    break
            # If it's not an exact duplicate, add it to the Lookup Master.
            if not duplicate:
                lookupCols = ['CM Sales', 'Design Sales', 'CM Split', 'CM',
                              'T-Name', 'T-End Cust', 'Reported Customer',
                              'Principal', 'Part Number', 'City']
                newLookup = running_com.loc[row, lookupCols]
                newLookup['Date Added'] = datetime.datetime.now().date()
                newLookup['Last Used'] = datetime.datetime.now().date()
                look_mast = look_mast.append(newLookup, ignore_index=True)

        # --------------------------------------------------------------
        # Append the new Running Commissions to the Commissions Master.
        # --------------------------------------------------------------
        com_mast = com_mast.append(running_com, ignore_index=True, sort=False)
        master_files = master_files.append(files_processed, ignore_index=True,
                                           sort=False)
        # Make sure all the dates are formatted correctly.
        com_mast['Invoice Date'] = com_mast['Invoice Date'].map(
            lambda x: form_date(x))
        master_files['Date Added'] = master_files['Date Added'].map(
            lambda x: form_date(x))
        master_files['Paid Date'] = master_files['Paid Date'].map(
            lambda x: form_date(x))
        # Convert commission dollars to numeric.
        master_files['Total Commissions'] = pd.to_numeric(
            master_files['Total Commissions'], errors='coerce').fillna(0)
        for col in num_cols:
            try:
                if col not in ['Actual Comm Paid', 'Sales Commission']:
                    fill = ''
                else:
                    fill = 0
                com_mast[col] = pd.to_numeric(com_mast[col],
                                              errors='coerce').fillna(fill)
            except KeyError:
                pass
        # Convert individual numbers to numeric in rest of columns.
        mixed_cols = [col for col in list(com_mast) if col not in num_cols]
        # Invoice/part numbers sometimes has leading zeros we'd like to keep.
        mixed_cols.remove('Invoice Number')
        mixed_cols.remove('Part Number')
        # The INF gets read in as infinity, so skip the principal column.
        mixed_cols.remove('Principal')
        for col in mixed_cols:
            com_mast[col] = com_mast[col].map(
                lambda x: pd.to_numeric(x, errors='ignore'))

    # %%
    # Save the files.
    fname1 = data_dir + '\\Commissions Master.xlsx'
    fname2 = (reports_dir + '\\Running Commissions ' + currentYrMo
              + ' Reported' + RC_addon + '.xlsx')
    fname3 = look_dir + '\\Lookup Master - Current.xlsx'

    if save_error(fname1, fname2, fname3):
        print('---\n'
              'One or more of these files are currently open in Excel:\n'
              'Running Commissions, Entries Need Fixing, Lookup Master.\n'
              'Please close these files and try again.\n'
              '*Program Terminated*')
        return

    if run_com:
        # Write the Commissions Master file.
        writer = pd.ExcelWriter(fname1, engine='xlsxwriter',
                                datetime_format='mm/dd/yyyy')
        com_mast.to_excel(writer, sheet_name='Master Data', index=False)
        master_files.to_excel(writer, sheet_name='Files Processed', index=False)
        # Format everything in Excel.
        table_format(com_mast, 'Master Data', writer)
        table_format(master_files, 'Files Processed', writer)

        # Write the Lookup Master.
        writer2 = pd.ExcelWriter(fname3, engine='xlsxwriter',
                                 datetime_format='mm/dd/yyyy')
        look_mast.to_excel(writer2, sheet_name='Lookup', index=False)
        # Format everything in Excel.
        table_format(look_mast, 'Lookup', writer2)

    # Write the Running Commissions report.
    writer1 = pd.ExcelWriter(fname2, engine='xlsxwriter',
                             datetime_format='mm/dd/yyyy')
    running_com.to_excel(writer1, sheet_name='Data', index=False)
    if run_com:
        files_processed.to_excel(writer1, sheet_name='Files Processed',
                                 index=False)
        table_format(files_processed, 'Files Processed', writer1)
    sales_tot.to_excel(writer1, sheet_name='Salesperson Totals',
                       index=False)
    princ_tab.to_excel(writer1, sheet_name='Principal Totals',
                       index=False)
    # Format as table in Excel.
    table_format(running_com, 'Data', writer1)
    table_format(sales_tot, 'Salesperson Totals', writer1)
    table_format(princ_tab, 'Principal Totals', writer1)

    # Save the files.
    if run_com:
        writer.save()
        writer2.save()
    writer1.save()
    print('---\nSales reports finished successfully!\n')
    if run_com:
        print('---\nCommissions Master updated.\n'
              'Lookup Master updated.\n+++')
