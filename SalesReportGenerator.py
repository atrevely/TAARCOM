import pandas as pd
import numpy as np
import time
import datetime
import os
from dateutil.parser import parse
from RCExcelTools import tab_save_prep, save_error, PivotTables
from FileIO import load_salespeople_info, load_com_master, load_run_com, load_acct_list, load_lookup_master
# from PDFReportGenerator import pdfReport

# Set the numerical columns.
num_cols = ['Quantity', 'Ext. Cost', 'Invoiced Dollars', 'Paid-On Revenue', 'Actual Comm Paid',
            'Unit Cost', 'Unit Price', 'CM Split', 'Year', 'Sales Commission',
            'Split Percentage', 'Commission Rate', 'Gross Rev Reduction', 'Shared Rev Tier Rate']

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
    """
    Returns all the data for a particular salesperson, with sales commission scaled down
    by split percentage.
    """
    # Determine the salesperson's commission percentage.
    sales = sales_info[sales_info['Sales Initials'] == salesperson]
    comm_pct = sales['Sales Percentage'].iloc[0] / 100
    # Grab the data that has either CM or design sales for this person.
    CM = input_data['CM Sales'] == salesperson
    design = input_data['Design Sales'] == salesperson
    sales_data = input_data[np.logical_or(CM, design)]
    # Get the lines that are shared with other salespeople.
    shared_sales = np.logical_and(sales_data['CM Sales'] != '', sales_data['Design Sales'] != '')
    # Scale the commission data by split percentage.
    for row in shared_sales[shared_sales].index:
        if CM.loc[row] and not design.loc[row]:
            split = sales_data.loc[row, 'CM Split'] / 100
        elif design.loc[row] and not CM.loc[row]:
            split = 1 - sales_data.loc[row, 'CM Split'] / 100
        else:
            split = 1
        sales_data.loc[row, 'Actual Comm Paid'] *= split
        sales_data.loc[row, 'Sales Commission'] = comm_pct * sales_data.loc[row, 'Actual Comm Paid']
    sales_data.reset_index(drop=True, inplace=True)
    return sales_data


def data_by_princ_tab(input_data):
    """Builds a DataFrame of the provided data broken down by principal."""
    princ_tab = pd.DataFrame(columns=['Principal', 'Paid-On Revenue', 'Actual Comm Paid', 'Sales Commission'])
    # Tally up totals for each principal.
    for row, principal in enumerate(input_data['Principal'].unique()):
        princ_sales = input_data[input_data['Principal'] == principal]
        princ_tab.loc[row, 'Principal'] = principal
        for col in princ_tab.columns[1:]:
            princ_tab.loc[row, col] = sum(princ_sales[col])
    # Sort principals in descending order alphabetically.
    princ_tab.sort_values(by=['Principal'], inplace=True)
    princ_tab.reset_index(drop=True, inplace=True)
    # Fill in overall totals on the last row.
    princ_tab.loc[princ_tab.__len__(), 'Principal'] = 'Grand Total'
    for col in princ_tab.columns[1:]:
        princ_tab.loc[princ_tab.__len__(), col] = np.sum(princ_tab[col])
    return princ_tab


def create_quarterly_report(comm_data, comm_qtr, salespeople, sales_info):
    """Builds the report that runs at the end of each quarter."""
    print('---\nCreating end-of-quarter report.')
    # ---------------------------------------------------------
    # Build the tab with commissions broken down by principal.
    # ---------------------------------------------------------
    princ_tab = data_by_princ_tab(input_data=comm_data)
    for princ_row in princ_tab.index:
        actual_comm = princ_tab.loc[princ_row, 'Actual Comm Paid']
        revenue = princ_tab.loc[princ_row, 'Paid-On Revenue']
        if actual_comm and revenue:
            true_comm_pct = actual_comm / revenue
        else:
            true_comm_pct = ''
        princ_tab.loc[princ_row, 'True Comm %'] = true_comm_pct
    # ---------------------------------------------------------------------------
    # Build the tab with commissions broken down by salesperson, then principal.
    # ---------------------------------------------------------------------------
    sales_tab = pd.DataFrame(columns=['Salesperson', 'Principal', 'Actual Comm Paid', 'Sales Commission'])
    row = 0
    for person in salespeople:
        sales_data = get_sales_comm_data(salesperson=person, input_data=comm_data, sales_info=sales_info)
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
    # -----------------
    # Save the report.
    # -----------------
    filename = os.path.join(reports_dir, 'Quarterly Commission Report ' + comm_qtr + '.xlsx')
    writer = pd.ExcelWriter(filename, engine='xlsxwriter', datetime_format='mm/dd/yyyy')
    tab_save_prep(writer=writer, data=comm_data, sheet_name='Comm Data')
    tab_save_prep(writer=writer, data=princ_tab, sheet_name='Principals')
    tab_save_prep(writer=writer, data=sales_tab, sheet_name='Salespeople')
    if not save_error(filename):
        writer.save()
    else:
        print('Error saving quarter commission report.')


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
    sales_info = load_salespeople_info()
    acct_list = load_acct_list()
    com_mast, master_files = load_com_master()
    if any([acct_list.empty, sales_info.empty, com_mast.empty, master_files.empty]):
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
        look_mast = load_lookup_master()
        running_com, files_processed = load_run_com(file_path=run_com)
        if any([look_mast.empty, running_com.empty, files_processed.empty]):
            print('*Program Terminated*')
            return
        # Fill in the Sales Report Date in Running Commissions.
        running_com.loc[running_com['Sales Report Date'] == '', 'Sales Report Date'] = time.strftime('%m/%d/%Y')
        # -------------------------------------------------------------------
        # Check to make sure new files aren't already in Commissions Master.
        # -------------------------------------------------------------------
        # Check if we've duplicated any files.
        filenames = master_files['Filename']
        duplicates = list(set(filenames).intersection(files_processed['Filename']))
        # Don't let duplicate files get processed.
        if duplicates:
            # Let us know we found duplicates and removed them.
            print('---\nThe following files are already in Commissions Master:\n%s'
                  % ', '.join(map(str, duplicates)) + '\nPlease check '
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
        current_yr_mo = str(current_year) + '-' + str(current_month)
        running_com['Comm Month'] = current_yr_mo
        # No add-on since this is the first RC run.
        RC_addon = ''
    else:
        # -----------------------------------------------------------------
        # Use the most recent commission month as the RC from here on out.
        # -----------------------------------------------------------------
        current_month = last_month.month
        current_year = last_month.year
        current_yr_mo = str(current_year) + '-' + str(current_month)
        running_com = com_mast[com_mast['Comm Month'] == current_yr_mo]
        # Indicate that this is a rerun.
        RC_addon = ' (Rerun)'
        print('No new RC supplied. Reporting on latest quarter in the Commissions Master.')

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
        if np.sum(acct_list['ProperName'] == cust) == 1:
            try:
                sales = acct_list[acct_list['ProperName'] == cust]['SLS'].iloc[0]
                cust_ID = revenue_data[revenue_data['T-End Cust'] == cust].index
                revenue_data.loc[cust_ID, 'CDS'] = sales
            except KeyError:
                print('Error reading column names in Account List!\n'
                      'Please make sure the columns ProperName and SLS are in '
                      'the Account List.\n*Program Terminated*')
                return
    # Fill in the CDS (current design sales) for missing entries as simply the
    # Design Sales for that line.
    for row in revenue_data[pd.isna(revenue_data['CDS'])].index:
        revenue_data.loc[row, 'CDS'] = revenue_data.loc[row, 'Design Sales']
        # If no design sales, use CM sales.
        if not revenue_data.loc[row, 'CDS']:
            revenue_data.loc[row, 'CDS'] = revenue_data.loc[row, 'CM Sales']
    # Also grab the section of the data that aren't 80/20 splits.
    split_data = revenue_data[revenue_data['CM Split'] != 20]

    # ---------------------------------------------------------
    # Combine and tag commission data for the current quarter.
    # ---------------------------------------------------------
    # Figure out what slice of commissions data is in the current quarter.
    com_mast_tracked = com_mast[com_mast['Comm Month'] != '']
    try:
        com_mast_tracked['Comm Month'].map(lambda x: parse(str(x)))
    except (TypeError, ValueError):
        print('Error reading month in Comm Month column!\n'
              'Please make sure all months are in YYYY-MM format.\n*Program Terminated*')
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
    print('Found the following sales initials in the Salespeople Info file: ' + ', '.join(salespeople))
    # Create the dataframe with the commission information by salesperson.
    sales_tot = pd.DataFrame(columns=['Salesperson', 'Principal', 'Paid-On Revenue', 'Actual Comm Paid',
                                      'Sales Commission'])

    # Go through each salesperson and prepare their reports.
    print('Running reports...')
    for person in salespeople:
        # ------------------------------------------------------------
        # Create the revenue reports for each salesperson, using only
        # design data.
        # ------------------------------------------------------------
        # Grab the raw data for this salesperson's design sales.
        design_data = revenue_data[revenue_data['CDS'] == person]
        # Also grab any nonstandard splits.
        cm_data = split_data[split_data['CM Sales'] == person]
        cm_data = cm_data[cm_data['CDS'] != person]
        design_data = design_data.append(cm_data, ignore_index=True, sort=False)
        # Get rid of empty Quarter Shipped lines.
        design_data = design_data[design_data['Quarter Shipped'] != '']
        design_data.reset_index(drop=True, inplace=True)
        # Write the raw data to a file.
        filename = (reports_dir + '\\' + person + ' Revenue Report - ' + current_yr_mo + '.xlsx')
        writer = pd.ExcelWriter(filename, engine='xlsxwriter', datetime_format='mm/dd/yyyy')
        tab_save_prep(writer=writer, data=design_data, sheet_name='Raw Data')
        # Try saving the report.
        try:
            writer.save()
        except IOError:
            print('---\nA salesperson report file is open in Excel!\n'
                  'Please close the file(s) and try again.\n*Program Terminated*')
            return
        # Create the revenue pivot table.
        pivots.create_pivot_table(excel_file=filename, data_sheet_name='Raw Data',
                                  pivot_sheet_name='Revenue Table',
                                  row_fields=['T-End Cust', 'Part Number', 'CM'],
                                  col_field='Quarter Shipped', data_field='Paid-On Revenue',
                                  page_field='Principal')

        # -----------------------------------------------
        # Grab the QQ entries and combine into one line.
        # -----------------------------------------------
        qq_data = comm_data[comm_data['Design Sales'] == 'QQ']
        qq_condensed = pd.DataFrame(columns=master_cols)
        qq_condensed.loc[0, 'T-End Cust'] = 'MISC POOL'
        qq_condensed.loc[0, 'Sales Commission'] = sum(qq_data['Sales Commission'])
        qq_condensed.loc[0, 'Design Sales'] = 'QQ'
        qq_condensed.loc[0, 'Principal'] = 'VARIOUS (MISC POOL)'
        qq_condensed.loc[0, 'Comm Month'] = current_yr_mo
        # Scale down the QQ entries based on the salesperson's share.
        qq_person = sales_info[sales_info['Sales Initials'] == person]
        try:
            qq_scale = qq_person['QQ Split'].iloc[0]
            qq_condensed.loc[0, 'Sales Commission'] *= qq_scale / 100
        except IndexError:
            # No salesperson QQ split found, so empty it out.
            qq_condensed = pd.DataFrame(columns=master_cols)

        # --------------------------------------------------------------------
        # Create the commission reports for each salesperson, using all data.
        # --------------------------------------------------------------------
        final_report = get_sales_comm_data(salesperson=person, input_data=comm_data,
                                           sales_info=sales_info)
        # Append the data.
        final_report = final_report.append(qq_condensed, ignore_index=True, sort=False)
        # Total up the Paid-On Revenue and Actual/Sales Commission.
        report_total = pd.DataFrame(columns=['Salesperson', 'Paid-On Revenue', 'Actual Comm Paid',
                                             'Sales Commission'], index=[0])
        report_total['Salesperson'] = person
        report_total['Principal'] = ''
        actual_comm = sum(final_report['Actual Comm Paid'])
        sales_comm = sum(final_report['Sales Commission'])
        report_total['Paid-On Revenue'] = sum(final_report['Paid-On Revenue'])
        report_total['Actual Comm Paid'] = actual_comm
        report_total['Sales Commission'] = sales_comm
        # Build table of sales by principal.
        princ_tab = data_by_princ_tab(input_data=final_report)
        # Append to Sales Totals.
        person_total = princ_tab[princ_tab['Principal'] == 'Grand Total']
        person_total['Salesperson'] = person
        person_total['Principal'] = ''
        sales_tot = sales_tot.append(person_total, ignore_index=True, sort=False)
        sales_tot = sales_tot.append(princ_tab[princ_tab['Principal'] != 'Grand Total'],
                                     ignore_index=True, sort=False)
        # Write report to file.
        filename = (reports_dir + '\\' + person + ' Commission Report - ' + current_yr_mo + '.xlsx')
        writer = pd.ExcelWriter(filename, engine='xlsxwriter', datetime_format='mm/dd/yyyy')
        # Prepare the data in Excel.
        tab_save_prep(writer=writer, data=princ_tab, sheet_name='Principals')
        tab_save_prep(writer=writer, data=final_report, sheet_name='Raw Data')
        # Try saving the file, exit with error if file is currently open.
        try:
            writer.save()
        except IOError:
            print('---\nA salesperson report file is open in Excel!\n'
                  'Please close the file(s) and try again.\n*Program Terminated*')
            return
        # Create the commission pivot table.
        pivots.create_pivot_table(excel_file=filename,  data_sheet_name='Raw Data',
                                  pivot_sheet_name='Comm Table', row_fields=['T-End Cust', 'Principal'],
                                  col_field='Comm Month', data_field='Sales Commission')

    # -------------------------------------------------------------------
    # If we're at the end of a quarter, create the quarterly/PDF report.
    # -------------------------------------------------------------------
    if num_prev_mos == 2:
        current_qtr = str(current_year) + 'Q' + str(int(current_month / 3))
        create_quarterly_report(comm_data=comm_data, comm_qtr=current_qtr,
                                salespeople=salespeople, sales_info=sales_info)
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
    filename = (reports_dir + '\\' + 'Revenue Report - ' + current_yr_mo + RC_addon + '.xlsx')
    writer = pd.ExcelWriter(filename, engine='xlsxwriter', datetime_format='mm/dd/yyyy')
    tab_save_prep(writer=writer, data=revenue_data, sheet_name='Raw Data')
    # Try saving the report.
    try:
        writer.save()
    except IOError:
        print('---\nRevenue report file is open in Excel!\n'
              'Please close the file(s) and try again.\n'
              '*Program Terminated*')
        return
    # Add the pivot table for revenue by quarter.
    pivots.create_pivot_table(excel_file=filename,  data_sheet_name='Raw Data',
                              pivot_sheet_name='Revenue Table',
                              row_fields=['T-End Cust', 'CM', 'Part Number'],
                              col_field='Quarter Shipped', data_field='Paid-On Revenue',
                              page_field='Principal')

    # -------------------------------------------------------------------------
    # Go through each line of the finished Running Commissions and use them to
    # update the Lookup Master.
    # -------------------------------------------------------------------------
    if run_com:
        # Don't copy over INDIVIDUAL, MISC, or ALLOWANCE.
        no_copy = ['INDIVIDUAL', 'UNKNOWN', 'ALLOWANCE']
        pared_ID = [i for i in running_com.index
                    if not any(j in running_com.loc[i, 'T-End Cust'].upper() for j in no_copy)]
        for row in pared_ID:
            # First match reported customer.
            reported_cust = str(running_com.loc[row, 'Reported Customer']).lower()
            POS_cust = look_mast['Reported Customer'].map(lambda x: str(x).lower())
            cust_matches = look_mast[reported_cust == POS_cust]
            # Now match part number.
            part_num = str(running_com.loc[row, 'Part Number']).lower()
            PPN = look_mast['Part Number'].map(lambda x: str(x).lower())
            full_matches = cust_matches[PPN == part_num]
            # Figure out if this entry is a duplicate of any existing entry.
            duplicate = False
            for matchID in full_matches.index:
                match_cols = ['CM Sales', 'Design Sales', 'CM', 'T-Name', 'T-End Cust']
                duplicate = all(full_matches.loc[matchID, i] == running_com.loc[row, i] for i in match_cols)
                if duplicate:
                    break
            # If it's not an exact duplicate, add it to the Lookup Master.
            if not duplicate:
                lookup_cols = ['CM Sales', 'Design Sales', 'CM Split', 'CM', 'T-Name', 'T-End Cust',
                               'Reported Customer', 'Principal', 'Part Number', 'City']
                new_lookup = running_com.loc[row, lookup_cols]
                new_lookup['Date Added'] = datetime.datetime.now().date()
                new_lookup['Last Used'] = datetime.datetime.now().date()
                look_mast = look_mast.append(new_lookup, ignore_index=True)

        # --------------------------------------------------------------
        # Append the new Running Commissions to the Commissions Master.
        # --------------------------------------------------------------
        com_mast = com_mast.append(running_com, ignore_index=True, sort=False)
        master_files = master_files.append(files_processed, ignore_index=True, sort=False)
        # Convert commission dollars to numeric.
        master_files['Total Commissions'] = pd.to_numeric(master_files['Total Commissions'],
                                                          errors='coerce').fillna(0)

    # ----------------
    # Save the files.
    # ----------------
    filename_1 = data_dir + '\\Commissions Master.xlsx'
    filename_2 = look_dir + '\\Lookup Master - Current.xlsx'
    filename_3 = reports_dir + '\\Running Commissions ' + current_yr_mo + ' Reported' + RC_addon + '.xlsx'

    if save_error(filename_1, filename_2, filename_3):
        print('---\nOne or more of these files are currently open in Excel:\n'
              'Running Commissions, Commissions Master, Lookup Master.\n'
              'Please close these files and try again.\n*Program Terminated*')
        return

    if run_com:
        # Write the Commissions Master file if new RC data was added to it.
        writer1 = pd.ExcelWriter(filename_1, engine='xlsxwriter', datetime_format='mm/dd/yyyy')
        tab_save_prep(writer=writer1, data=com_mast, sheet_name='Master')
        tab_save_prep(writer=writer1, data=master_files, sheet_name='Files Processed')

        # Write the Lookup Master.
        writer2 = pd.ExcelWriter(filename_2, engine='xlsxwriter', datetime_format='mm/dd/yyyy')
        tab_save_prep(writer=writer2, data=look_mast, sheet_name='Lookup')

    # Write the Running Commissions report.
    writer3 = pd.ExcelWriter(filename_3, engine='xlsxwriter', datetime_format='mm/dd/yyyy')
    tab_save_prep(writer=writer3, data=running_com, sheet_name='Data')
    if run_com:
        # Only write the Files Processed tab if it's a new RC.
        tab_save_prep(writer=writer3, data=files_processed, sheet_name='Files Processed')
    tab_save_prep(writer=writer3, data=sales_tot, sheet_name='Salesperson Totals')
    tab_save_prep(writer=writer3, data=princ_tab, sheet_name='Principal Totals')

    # Save the files.
    if run_com:
        writer1.save()
        writer2.save()
    writer3.save()
    print('---\nSales reports finished successfully!')
    if run_com:
        print('---\nCommissions Master updated.\nLookup Master updated.')
    print('+Program Complete+')
