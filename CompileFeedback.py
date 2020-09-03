import pandas as pd
import os
import time
from RCExcelTools import save_error
from FileLoader import load_salespeople_info, load_root_customer_mappings, load_acct_list, load_digikey_master
from xlrd import XLRDError

# Set the directory for the data input/output.
if os.path.exists('Z:\\'):
    data_dir = 'W:\\'
    look_dir = 'Z:\\Commissions Lookup'
else:
    data_dir = os.getcwd()
    look_dir = os.getcwd()


def table_format(sheet_data, sheet_name, workbook):
    """Formats the Excel output as a table with correct column formatting."""
    # Nothing to format, so return.
    if sheet_data.shape[0] == 0:
        return
    sheet = workbook.sheets[sheet_name]
    sheet.freeze_panes(1, 0)
    # Set the autofilter for the sheet.
    sheet.autofilter(0, 0, sheet_data.shape[0], sheet_data.shape[1] - 1)
    # Set document formatting.
    doc_format = workbook.book.add_format({'font': 'Calibri', 'font_size': 11})
    accounting_format = workbook.book.add_format({'font': 'Calibri', 'font_size': 11, 'num_format': 44})
    comma_format = workbook.book.add_format({'font': 'Calibri', 'font_size': 11, 'num_format': 3})
    # Format and fit each column.
    i = 0
    # Columns which get shrunk down in reports.
    hidden_cols = ['Technology', 'Excel Part Link', 'Report Part Nbr Link', 'MFG Part Description', 'Focus',
                   'Part Class Name', 'Vendor ID', 'Invoice Detail Nbr', 'Assigned Account Rep',
                   'Recipient', 'DKLI Report Date', 'Invoice Date Group', 'Comments', 'Sales Channel']
    core_cols = ['Must Contact', 'End Product', 'How Contacted', 'Information for Digikey']
    for col in sheet_data.columns:
        accounting_cols = ['Unit Price', 'Invoiced Dollars']
        if col in accounting_cols:
            formatting = accounting_format
        elif col == 'Quantity':
            formatting = comma_format
        else:
            formatting = doc_format
        max_width = max(len(str(val)) for val in sheet_data[col].values)
        # Set maximum column width at 50.
        max_width = min(max_width, 50)
        if col in hidden_cols:
            max_width = 0
        elif col in core_cols:
            max_width = 25
        sheet.set_column(i, i, max_width+0.8, formatting)
        i += 1
    # Set the autofilter for the sheet.
    sheet.autofilter(0, 0, sheet_data.shape[0], sheet_data.shape[1]-1)


# The main function.
def main(filepaths):
    """Combine files into one finalized monthly Digikey file, and append it
    to the Digikey Insights Master file. Also updates the rootCustomerMappings file.

    Arguments:
    filepaths -- The filepaths to the files with new comments.
    """
    # --------------------------------------------------------
    # Load in the supporting files, exit if any aren't found.
    # --------------------------------------------------------
    sales_info = load_salespeople_info(file_dir=look_dir)
    customer_mappings = load_root_customer_mappings(file_dir=look_dir)
    acct_list = load_acct_list(file_dir=look_dir)
    digikey_master, files_processed = load_digikey_master(file_dir=data_dir)

    if any([sales_info.empty, customer_mappings.empty, acct_list.empty, digikey_master.empty]):
        print('*Program Terminated*')
        return

    # ------------------------
    # Load the Insight files.
    # ------------------------
    # Strip the root off of the filepaths and leave just the filenames.
    filenames = [os.path.basename(i) for i in filepaths]
    try:
        input_data = [pd.read_excel(i) for i in filepaths]
    except XLRDError:
        print('---\nError reading in files!\n*Program Terminated*')
        return

    # -----------------------------------------------
    # Combine the report data from each salesperson.
    # -----------------------------------------------
    # Make sure each filename has a salesperson initials.
    salespeople = sales_info['Sales Initials'].values
    initials_list = []
    for filename in filenames:
        initials = filename[0:2].upper()
        if initials not in salespeople:
            print('Salesperson initials ' + initials + ' not recognized!\n'
                  'Make sure the first two letters of each filename are salesperson initials.\n'
                  '*Program Terminated*')
            return
        elif initials in initials_list:
            print('Salesperson initials ' + initials + ' duplicated!\n'
                  'Make sure each salesperson has at most one file.\n'
                  '*Program Terminated*')
            return
        initials_list.append(initials)
    # Create the master dataframe to append to.
    final_data = pd.DataFrame(columns=digikey_master.columns)
    # Copy over the comments.
    file_num = 0
    for sheet in input_data:
        print('---\nCopying comments from file: ' + filenames[file_num])
        # Grab only the salesperson's data.
        sales = filenames[file_num][0:2]
        sheet_data = sheet[sheet['Sales'] == sales]
        # Append data to the output dataframe.
        final_data = final_data.append(sheet_data, ignore_index=True, sort=False)
        # Next file.
        file_num += 1
    # Drop any unnamed columns that got processed.
    try:
        final_data = final_data.loc[:, ~final_data.columns.str.contains('^Unnamed')]
        final_data = final_data.loc[:, digikey_master.columns]
    except AttributeError:
        pass

    # --------------------------------------
    # Update the rootCustomerMappings file.
    # --------------------------------------
    for row in final_data.index:
        # Get root customer and salesperson.
        cust = str(final_data.loc[row, 'Root Customer..']).lower().rstrip()
        salesperson = final_data.loc[row, 'Sales']
        try:
            indiv = final_data.loc[row, 'Root Customer Class'].lower().rstrip() == 'individual'
        except AttributeError:
            indiv = False
        if cust and salesperson and not indiv:
            # Find match in rootCustomerMappings.
            cust_match = customer_mappings['Root Customer'].astype(str).str.lower() == cust
            if sum(cust_match) == 1:
                match_id = customer_mappings[cust_match].index
                # Input (possibly new) salesperson.
                customer_mappings.loc[match_id, 'Salesperson'] = salesperson
            elif not cust_match.any():
                # New customer (no match), so append to mappings.
                new_cust = pd.DataFrame({'Root Customer': [cust], 'Salesperson': [salesperson]})
                customer_mappings = customer_mappings.append(new_cust, ignore_index=True, sort=False)
            else:
                print('There appears to be a duplicate customer in rootCustomerMappings:\n'
                      + str(cust) + '\nPlease trim to one entry and try again.'
                      + '\n*Program Terminated*')
                return

    # ----------------------------------------------------------------------------------------
    # Append the new data to the Digikey Insight Master, then update the Current Salesperson.
    # ----------------------------------------------------------------------------------------
    master_cols = list(digikey_master)
    master_cols.remove('Current Sales')
    digikey_master = digikey_master.append(final_data[master_cols], ignore_index=True, sort=False)
    digikey_master.fillna('', inplace=True)
    final_data.fillna('', inplace=True)
    # Go through each root customer and update current salesperson.
    for cust in digikey_master['Root Customer..'].unique():
        current_sales = ''
        # First check the Account List.
        acct_match = acct_list[acct_list['ProperName'] == cust]
        if not acct_match.empty:
            current_sales = acct_match['SLS'].iloc[0]
        # Next try rootCustomerMappings.
        map_match = customer_mappings[customer_mappings['Root Customer'] == cust]
        if acct_match.empty and not map_match.empty:
            try:
                current_sales = map_match['Current Sales'].iloc[0]
            except KeyError:
                pass
        # Update current salesperson, if a new one is found.
        if current_sales:
            match_id = digikey_master[digikey_master['Root Customer..'] == cust].index
            digikey_master.loc[match_id, 'Current Sales'] = current_sales

    # ---------------------------------------------------------------------
    # Try saving the files, exit with error if any file is currently open.
    # ---------------------------------------------------------------------
    fname1 = data_dir + '\\Digikey Insight Final ' + time.strftime('%Y-%m-%d') + '.xlsx'
    # Append the new file to files processed.
    new_file = pd.DataFrame(columns=files_processed.columns)
    new_file.loc[0, 'Filename'] = fname1
    files_processed = files_processed.append(new_file, ignore_index=True, sort=False)
    fname2 = data_dir + '\\Digikey Insight Master.xlsx'
    fname3 = look_dir + '\\rootCustomerMappings.xlsx'
    if save_error(fname1, fname2, fname3):
        print('---\nInsight Master, Insight Final, or rootCustomerMapings is currently open in Excel!\n'
              'Please close the file(s) and try again.\n*Program Terminated*')
        return
    # Write the finished Insight file.
    writer1 = pd.ExcelWriter(fname1, engine='xlsxwriter', datetime_format='mm/dd/yyyy')
    final_data.to_excel(writer1, sheet_name='Master', index=False)
    table_format(final_data, 'Master', writer1)
    # Write the Insight Master file.
    writer2 = pd.ExcelWriter(fname2, engine='xlsxwriter', datetime_format='mm/dd/yyyy')
    digikey_master.to_excel(writer2, sheet_name='Master', index=False)
    files_processed.to_excel(writer2, sheet_name='Files Processed', index=False)
    table_format(digikey_master, 'Master', writer2)
    table_format(files_processed, 'Files Processed', writer2)
    # Write the new rootCustomerMappings file.
    writer3 = pd.ExcelWriter(fname3, engine='xlsxwriter', datetime_format='mm/dd/yyyy')
    customer_mappings.to_excel(writer3, sheet_name='Sales Lookup', index=False)
    table_format(customer_mappings, 'Sales Lookup', writer3)
    # Save the files.
    writer1.save()
    writer2.save()
    writer3.save()
    print('---\nUpdates completed successfully!\n---\nDigikey Master updated.\n'
          'rootCustomerMappings updated.\n+Program Complete+')
