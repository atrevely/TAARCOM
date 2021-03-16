import pandas as pd
import numpy as np
from dateutil.parser import parse
from xlrd import XLRDError
import time
import calendar
import math
import os.path
import re
import datetime
from uuid import uuid4
from FileLoader import load_lookup_master, load_run_com, load_entries_need_fixing
from FileSaver import prepare_save_file, save_files
from RCExcelTools import save_error, form_date

# Set the directory for the data input/output.
if os.path.exists('Z:\\'):
    out_dir = 'Z:\\MK Working Commissions'
    look_dir = 'Z:\\Commissions Lookup'
    match_dir = 'Z:\\Matched Raw Data Files'
else:
    out_dir = os.getcwd()
    look_dir = os.getcwd()
    match_dir = os.getcwd()


def preprocess_by_principal(princ, sheet, sheet_name):
    """Do special pre-processing tailored to the principal input. Primarily,
    this involves renaming columns that would get looked up incorrectly
    in the Field Mappings.

    This function modifies a dataframe inplace.
    """
    # Initialize the rename_dict in case it doesn't get set later.
    rename_dict = {}
    # ------------------------------
    # Osram special pre-processing.
    # ------------------------------
    if princ == 'OSR':
        rename_dict = {'Item': 'Unmapped', 'Material Number': 'Unmapped 2',
                       'Customer Name': 'Unmapped 3', 'Sales Date': 'Unmapped 4'}
        sheet.rename(columns=rename_dict, inplace=True)
        # Combine Rep 1 % and Rep 2 %.
        if 'Rep 1 %' in list(sheet) and 'Rep 2 %' in list(sheet):
            print('Copying Rep 2 % into empty Rep 1 % lines.\n---')
            for row in sheet.index:
                if sheet.loc[row, 'Rep 2 %'] and not sheet.loc[row, 'Rep 1 %']:
                    sheet.loc[row, 'Rep 1 %'] = sheet.loc[row, 'Rep 2 %']
    # -----------------------------
    # ISSI special pre-processing.
    # -----------------------------
    elif princ == 'ISS':
        rename_dict = {'Commission Due': 'Unmapped', 'Name': 'OEM/POS'}
        sheet.rename(columns=rename_dict, inplace=True)
    # ----------------------------
    # ATS special pre-processing.
    # ----------------------------
    elif princ == 'ATS':
        rename_dict = {'Resale': 'Extended Resale', 'Cost': 'Extended Cost'}
        sheet.rename(columns=rename_dict, inplace=True)
    # ----------------------------
    # QRF special pre-processing.
    # ----------------------------
    elif princ == 'QRF':
        if sheet_name in ['OEM', 'OFF']:
            rename_dict = {'End Customer': 'Unmapped 2', 'Item': 'Unmapped 3'}
            sheet.rename(columns=rename_dict, inplace=True)
        elif sheet_name == 'POS':
            rename_dict = {'Company': 'Distributor', 'BillDocNo': 'Unmapped',
                           'End Customer': 'Unmapped 2', 'Item': 'Unmapped 3'}
            sheet.rename(columns=rename_dict, inplace=True)
    # ----------------------------
    # XMO special pre-processing.
    # ----------------------------
    elif princ == 'XMO':
        rename_dict = {'Amount': 'Commission', 'Commission Due': 'Unmapped'}
        sheet.rename(columns=rename_dict, inplace=True)
    # Return the rename_dict for future use in the matched raw file.
    if rename_dict:
        print('The following columns were renamed automatically on this sheet:')
        [print(i + ' --> ' + j) for i, j in zip(rename_dict.keys(), rename_dict.values())]
    return rename_dict


def process_by_principal(princ, sheet, sheet_name, disty_map):
    """Do special processing tailored to the principal input. This involves
    things like filling in commissions source as cost/resale, setting some
    commission rates that aren't specified in the data, etc.

    This function modifies a dataframe inplace.
    """
    # Make sure applicable entries exist and are numeric.
    invoice_dollars = True
    ext_cost = True
    try:
        sheet['Invoiced Dollars'] = pd.to_numeric(sheet['Invoiced Dollars'], errors='coerce').fillna(0)
    except KeyError:
        invoice_dollars = False
    try:
        sheet['Ext. Cost'] = pd.to_numeric(sheet['Ext. Cost'], errors='coerce').fillna(0)
    except KeyError:
        ext_cost = False
    # ----------------------------
    # Abracon special processing.
    # ----------------------------
    if princ == 'ABR':
        # Use the sheet names to figure out what processing needs to be done.
        if 'Adj' in sheet_name:
            # Input missing data. Commission Rate is always 3% here.
            sheet['Commission Rate'] = 0.03
            sheet['Paid-On Revenue'] = pd.to_numeric(sheet['Invoiced Dollars'], errors='coerce') * 0.7
            sheet['Actual Comm Paid'] = sheet['Paid-On Revenue'] * 0.03
            # These are paid on resale.
            sheet['Comm Source'] = 'Resale'
            print('Columns added from Abracon special processing:\n'
                  'Commission Rate, Paid-On Revenue, Actual Comm Paid\n---')
        elif 'MoComm' in sheet_name:
            # Fill down Distributor for their grouping scheme.
            sheet['Reported Distributor'].replace('', np.nan, inplace=True)
            sheet['Reported Distributor'].fillna(method='ffill', inplace=True)
            sheet['Reported Distributor'].fillna('', inplace=True)
            # Paid-On Revenue gets Invoiced Dollars.
            sheet['Paid-On Revenue'] = sheet['Invoiced Dollars']
            sheet['Comm Source'] = 'Resale'
            # Calculate the Commission Rate.
            comm_paid = pd.to_numeric(sheet['Actual Comm Paid'], errors='coerce')
            revenue = pd.to_numeric(sheet['Paid-On Revenue'], errors='coerce')
            comm_rate = round(comm_paid / revenue, 3)
            sheet['Commission Rate'] = comm_rate
            print('Columns added from Abracon special processing:\nCommission Rate\n---')
        else:
            print('Sheet not recognized!\n'
                  'Make sure the tab name contains either MoComm or Adj in the name.\n'
                  'Continuing without extra ABR processing.\n---')
    # -------------------------
    # ISSI special processing.
    # -------------------------
    if princ == 'ISS':
        if 'OEM/POS' in list(sheet):
            for row in sheet.index:
                # Deal with OEM idiosyncrasies.
                if 'OEM' in sheet.loc[row, 'OEM/POS']:
                    # Put Sales Region into City.
                    sheet.loc[row, 'City'] = sheet.loc[row, 'Sales Region']
                    # Check for distributor in Customer
                    cust = sheet.loc[row, 'Reported Customer']
                    dist_name = re.sub(r'[^a-zA-Z0-9]', '', str(cust).lower())
                    # Find matches in the Distributor Abbreviations.
                    dist_matches = [i for i in disty_map['Search Abbreviation']
                                    if i in dist_name]
                    if len(dist_matches) == 1:
                        # Copy to distributor column.
                        try:
                            sheet.loc[row, 'Reported Distributor'] = cust
                        except KeyError:
                            pass
        sheet['Comm Source'] = 'Resale'
    # ------------------------
    # ATS special processing.
    # ------------------------
    if princ == 'ATS':
        # Try setting the Paid-On Revenue as the Invoiced Dollars.
        try:
            sheet['Paid-On Revenue'] = sheet['Invoiced Dollars']
        except KeyError:
            pass
        # Try setting the cost/resale by the distributor.
        try:
            for row in sheet.index:
                dist = str(sheet.loc[row, 'Reported Distributor']).lower()
                # Digikey and Mouser are paid on cost, not resale.
                if 'digi' in dist or 'mous' in dist:
                    sheet.loc[row, 'Comm Source'] = 'Cost'
                else:
                    sheet.loc[row, 'Comm Source'] = 'Resale'
        except KeyError:
            pass
    # ----------------------------
    # Mill-Max special processing.
    # ----------------------------
    if princ == 'MIL':
        invoice_num = True
        try:
            sheet['Invoice Number']
        except KeyError:
            print('Found no Invoice Numbers on this sheet.\n---')
            invoice_num = False
        if ext_cost and not invoice_dollars:
            # Sometimes the Totals are written in the Part Number column.
            sheet.drop(sheet[sheet['Part Number'] == 'Totals'].index, inplace=True)
            sheet.reset_index(drop=True, inplace=True)
            # These commissions are paid on cost.
            sheet['Paid-On Revenue'] = sheet['Ext. Cost']
            sheet['Comm Source'] = 'Cost'
        elif 'Part Number' not in list(sheet) and invoice_num:
            # We need to load in the part number log.
            if os.path.exists(look_dir + '\\Mill-Max Invoice Log.xlsx'):
                millmax_log = pd.read_excel(look_dir + '\\Mill-Max Invoice Log.xlsx', dtype=str).fillna('')
                print('Looking up part numbers from invoice log.\n---')
            else:
                print('No Mill-Max Invoice Log found!\n'
                      'Please make sure the Invoice Log is in the Commission Lookup directory.\n'
                      'Skipping tab.\n---')
                return
            # Input part number from Mill-Max Invoice Log.
            for row in sheet.index:
                if sheet.loc[row, 'Invoice Number']:
                    match = millmax_log['Inv#'] == sheet.loc[row, 'Invoice Number']
                    if sum(match) == 1:
                        part_num = millmax_log[match].iloc[0]['Part Number']
                        sheet.loc[row, 'Part Number'] = part_num
                    else:
                        sheet.loc[row, 'Part Number'] = 'NOT FOUND'
            # These commissions are paid on resale.
            sheet['Comm Source'] = 'Resale'
    # --------------------------
    # Osram special processing.
    # --------------------------
    if princ == 'OSR':
        # For World Star POS tab, enter World Star as the distributor.
        if 'World' in sheet_name:
            sheet['Reported Distributor'] = 'World Star'
        try:
            sheet['Paid-On Revenue'] = sheet['Invoiced Dollars']
        except KeyError:
            pass
        sheet['Comm Source'] = 'Resale'
    # --------------------------
    # Cosel special processing.
    # --------------------------
    if princ == 'COS':
        # Only work with the Details tab.
        if sheet_name == 'Details' and ext_cost:
            print('Calculating commissions as 5% of Cost Ext.\n'
                  'For Allied shipments, 4.9% of Cost Ext.\n---')
            # Revenue is from cost.
            sheet['Paid-On Revenue'] = sheet['Ext. Cost']
            for row in sheet.index:
                ext_cost = sheet.loc[row, 'Ext. Cost']
                if sheet.loc[row, 'Reported Distributor'] == 'ALLIED':
                    sheet.loc[row, 'Commission Rate'] = 0.049
                    sheet.loc[row, 'Actual Comm Paid'] = 0.049 * ext_cost
                else:
                    sheet.loc[row, 'Commission Rate'] = 0.05
                    sheet.loc[row, 'Actual Comm Paid'] = 0.05 * ext_cost
        sheet['Comm Source'] = 'Cost'
    # ----------------------------
    # Globtek special processing.
    # ----------------------------
    if princ == 'GLO':
        try:
            sheet['Paid-On Revenue'] = sheet['Invoiced Dollars']
        except KeyError:
            print('No Invoiced Dollars found on this sheet!\n')
        if 'Commission Rate' not in sheet.columns:
            sheet['Commission Rate'] = 0.05
        if 'Actual Comm Paid' not in sheet.columns:
            try:
                sheet['Actual Comm Paid'] = sheet['Paid-On Revenue'] * 0.05
            except KeyError:
                print('No Paid-On Revenue found, could not calculate '
                      'Actual Comm Paid.\n---')
                return
        sheet['Comm Source'] = 'Resale'
    # -------------------------------------------------
    # Generic special processing for a few principals.
    # -------------------------------------------------
    if princ in ['QRF', 'GAN', 'SUR', 'XMO', 'TRI']:
        try:
            sheet['Paid-On Revenue'] = sheet['Invoiced Dollars']
        except KeyError:
            pass
        sheet['Comm Source'] = 'Resale'
    # ---------------------------
    # INF/LAT special processing.
    # ---------------------------
    if princ in ['INF', 'LAT']:
        sheet['Comm Source'] = 'Resale'


def main(filepaths, path_to_running_com, field_mappings):
    """Processes commission files and appends them to Running Commissions.

    Columns in individual commission files are identified and appended to the
    Running Commissions under the appropriate column, as identified by the
    field_mappings file. Entries are then passed through the Lookup Master in
    search of a match to Reported Customer and Part Number. Distributors are
    corrected to consistent names. Entries with missing information are copied
    to Entries Need Fixing for further attention.

    Arguments:
    filepaths -- paths for opening (Excel) files to process.
    path_to_running_com -- current Running Commissions file (in Excel) onto which we are
                  appending data.
    field_mappings -- dataframe which links Running Commissions columns to
                     file data columns.
    """
    # Grab lookup table data names.
    column_names = list(field_mappings)
    # Add in non-lookup'd data names.
    column_names[0:0] = ['CM Sales', 'Design Sales']
    column_names[3:3] = ['T-Name', 'CM', 'T-End Cust']
    column_names[7:7] = ['Principal', 'Distributor']
    column_names[18:18] = ['CM Sales Comm', 'Design Sales Comm', 'Sales Commission']
    column_names[22:22] = ['Quarter Shipped', 'Month', 'Year']
    column_names.extend(['CM Split', 'Paid Date', 'From File', 'Sales Report Date', 'T-Notes', 'Unique ID'])
    # Make sure that column names are unique.
    seen = []
    for col_name in column_names:
        if col_name not in seen:
            seen.append(col_name)
    column_names = seen

    # -------------------------------------------------------------------
    # Check to see if there's an existing Running Commissions to append
    # the new data onto. If so, we need to do some work to get it ready.
    # -------------------------------------------------------------------
    if path_to_running_com:
        running_com, files_processed = load_run_com(file_path=path_to_running_com)
        print('Appending files to Running Commissions...')
        # Make sure column names all match.
        if set(list(running_com)) != set(column_names):
            missing_cols = [i for i in column_names if i not in running_com]
            if not missing_cols:
                missing_cols = ['(None)']
            extra_cols = [i for i in running_com if i not in column_names]
            if not extra_cols:
                extra_cols = ['(None)']
            print('---\nColumns in Running Commissions do not match field_mappings.xlsx!\n'
                  'Missing columns:\n%s' % ', '.join(map(str, missing_cols))
                  + '\nExtra (erroneous) columns:\n%s' % ', '.join(map(str, extra_cols))
                  + '\n*Program terminated*')
            return
        # Load in the matching Entries Need Fixing file.
        ENF_path = out_dir + '\\Entries Need Fixing ' + path_to_running_com[-20:]
        entries_need_fixing = load_entries_need_fixing(file_dir=ENF_path)
        if any([running_com.empty, files_processed.empty]):
            print('Running commissions and/or files processed are empty!\n*Program Terminated*')
            return
        elif entries_need_fixing is None:
            print('No Entries Need Fixing found for the provided Running Commissions!\nProgram Terminated')
        run_com_len = len(running_com)
    # Start new Running Commissions.
    else:
        print('No Running Commissions file provided. Starting a new one.')
        run_com_len = 0
        running_com = pd.DataFrame(columns=column_names)
        entries_need_fixing = pd.DataFrame(columns=column_names)
        files_processed = pd.DataFrame(columns=['Filename', 'Total Commissions', 'Date Added', 'Paid Date'])

    # -------------------------------------------------------------------
    # Check to make sure we aren't duplicating files, then load in data.
    # -------------------------------------------------------------------
    # Strip the root off of the filepaths and leave just the filenames.
    filenames = [os.path.basename(val) for val in filepaths]
    duplicates = list(set(filenames).intersection(files_processed['Filename']))
    filenames = [val for val in filenames if val not in duplicates]
    if duplicates:
        # Let us know we found duplictes and removed them.
        print('---\nThe following files are already in Running Commissions:\n%s' %
              ', '.join(map(str, duplicates)) + '\nDuplicate files were removed from processing.')
        # Exit if no new files are left.
        if not filenames:
            print('---\nNo new commissions files selected.\n'
                  'Please try selecting files again.\n'
                  '*Program terminated*')
            return
    # Read in each new file with Pandas and store them as dictionaries.
    # Each dictionary has a dataframe for each sheet in the file.
    input_data = [pd.read_excel(filepath, None, dtype=str) for filepath in filepaths]

    # --------------------------------------------------------------
    # Read in disty_map. Terminate if not found or if errors in file.
    # --------------------------------------------------------------
    if os.path.exists(look_dir + '\\distributorLookup.xlsx'):
        try:
            disty_map = pd.read_excel(look_dir + '\\distributorLookup.xlsx', 'Distributors')
        except XLRDError:
            print('---\nError reading sheet name for distributorLookup.xlsx!\n'
                  'Please make sure the main tab is named Distributors.\n'
                  '*Program terminated*')
            return
        # Check the column names.
        disty_map_cols = ['Corrected Dist', 'Search Abbreviation']
        missing_cols = [i for i in disty_map_cols if i not in list(disty_map)]
        if missing_cols:
            print('The following columns were not detected in distributorLookup.xlsx:\n%s' %
                  ', '.join(map(str, missing_cols)) + '\n*Program terminated*')
            return
    else:
        print('---\nNo distributor lookup file found!\n'
              'Please make sure distributorLookup.xlsx is in the directory.\n'
              '*Program terminated*')
        return

    # Read in the Lookup Master. Terminate if not found or if errors in file.
    master_lookup = load_lookup_master()

    # -------------------------------------------------------------------------
    # Done loading in the data and supporting files, now go to work.
    # Iterate through each file that we're appending to Running Commissions.
    # -------------------------------------------------------------------------
    for file_num, filename in enumerate(filenames):
        # Grab the next file from the list.
        new_data = input_data[file_num]
        print('_' * 54 + '\nWorking on file: ' + filename + '\n' + '_' * 54)
        # Initialize total commissions for this file.
        total_comm = 0

        # -------------------------------------------------------------------
        # Detect principal from filename, terminate if not on approved list.
        # -------------------------------------------------------------------
        principal = filename[0:3]
        print('Principal detected as: ' + principal)
        principal_list = ['ABR', 'ACH', 'ATP', 'ATS', 'ATO', 'COS', 'EVE', 'GLO', 'INF', 'ISS', 'LAT', 'MIL', 'OSR',
                          'QRF', 'SUR', 'TOP', 'TRI', 'TRU', 'XMO', 'MPS', 'NET', 'GAN', 'EFI']
        if principal not in principal_list:
            print('Principal supplied is not valid!\n'
                  'Current valid principals: ' + ', '.join(map(str, principal_list))
                  + '\nRemember to capitalize the principal abbreviation at start of filename.'
                    '\n*Program terminated*')
            return

        # ----------------------------------------------------------------
        # Iterate over each dataframe in the ordered dictionary.
        # Each sheet in the file is its own dataframe in the dictionary.
        # ----------------------------------------------------------------
        for sheet_name in list(new_data):
            # Rework the index just in case it got read in wrong.
            sheet = new_data[sheet_name].reset_index(drop=True)
            sheet.index = sheet.index.map(int)
            sheet.replace('nan', '', inplace=True)
            # Create a duplicate of the sheet that stays unchanged, aside
            # from recording matches.
            raw_sheet = sheet.copy(deep=True)
            # Figure out if we've already added in the matches row.
            if filename.split('.')[0][-7:] != 'Matched':
                raw_sheet.index += 1
            # Strip whitespace from column names.
            sheet.rename(columns=lambda x: str(x).strip(), inplace=True)
            # Clear out unnamed columns. Attribute error means it's an empty
            # sheet, so simply pass it along (it'll get dealt with).
            try:
                sheet = sheet.loc[:, ~sheet.columns.str.contains('^Unnamed')]
            except AttributeError:
                pass
            # Do specialized pre-processing tailored to principal.
            rename_dict = preprocess_by_principal(principal, sheet, sheet_name)
            # Iterate over each column of data that we want to append.
            for data_name in list(field_mappings):
                # Grab list of names that the data could potentially be under.
                name_list = field_mappings[data_name].dropna().tolist()
                # Look for a match in the sheet column names.
                column_name = [val for val in sheet.columns if val in name_list]
                # If we found too many columns that match, then rename the column in the sheet to the master name.
                if len(column_name) > 1:
                    print('Found multiple matches for ' + data_name
                          + '\nMatching columns: %s' % ', '.join(map(str, column_name))
                          + '\nPlease fix column names and try again.\n*Program terminated*')
                    return
                elif len(column_name) == 1:
                    sheet.rename(columns={column_name[0]: data_name}, inplace=True)
                    if column_name[0] in rename_dict.values():
                        column_name[0] = [i for i in rename_dict.keys() if rename_dict[i] == column_name[0]][0]
                    raw_sheet.loc[0, column_name[0]] = data_name

            # Replace the old raw data sheet with the new one.
            raw_sheet.sort_index(inplace=True)
            new_data[sheet_name] = raw_sheet

            # Convert applicable columns to numeric.
            numeric_cols = ['Quantity', 'Ext. Cost', 'Invoiced Dollars', 'Paid-On Revenue', 'Actual Comm Paid',
                            'Unit Cost', 'Unit Price']
            for col in numeric_cols:
                try:
                    sheet[col] = pd.to_numeric(sheet[col], errors='coerce').fillna('')
                except KeyError:
                    pass

            # Fix Commission Rate if it got read in as a decimal.
            pct_cols = ['Commission Rate', 'Split Percentage', 'Gross Rev Reduction', 'Shared Rev Tier Rate']
            for pct_col in pct_cols:
                try:
                    # Remove '%' sign if present.
                    col = sheet[pct_col].astype(str).map(lambda x: x.strip('%'))
                    # Convert to numeric.
                    col = pd.to_numeric(col, errors='coerce')
                    # Identify and correct which entries are not decimal.
                    col[col > 1] /= 100
                    sheet[pct_col] = col.fillna(0)
                except (KeyError, TypeError):
                    pass

            # Do special processing for principal, if applicable.
            process_by_principal(princ=principal, sheet=sheet, sheet_name=sheet_name, disty_map=disty_map)
            # Drop entries with emtpy part number or reported customer.
            try:
                sheet.drop(sheet[sheet['Part Number'] == ''].index, inplace=True)
                sheet.reset_index(drop=True, inplace=True)
            except KeyError:
                pass

            # Now that we've renamed all of the relevant columns,
            # append the new sheet to Running Commissions, where only the
            # properly named columns are appended.
            if sheet.columns.duplicated().any():
                dupes = sheet.columns[sheet.columns.duplicated()].unique()
                print('Two items are being mapped to the same column!\n'
                      'These columns contain duplicates: %s' % ', '.join(map(str, dupes))
                      + '\n*Program terminated*')
                return
            elif 'Actual Comm Paid' not in list(sheet):
                # Tab has no commission data, so it is ignored.
                print('No commission dollars column found on this tab.\n'
                      'Skipping tab.\n-')
            elif 'Part Number' not in list(sheet):
                # Tab has no part number data, so it is ignored.
                print('No part number column found on this tab.\n'
                      'Skipping tab.\n-')
            elif 'Invoice Date' not in list(sheet):
                # Tab has no date column, so report and exit.
                print('No Invoice Date column found for this tab.\n'
                      'Please make sure the Invoice Date is mapped.\n'
                      '*Program terminated*')
                return
            else:
                # Report the number of rows that have part numbers.
                total_rows = sum(sheet['Part Number'] != '')
                print('Found ' + str(total_rows) + ' entries in the tab ' + sheet_name
                      + ' with valid part numbers.\n' + '-'*35)
                # Remove entries with no commissions dollars.
                sheet['Actual Comm Paid'] = pd.to_numeric(sheet['Actual Comm Paid'], errors='coerce').fillna(0)
                sheet = sheet[sheet['Actual Comm Paid'] != 0]

                # Add 'From File' column to track where data came from.
                sheet['From File'] = filename
                # Fill in principal.
                sheet['Principal'] = principal

                # Find matching columns.
                matching_columns = [val for val in list(sheet) if val in list(field_mappings)]
                if len(matching_columns) > 0:
                    # Sum commissions paid on sheet.
                    print('Commissions for this tab: '
                          + '${:,.2f}'.format(sheet['Actual Comm Paid'].sum()) + '\n-')
                    total_comm += sheet['Actual Comm Paid'].sum()
                    # Strip whitespace from all strings in dataframe.
                    string_cols = [val for val in list(sheet) if sheet[val].dtype == 'object']
                    for col in string_cols:
                        sheet[col] = sheet[col].fillna('').astype(str).map(lambda x: x.strip())
                    # Append matching columns of data.
                    app_cols = matching_columns + ['From File', 'Principal']
                    running_com = running_com.append(sheet[app_cols], ignore_index=True, sort=False)
                else:
                    print('Found no data on this tab. Moving on.\n-')

        # Show total commissions.
        print('Total commissions for this file: ${:,.2f}'.format(total_comm))
        # Append filename and total commissions to Files Processed sheet.
        current_date = datetime.datetime.now().date()
        new_file = pd.DataFrame({'Filename': [filename], 'Total Commissions': [total_comm],
                                 'Date Added': [current_date], 'Paid Date': ['']})
        files_processed = files_processed.append(new_file, ignore_index=True, sort=False)
        # Save the matched raw data file.
        fname = filename[:-5]
        if filename[-12:] != 'Matched.xlsx':
            fname += ' Matched.xlsx'
        else:
            fname += '.xlsx'
        if save_error(fname):
            print('---\nOne or more of the raw data files are open in Excel.\n'
                  'Please close these files and try again.\n*Program terminated*')
            return
        # Write the raw data file with matches.
        writer = pd.ExcelWriter(match_dir + '\\' + fname, engine='xlsxwriter', datetime_format='mm/dd/yyyy')
        for tab in list(new_data):
            new_data[tab].to_excel(writer, sheet_name=tab, index=False)
            # Format and fit each column.
            sheet = writer.sheets[tab]
            index = 0
            for col in new_data[tab].columns:
                # Set column width and formatting.
                try:
                    max_width = max(len(str(val)) for val in new_data[tab][col].values)
                except ValueError:
                    max_width = 0
                max_width = max(10, max_width)
                sheet.set_column(index, index, max_width + 0.8)
                index += 1
        # Save the file.
        writer.save()

    # Fill NaNs left over from appending.
    running_com.fillna('', inplace=True)
    # Find matches in Lookup Master and extract data from them.
    # Let us know how many rows are being processed.
    num_rows = '{:,.0f}'.format(len(running_com) - run_com_len)
    if num_rows == '0':
        print('---\nNo new valid data provided.\n'
              'Please check the new files for missing data or column matches.\n'
              '*Program terminated*')
        return
    print('---\nBeginning processing on ' + num_rows + ' rows of data.')
    running_com.reset_index(inplace=True, drop=True)

    # Iterate over each row of the newly appended data.
    for row in range(run_com_len, len(running_com)):
        # ------------------------------------------
        # Try to find a match in the Lookup Master.
        # ------------------------------------------
        # First assign a new Unique ID to this entry.
        running_com.loc[row, 'Unique ID'] = uuid4()
        lookup_matches = 0
        # Don't look up correction lines.
        if 'correction' not in str(running_com.loc[row, 'T-Notes']).lower():
            # First match reported customer.
            reported_cust = str(running_com.loc[row, 'Reported Customer']).lower()
            POS_cust = master_lookup['Reported Customer'].map(lambda x: str(x).lower())
            cust_matches = master_lookup[reported_cust == POS_cust]
            # Now match part number.
            part_num = str(running_com.loc[row, 'Part Number']).lower()
            PPN = master_lookup['Part Number'].map(lambda x: str(x).lower())
            # Reset index, but keep it around for updating usage below.
            full_match = cust_matches[part_num == PPN].reset_index()
            # Record number of Lookup Master matches.
            lookup_matches = len(full_match)
            # If we found one match we're good, so copy it over.
            if lookup_matches == 1:
                full_match = full_match.iloc[0]
                # If there are no salespeople, it means we found a
                # "soft match."
                # These have unknown End Customers and should go to
                # Entries Need Fixing. So, set them to zero matches.
                if full_match['CM Sales'] == full_match['Design Sales'] == '':
                    lookup_matches = 0
                # Grab primary and secondary sales people from Lookup Master.
                running_com.loc[row, 'CM Sales'] = full_match['CM Sales']
                running_com.loc[row, 'Design Sales'] = full_match['Design Sales']
                running_com.loc[row, 'T-Name'] = full_match['T-Name']
                running_com.loc[row, 'CM'] = full_match['CM']
                running_com.loc[row, 'T-End Cust'] = full_match['T-End Cust']
                running_com.loc[row, 'CM Split'] = full_match['CM Split']
                # Update usage in lookup Master.
                master_lookup.loc[full_match['index'], 'Last Used'] = datetime.datetime.now().date()
                # Update OOT city if not already filled in.
                if full_match['T-Name'][0:3] == 'OOT' and not full_match['City']:
                    master_lookup.loc[full_match['index'], 'City'] = running_com.loc[row, 'City']
            # If we found multiple matches, then fill in all the options.
            elif lookup_matches > 1:
                lookup_cols = ['CM Sales', 'Design Sales', 'T-Name', 'CM', 'T-End Cust', 'CM Split']
                # Write list of all unique entries for each column.
                for col in lookup_cols:
                    running_com.loc[row, col] = ', '.join(full_match[col].map(lambda x: str(x)).unique())

        # -----------------------------------------------------------
        # Format the date correctly and fill in the Quarter Shipped.
        # -----------------------------------------------------------
        # Try parsing the date.
        date_error = False
        date_given = running_com.loc[row, 'Invoice Date']
        # Check if the date is read in as a float/int, and convert to string.
        if isinstance(date_given, (float, int)):
            date_given = str(int(date_given))
        # Check if Pandas read it in as a Timestamp object.
        # If so, turn it back into a string (a bit roundabout, oh well).
        elif isinstance(date_given, (pd.Timestamp, datetime.datetime)):
            date_given = str(date_given)
        try:
            parse(date_given)
        except (ValueError, TypeError):
            # The date isn't recognized by the parser.
            date_error = True
        except KeyError:
            print('---\nThere is no Invoice Date column in Running Commissions!\n'
                  'Please check to make sure an Invoice Date column exists.\n'
                  'Note: Spelling, whitespace, and capitalization matter.\n---')
            date_error = True
        # If no error found in date, fill in the month/year/quarter
        if not date_error:
            date = parse(date_given).date()
            # Make sure the date actually makes sense.
            current_year = int(time.strftime('%Y'))
            if current_year - date.year not in [0, 1] or date > current_date:
                date_error = True
            else:
                # Cast date format into mm/dd/yyyy.
                running_com.loc[row, 'Invoice Date'] = date
                # Fill in quarter/year/month data.
                running_com.loc[row, 'Year'] = date.year
                month = calendar.month_name[date.month][0:3]
                running_com.loc[row, 'Month'] = month
                qtr = str(math.ceil(date.month / 3))
                running_com.loc[row, 'Quarter Shipped'] = (str(date.year) + 'Q' + qtr)

        # ---------------------------------------------------
        # Try to correct the distributor to consistent name.
        # ---------------------------------------------------
        # Strip extraneous characters and all spaces, and make lowercase.
        reported_dist = str(running_com.loc[row, 'Reported Distributor'])
        dist_name = re.sub(r'[^a-zA-Z0-9]', '', reported_dist).lower()

        # Find matches for the dist_name in the Distributor Abbreviations.
        dist_matches = [i for i in disty_map['Search Abbreviation'] if i in dist_name]
        if len(dist_matches) == 1:
            # Find and input corrected distributor name.
            match_loc = disty_map['Search Abbreviation'] == dist_matches[0]
            corrected_dist = disty_map[match_loc].iloc[0]['Corrected Dist']
            running_com.loc[row, 'Distributor'] = corrected_dist
        elif not dist_name:
            running_com.loc[row, 'Distributor'] = ''
            dist_matches = ['Empty']

        # -----------------------------------------------------------------
        # Go through each column and convert applicable entries to numeric.
        # -----------------------------------------------------------------
        cols = list(running_com)
        # Invoice number sometimes has leading zeros we'd like to keep.
        cols.remove('Invoice Number')
        # The INF gets read in as infinity, so skip the principal column.
        cols.remove('Principal')
        for col in cols:
            running_com.loc[row, col] = pd.to_numeric(running_com.loc[row, col], errors='ignore')

        # -----------------------------------------------------------------
        # If any data isn't found/parsed, copy over to Entries Need Fixing.
        # -----------------------------------------------------------------
        if lookup_matches != 1 or len(dist_matches) != 1 or date_error:
            entries_need_fixing = entries_need_fixing.append(running_com.loc[row, :], sort=False)
            entries_need_fixing.loc[row, 'Running Com Index'] = row
            entries_need_fixing.loc[row, 'Distributor Matches'] = len(dist_matches)
            entries_need_fixing.loc[row, 'Lookup Master Matches'] = lookup_matches
            entries_need_fixing.loc[row, 'Date Added'] = datetime.datetime.now().date()
        else:
            # Fill in the Sales Commission info.
            sales_com = 0.45 * running_com.loc[row, 'Actual Comm Paid']
            running_com.loc[row, 'Sales Commission'] = sales_com
            if running_com.loc[row, 'CM Sales']:
                # Grab split with default to 20.
                split = running_com.loc[row, 'CM Split'] or 20
            else:
                # No CM Sales, so no split.
                split = 0
            running_com.loc[row, 'CM Sales Comm'] = split * sales_com / 100
            running_com.loc[row, 'Design Sales Comm'] = (100 - split) * sales_com / 100

        # Update progress every 100 rows.
        if (row - run_com_len) % 100 == 0 and row > run_com_len:
            print('Done with row ' '{:,.0f}'.format(row - run_com_len))
    # -----------------------------
    # Clean up the finalized data.
    # -----------------------------
    # Reorder columns to match the desired layout in column_names.
    running_com.fillna('', inplace=True)
    running_com = running_com.loc[:, column_names]
    column_names.extend(['Distributor Matches', 'Lookup Master Matches', 'Date Added', 'Running Com Index',
                         'Unique ID'])
    # Fix up the Entries Need Fixing file.
    entries_need_fixing = entries_need_fixing.loc[:, column_names]
    entries_need_fixing.reset_index(drop=True, inplace=True)
    entries_need_fixing.fillna('', inplace=True)
    # Make sure all the dates are formatted correctly.
    running_com['Invoice Date'] = running_com['Invoice Date'].map(lambda x: form_date(x))
    entries_need_fixing['Invoice Date'] = entries_need_fixing['Invoice Date'].map(lambda x: form_date(x))
    entries_need_fixing['Date Added'] = entries_need_fixing['Date Added'].map(lambda x: form_date(x))
    master_lookup['Last Used'] = master_lookup['Last Used'].map(lambda x: form_date(x))
    master_lookup['Date Added'] = master_lookup['Date Added'].map(lambda x: form_date(x))

    # Get ready to save files.
    current_time = time.strftime('%Y-%m-%d-%H%M')
    fname1 = out_dir + '\\Running Commissions ' + current_time + '.xlsx'
    fname2 = out_dir + '\\Entries Need Fixing ' + current_time + '.xlsx'
    fname3 = look_dir + '\\Lookup Master - Current.xlsx'
    writer1 = prepare_save_file(filename=fname1, tab_data=[running_com, files_processed],
                                tab_names=['Master', 'Files Processed'])
    writer2 = prepare_save_file(filename=fname2, tab_data=entries_need_fixing, tab_names='Data')
    writer3 = prepare_save_file(filename=fname3, tab_data=master_lookup, tab_names='Lookup')

    success = save_files([writer1, writer2, writer3])
    if success:
        print('---\nUpdates completed successfully!\n---\nRunning Commissions updated.\n'
              'Lookup Master updated.\nEntries Need Fixing updated.\n+Program Complete+')
