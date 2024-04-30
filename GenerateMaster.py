import pandas as pd
import numpy as np
from dateutil.parser import parse
import time
import calendar
import math
import os.path
import re
import datetime
import logging
from uuid import uuid4
import GenerateMasterUtils as Utils
from FileIO import (load_lookup_master, load_run_com, load_entries_need_fixing, load_principal_info,
                    load_distributor_map, save_excel_file)
from RCExcelTools import save_error, form_date
from PrincipalSpecialProcessing import process_by_principal, preprocess_by_principal

logger = logging.getLogger(__name__)


def main(filepaths, path_to_running_com, field_mappings):
    """
    Processes commission files and appends them to Running Commissions.

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
    logger.info('Starting program: Generate Master')
    # -------------------------------------------------------------------
    # Check to see if there's an existing Running Commissions to append
    # the new data onto. If so, we need to do some work to get it ready.
    # -------------------------------------------------------------------
    # Get the correct column names for the commission file.
    column_names = Utils.get_column_names(field_mappings)

    if path_to_running_com:
        running_com, files_processed = load_run_com(file_path=path_to_running_com)
        logger.info(f'Appending files to {path_to_running_com}...')

        # Check to make sure that all columns are present and match between the files
        missing_cols = [i for i in column_names if i not in running_com]
        extra_cols = [i for i in running_com if i not in column_names]
        if missing_cols or extra_cols:
            logger.error('Columns in Running Commissions do not match field_mappings.xlsx!\n'
                         f'Missing columns:\n{', '.join(map(str, missing_cols))}'
                         f'\nExtra (erroneous) columns:\n{', '.join(map(str, extra_cols))}\n*Program Terminated*')
            return

        # Load in the matching Entries Need Fixing file.
        ENF_path = os.path.join(Utils.DIRECTORIES.get('COMM_WORKING_DIR'),
                                f'Entries Need Fixing {path_to_running_com[-20:]}')
        entries_need_fixing = load_entries_need_fixing(file_dir=ENF_path)

        if any([running_com.empty, files_processed.empty]):
            logger.error('Running commissions and/or files processed are empty!\n*Program Terminated*')
            return
        elif entries_need_fixing is None:
            logger.error('No Entries Need Fixing found for the provided Running Commissions!\n*Program Terminated*')
            return
        running_com_input_len = running_com.shape[0]
    # ---------------------------------------------------------
    # Start new Running Commissions; no existing one provided.
    # ---------------------------------------------------------
    else:
        logger.info('No Running Commissions file provided. Starting a new one.')
        running_com_input_len = 0
        running_com = pd.DataFrame(columns=column_names)
        entries_need_fixing = pd.DataFrame(columns=column_names)
        files_processed = pd.DataFrame(columns=['Filename', 'Total Commissions', 'Date Added', 'Paid Date'])

    filenames = Utils.filter_duplicate_files(filepaths, files_processed)
    # Exit if no new files are left after filtering.
    if not filenames:
        logger.error('No new commissions files selected. Please try selecting files again.\n*Program terminated*')
        return

    # Read in each new file with Pandas and store them as a list of dictionaries.
    # Each dictionary has a dataframe for each sheet in the file.
    input_data = [pd.read_excel(filepath, sheet_name=None, dtype=str) for filepath in filepaths]

    # Load the supporting files.
    distributor_map = load_distributor_map()
    master_lookup = load_lookup_master()
    principal_info = load_principal_info()
    if any([distributor_map.empty, master_lookup.empty, principal_info.empty]):
        logger.error('Error loading supporting files.\n*Program terminated*')
        return
    principal_list = principal_info['Abbreviation'].to_list()

    # -------------------------------------------------------------------------
    # Done loading in the data and supporting files, now go to work.
    # Iterate through each file that we're appending to Running Commissions.
    # -------------------------------------------------------------------------
    for file_num, filename in enumerate(filenames):
        # Grab the next file from the list.
        new_data = input_data[file_num]
        logger.info(f'Working on file: {filename}')
        # Initialize total commissions for this file.
        total_comm = 0

        # Detect principal from filename, terminate if not on approved list.
        principal = filename[0:3]
        logger.info(f'Principal detected as: {principal}')
        if principal not in principal_list:
            logger.error(f'Principal supplied is not valid! Current valid principals: '
                         f'{', '.join(map(str, principal_list))}\nRemember to capitalize the principal abbreviation '
                         f'at start of filename.\n*Program terminated*')
            return

        # ----------------------------------------------------------------
        # Iterate over each dataframe in the ordered dictionary.
        # Each sheet in the file is its own dataframe in the dictionary.
        # ----------------------------------------------------------------
        for sheet_name in list(new_data):
            # Rework the index just in case it got read in wrong, then clean up the dataframe.
            sheet = new_data[sheet_name].reset_index(drop=True)
            sheet.index = sheet.index.map(int)
            sheet.replace(to_replace=['nan', np.nan], value='', inplace=True)
            sheet.rename(columns=lambda x: str(x).strip(), inplace=True)

            # Clear out unnamed columns.
            try:
                sheet = sheet.loc[:, ~sheet.columns.str.contains('^Unnamed')]
            except AttributeError:
                # It's an empty dataframe, so simply pass it along (it'll get dealt with).
                pass

            if sheet.empty:
                logger.info(f'Skipping empty sheet {sheet_name}')
                continue

            # Do specialized pre-processing tailored to principal (mostly renaming columns).
            preprocess_by_principal(principal=principal, sheet=sheet, sheet_name=sheet_name)

            # Iterate over each column of data that we want to append.
            for data_name in list(field_mappings):
                # Grab list of names that the data could potentially be under.
                name_list = field_mappings[data_name].dropna().tolist()
                # Look for a match in the sheet column names.
                column_name = [val for val in sheet.columns if val in name_list]
                # If we found too many columns that match, then rename the column in the sheet to the master name.
                if len(column_name) > 1:
                    logger.error(f'Found multiple mappings for {data_name}'
                                 f'\nMatching columns: {', '.join(map(str, column_name))}'
                                 '\nPlease fix column names and try again.\n*Program terminated*')
                    return
                elif len(column_name) == 1:
                    sheet.rename(columns={column_name[0]: data_name}, inplace=True)

            sheet = Utils.format_pct_numeric_cols(dataframe=sheet)

            # Do special processing for principal, if applicable.
            process_by_principal(principal=principal, sheet=sheet, sheet_name=sheet_name, disty_map=distributor_map)

            # Drop entries with emtpy part number, or skip tab if no part number column is found.
            try:
                num_entries = sheet['Part Number'].shape[0]
                sheet.drop(sheet[sheet['Part Number'] == ''].index, inplace=True)
                sheet.reset_index(drop=True, inplace=True)
                if sheet['Part Number'].shape[0] < num_entries:
                    logger.info(f'Dropped {num_entries - sheet['Part Number'].shape[0]:,} lines with no part number.')
            except KeyError:
                logger.warning(f'No part number column found on tab {sheet_name}. Skipping tab.')
                continue

            # Now that we've renamed all of the relevant columns,
            # append the new sheet to Running Commissions, where only the properly named columns are appended.
            if sheet.columns.duplicated().any():
                duplicates = sheet.columns[sheet.columns.duplicated()].unique()
                logger.error('Two items are being mapped to the same column!\n'
                             f'These columns contain duplicates: {', '.join(map(str, duplicates))}'
                             f'\n*Program terminated*')
                return
            elif 'Actual Comm Paid' not in list(sheet):
                # Tab has no commission data, so it is ignored.
                logger.warning(f'No commission dollars column found on tab {sheet_name}. Skipping tab.')
            elif 'Invoice Date' not in list(sheet):
                # Tab has no date column, so report and exit.
                logger.warning(f'No Invoice Date column found on tab {sheet_name}. Skipping tab.')
            else:
                logger.info(f'Found {sheet.shape[0]} entries in the tab {sheet_name} with valid part numbers.')
                # Remove entries with no commissions dollars.
                sheet['Actual Comm Paid'] = pd.to_numeric(sheet['Actual Comm Paid'], errors='coerce').fillna(0)
                sheet = sheet[sheet['Actual Comm Paid'] != 0]
                # Add 'From File' column to track where data came from.
                sheet['From File'] = filename
                # Fill in the principal.
                sheet['Principal'] = principal

                # Find matching columns.
                matching_columns = [val for val in list(sheet) if val in list(field_mappings)]
                if len(matching_columns) > 0:
                    # Sum commissions paid on sheet.
                    logger.info(f'Commissions for this tab: ${sheet['Actual Comm Paid'].sum():,.2f}')
                    total_comm += sheet['Actual Comm Paid'].sum()
                    # Strip whitespace from all strings in dataframe.
                    string_cols = [val for val in list(sheet) if sheet[val].dtype == 'object']
                    for col in string_cols:
                        sheet[col] = sheet[col].fillna('').astype(str).map(lambda x: x.strip())
                    # Append matching columns of data.
                    app_cols = matching_columns + ['From File', 'Principal']
                    running_com = pd.concat((running_com, sheet[app_cols]), ignore_index=True, sort=False)
                else:
                    logger.info(f'Found no data tab {sheet_name}. Skipping.')

        # Show total commissions.
        logger.info(f'Total commissions for {filename}: ${total_comm:,.2f}')
        # Append filename and total commissions to Files Processed sheet.
        new_file = pd.DataFrame({'Filename': [filename], 'Total Commissions': [total_comm],
                                 'Date Added': [datetime.datetime.now().date()], 'Paid Date': ['']})
        files_processed = pd.concat((files_processed, new_file), ignore_index=True, sort=False)

    # ----------------------------------------------------------------
    # Done appending new data, now find matches in the Lookup Master.
    # ----------------------------------------------------------------
    # Let us know how many rows are being processed.
    if running_com.shape[0] - running_com_input_len <= 0:
        logger.error('No new valid data provided. Please check the new files for missing data or column matches.\n'
                     '*Program terminated*')
        return
    logger.info(f'Beginning processing on {running_com.shape[0] - running_com_input_len} rows of data.')

    # Fill NaNs left over from appending.
    running_com.replace(to_replace=np.nan, value='', inplace=True)
    running_com.reset_index(inplace=True, drop=True)

    # Get the part numbers and customers out of the lookup for use below.
    lookup_customers = master_lookup['Reported Customer'].map(lambda x: str(x).lower())
    lookup_part_numbers = master_lookup['Part Number'].map(lambda x: str(x).lower())

    # Iterate over each row of the newly appended data.
    for row in range(running_com_input_len, len(running_com)):
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
            cust_matches = reported_cust == lookup_customers
            # Now match part number.
            part_num = str(running_com.loc[row, 'Part Number']).lower()
            part_matches = part_num == lookup_part_numbers
            # Reset index, but keep it around for updating usage below.
            full_match = master_lookup[cust_matches & part_matches].reset_index()
            # Record number of Lookup Master matches.
            lookup_matches = full_match.shape[0]

            # If we found one match we're good, so copy it over.
            if lookup_matches == 1:
                full_match = full_match.iloc[0]
                # If there are no salespeople, it means we found a "soft match."
                # These have unknown End Customers and should go to Entries Need Fixing.
                # So, set them to zero matches.
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
                master_lookup.loc[full_match['index'], 'Last Used'] = pd.to_datetime(datetime.datetime.now().date())
                # Update OOT city if not already filled in.
                if full_match['T-Name'][0:3] == 'OOT' and not full_match['City']:
                    master_lookup.loc[full_match['index'], 'City'] = running_com.loc[row, 'City']
            # If we found multiple matches, then fill in all the options and let the user fix later.
            elif lookup_matches > 1:
                lookup_cols = ['CM Sales', 'Design Sales', 'T-Name', 'CM', 'T-End Cust', 'CM Split']
                # Write list of all unique entries for each column.
                for col in lookup_cols:
                    running_com.loc[row, col] = ', '.join(full_match[col].map(lambda x: str(x)).unique())

        # -----------------------------------------------------------
        # Format the date correctly and fill in the Quarter Shipped.
        # -----------------------------------------------------------
        # Try parsing the date.
        invoice_date = running_com.loc[row, 'Invoice Date']
        date_error = Utils.check_for_date_errors(date=invoice_date)

        # If no error found in date, fill in the month/year/quarter
        if not date_error:
            date = parse(invoice_date).date()
            # Make sure the date actually makes sense.
            current_year = int(time.strftime('%Y'))
            if current_year - date.year not in [0, 1] or date > datetime.datetime.now().date():
                date_error = True
            else:
                # Cast date format into mm/dd/yyyy.
                running_com.loc[row, 'Invoice Date'] = date
                # Fill in quarter/year/month data.
                running_com.loc[row, 'Year'] = date.year
                month = calendar.month_name[date.month][0:3]
                running_com.loc[row, 'Month'] = month
                running_com.loc[row, 'Quarter Shipped'] = f'{date.year}Q{math.ceil(date.month / 3)}'

        # ---------------------------------------------------
        # Try to correct the distributor to consistent name.
        # ---------------------------------------------------
        # Strip extraneous characters and all spaces, and make lowercase.
        reported_dist = str(running_com.loc[row, 'Reported Distributor'])
        dist_name = re.sub(r'[^a-zA-Z0-9]', '', reported_dist).lower()

        # Find matches for the dist_name in the Distributor Abbreviations.
        dist_matches = [i for i in distributor_map['Search Abbreviation'] if i in dist_name]
        if len(dist_matches) == 1:
            # Find and input corrected distributor name.
            match_loc = distributor_map['Search Abbreviation'] == dist_matches[0]
            corrected_dist = distributor_map[match_loc].iloc[0]['Corrected Dist']
            running_com.loc[row, 'Distributor'] = corrected_dist
        elif not dist_name:
            running_com.loc[row, 'Distributor'] = ''
            dist_matches = ['Empty']

        # -----------------------------------------------------------------
        # If any data isn't found/parsed, copy over to Entries Need Fixing.
        # -----------------------------------------------------------------
        if lookup_matches != 1 or len(dist_matches) != 1 or date_error:
            entries_need_fixing = pd.concat((entries_need_fixing, running_com.loc[row, :]), sort=False)
            entries_need_fixing.loc[row, 'Running Com Index'] = row
            entries_need_fixing.loc[row, 'Distributor Matches'] = len(dist_matches)
            entries_need_fixing.loc[row, 'Lookup Master Matches'] = lookup_matches
            entries_need_fixing.loc[row, 'Date Added'] = pd.to_datetime(datetime.datetime.now().date())
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
        if (row - running_com_input_len) % 100 == 0 and row > running_com_input_len:
            logger.info(f'Done with row {'{:,.0f}'.format(row - running_com_input_len)}')

    # -----------------------------
    # Clean up the finalized data.
    # -----------------------------
    # Reorder columns to match the desired layout in column_names.
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

    logger.info('Saving files.')
    current_time = time.strftime('%Y-%m-%d-%H%M')

    filepath_RC = os.path.join(Utils.DIRECTORIES.get('COMM_WORKING_DIR'), f'Running Commissions {current_time}.xlsx')
    filepath_ENF = os.path.join(Utils.DIRECTORIES.get('COMM_WORKING_DIR'), f'Entries Need Fixing {current_time}.xlsx')
    filepath_LM = os.path.join(Utils.DIRECTORIES.get('COMM_LOOKUPS_DIR'), f'Lookup Master - Current.xlsx')
    if save_error(filepath_RC, filepath_ENF, filepath_LM):
        logger.error('One or more of the RC/ENF/Lookup files are currently open in Excel! '
                     'Please close the files and try again.\n*Program Teminated*')
        return

    save_excel_file(filename=filepath_RC, tab_data=[running_com, files_processed],
                    tab_names=['Master', 'Files Processed'])
    save_excel_file(filename=filepath_ENF, tab_data=entries_need_fixing, tab_names='Data')
    save_excel_file(filename=filepath_LM, tab_data=master_lookup, tab_names='Lookup')
    return True
