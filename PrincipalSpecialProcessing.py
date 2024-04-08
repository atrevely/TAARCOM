import os
import re
import logging
import pandas as pd
import numpy as np

logger = logging.getLogger(__name__)

# Set the directory for the data input/output.
if os.path.exists('Z:\\'):
    OUT_DIR = 'Z:\\MK Working Commissions'
    LOOK_DIR = 'Z:\\Commissions Lookup'
    MATCH_DIR = 'Z:\\Matched Raw Data Files'
else:
    OUT_DIR = os.getcwd()
    LOOK_DIR = os.getcwd()
    MATCH_DIR = os.getcwd()


def preprocess_by_principal(principal, sheet, sheet_name):
    """
    Do special pre-processing tailored to the principal input. Primarily, this involves renaming
    columns that would get looked up incorrectly in the Field Mappings.

    This function modifies a dataframe inplace.
    """
    # Initialize the rename_dict in case it doesn't get set by any matching principal.
    rename_dict = {}

    match principal:
        case 'OSR':
            rename_dict = {'Item': 'Unmapped', 'Material Number': 'Unmapped 2',
                           'Customer Name': 'Unmapped 3', 'Sales Date': 'Unmapped 4'}
            sheet.rename(columns=rename_dict, inplace=True)
            # Combine Rep 1 % and Rep 2 %.
            if 'Rep 1 %' in list(sheet) and 'Rep 2 %' in list(sheet):
                logger.info('Copying Rep 2 % into empty Rep 1 % lines.')
                for row in sheet.index:
                    if sheet.loc[row, 'Rep 2 %'] and not sheet.loc[row, 'Rep 1 %']:
                        sheet.loc[row, 'Rep 1 %'] = sheet.loc[row, 'Rep 2 %']

        case 'ISS':
            rename_dict = {'Commission Due': 'Unmapped', 'Name': 'OEM/POS'}
            sheet.rename(columns=rename_dict, inplace=True)

        case 'ATS':
            rename_dict = {'Resale': 'Extended Resale', 'Cost': 'Extended Cost'}
            sheet.rename(columns=rename_dict, inplace=True)

        case 'QRF':
            if sheet_name in ['OEM', 'OFF']:
                rename_dict = {'End Customer': 'Unmapped 2', 'Item': 'Unmapped 3'}
                sheet.rename(columns=rename_dict, inplace=True)
            elif sheet_name == 'POS':
                rename_dict = {'Company': 'Distributor', 'BillDocNo': 'Unmapped',
                               'End Customer': 'Unmapped 2', 'Item': 'Unmapped 3'}
                sheet.rename(columns=rename_dict, inplace=True)

        case 'XMO':
            rename_dict = {'Amount': 'Commission', 'Commission Due': 'Unmapped'}
            sheet.rename(columns=rename_dict, inplace=True)

    # Return the rename_dict for future use in the matched raw file.
    if rename_dict:
        logger.info(f'The following columns were renamed automatically on this sheet ({principal}):\n'
                    f'{', '.join([f'{i} --> {j}' for i, j in zip(rename_dict.keys(), rename_dict.values())])}')
    return rename_dict


def process_by_principal(principal, sheet, sheet_name, disty_map):
    """
    Do special processing tailored to the principal input. This involves
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

    match principal:
        case 'ABR':
            # Use the sheet names to figure out what processing needs to be done.
            if 'Adj' in sheet_name:
                # Input missing data. Commission Rate is always 3% here.
                sheet['Commission Rate'] = 0.03
                sheet['Paid-On Revenue'] = pd.to_numeric(sheet['Invoiced Dollars'], errors='coerce') * 0.7
                sheet['Actual Comm Paid'] = sheet['Paid-On Revenue'] * 0.03
                # These are paid on resale.
                sheet['Comm Source'] = 'Resale'
                logger.info('Columns added from Abracon special processing:'
                            'Commission Rate, Paid-On Revenue, Actual Comm Paid')
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
                logger.info('Columns added from Abracon special processing: Commission Rate')
            else:
                logger.warning('Sheet not recognized! Make sure the tab name contains either MoComm or Adj in the name.'
                               'Continuing without extra ABR processing.')

        case 'ISS':
            if 'OEM/POS' in list(sheet):
                for row in sheet.index:
                    # Deal with OEM idiosyncrasies.
                    if 'OEM' in sheet.loc[row, 'OEM/POS']:
                        # Put Sales Region into City.
                        sheet.loc[row, 'City'] = sheet.loc[row, 'Sales Region']
                        # Check for distributor in Customer
                        cust = sheet.loc[row, 'Reported Customer']
                        dist_name = re.sub(pattern='[^a-zA-Z0-9]', repl='', string=str(cust).lower())
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

        case 'ATS':
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

        case 'MIL':
            invoice_num = True
            try:
                sheet['Invoice Number']
            except KeyError:
                logger.info('Found no Invoice Numbers on this sheet.')
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
                if os.path.exists(LOOK_DIR + '\\Mill-Max Invoice Log.xlsx'):
                    millmax_log = pd.read_excel(os.path.join(LOOK_DIR, 'Mill-Max Invoice Log.xlsx'), dtype=str)
                    millmax_log = millmax_log.fillna('')
                    logger.info('Looking up part numbers from invoice log.')
                else:
                    logger.warning('No Mill-Max Invoice Log found! Please make sure the Invoice Log is in the '
                                   'Commission Lookup directory. Skipping tab.')
                    return
                # Input part number from Mill-Max Invoice Log.
                for row in sheet.index:
                    if sheet.loc[row, 'Invoice Number']:
                        matches = millmax_log['Inv#'] == sheet.loc[row, 'Invoice Number']
                        if sum(matches) == 1:
                            part_num = millmax_log[matches].iloc[0]['Part Number']
                            sheet.loc[row, 'Part Number'] = part_num
                        else:
                            sheet.loc[row, 'Part Number'] = 'NOT FOUND'
                # These commissions are paid on resale.
                sheet['Comm Source'] = 'Resale'

        case 'OSR':
            # For World Star POS tab, enter World Star as the distributor.
            if 'World' in sheet_name:
                sheet['Reported Distributor'] = 'World Star'
            try:
                sheet['Paid-On Revenue'] = sheet['Invoiced Dollars']
            except KeyError:
                pass
            sheet['Comm Source'] = 'Resale'

        case 'GLO':
            try:
                sheet['Paid-On Revenue'] = sheet['Invoiced Dollars']
            except KeyError:
                logger.info('No Invoiced Dollars found on this sheet!')
            if 'Commission Rate' not in sheet.columns:
                sheet['Commission Rate'] = 0.05
            if 'Actual Comm Paid' not in sheet.columns:
                try:
                    sheet['Actual Comm Paid'] = sheet['Paid-On Revenue'] * 0.05
                except KeyError:
                    logger.warning('No Paid-On Revenue found, could not calculate Actual Comm Paid.')
                    return
            sheet['Comm Source'] = 'Resale'

    # Generic special processing for a few principals.
    if principal in ['QRF', 'GAN', 'XMO', 'TRI']:
        try:
            sheet['Paid-On Revenue'] = sheet['Invoiced Dollars']
        except KeyError:
            pass
        sheet['Comm Source'] = 'Resale'
