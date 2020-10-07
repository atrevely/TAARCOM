import pandas as pd
import numpy as np
import os
from FileLoader import load_acct_list, load_root_customer_mappings, load_salespeople_info


def tableFormat(sheetData, sheetName, wbook):
    """Formats the Excel output as a table with correct column formatting."""
    # Nothing to format, so return.
    if sheetData.shape[0] == 0:
        return
    sheet = wbook.sheets[sheetName]
    sheet.freeze_panes(1, 0)
    # Set the autofilter for the sheet.
    sheet.autofilter(0, 0, sheetData.shape[0], sheetData.shape[1]-1)
    # Set document formatting.
    docFormat = wbook.book.add_format({'font': 'Calibri', 'font_size': 11})
    acctFormat = wbook.book.add_format({'font': 'Calibri', 'font_size': 11, 'num_format': 44})
    commaFormat = wbook.book.add_format({'font': 'Calibri', 'font_size': 11, 'num_format': 3})
    newFormat = wbook.book.add_format({'font': 'Calibri', 'font_size': 11, 'bg_color': 'yellow'})
    movedFormat = wbook.book.add_format({'font': 'Calibri', 'font_size': 11, 'bg_color': '#FF9900'})
    # Format and fit each column.
    i = 0
    # Columns which get shrunk down in reports.
    hideCols = ['Technology', 'Excel Part Link', 'Report Part Nbr Link', 'MFG Part Description',
                'Focus', 'Part Class Name',  'Vendor ID', 'Invoice Detail Nbr', 'Assigned Account Rep',
                'Recipient', 'DKLI Report Date', 'Invoice Date Group', 'Comments', 'Sales Channel']
    coreCols = ['Must Contact', 'End Product', 'How Contacted', 'Information for Digikey']
    for col in sheetData.columns:
        acctCols = ['Unit Price', 'Invoiced Dollars']
        if col in acctCols:
            formatting = acctFormat
        elif col == 'Quantity':
            formatting = commaFormat
        else:
            formatting = docFormat
        maxWidth = max(len(str(val)) for val in sheetData[col].values)
        # Set maximum column width at 50.
        maxWidth = min(maxWidth, 50)
        if col in hideCols:
            maxWidth = 0
        elif col in coreCols:
            maxWidth = 25
        sheet.set_column(i, i, maxWidth+0.8, formatting)
        i += 1
    # Highlight new root customer and moved city rows.
    try:
        for row in sheetData.index:
            ind = str(sheetData.loc[row, 'TAARCOM Comments']).lower().rstrip() == 'individual'
            no_root_cust = sheetData.loc[row, 'Root Customer..'] == ''
            if ind or no_root_cust:
                continue
            root_cust_loc = int(np.where(sheetData.columns == 'Root Customer..')[0])
            city_loc = int(np.where(sheetData.columns == 'City on Acct List')[0])
            if sheetData.loc[row, 'New T-Cust'] == 'Y':
                sheet.write(row + 1, root_cust_loc, sheetData.loc[row, 'Root Customer..'], newFormat)
            elif not sheetData.loc[row, 'City on Acct List']:
                pass
            elif sheetData.loc[row, 'Customer City'] not in sheetData.loc[row, 'City on Acct List'].split(', '):
                sheet.write(row + 1, root_cust_loc, sheetData.loc[row, 'Root Customer..'],  movedFormat)
                sheet.write(row + 1, city_loc, sheetData.loc[row, 'City on Acct List'], movedFormat)
    except KeyError:
        print('Error locating Sales and/or City on Acct List columns.\n'
              'Unable to highlight without these columns.\n---')


def saveError(*excelFiles):
    """Check Excel files and return True if any file is open."""
    for file in excelFiles:
        try:
            open(file, 'r+')
        except FileNotFoundError:
            pass
        except PermissionError:
            return True
    return False


def main(filepath):
    """Looks up the salespeople for a Digikey Local Insight file.

    Arguments:
    filepath -- The filepath to the new Digikey Insight file.
    """
    # Set the directory paths to the server.
    lookup_dir = 'Z:/Commissions Lookup/'
    if not os.path.exists(lookup_dir):
        lookup_dir = os.getcwd()

    # Load the Root Customer Mappings file.
    root_cust_map = load_root_customer_mappings(lookup_dir)

    # Load the Master Account List file.
    acct_list = load_acct_list(lookup_dir)

    # Load the Salesperson Info file.
    salespeople_info = load_salespeople_info(lookup_dir)

    if any(i.empty for i in (root_cust_map, acct_list, salespeople_info)):
        return

    # Check for duplicate cities in the Salespeople Info.
    city_list = [str(i).split(', ') for i in salespeople_info['Territory Cities'] if i != '']
    city_list = [j for i in city_list for j in i]
    duplicates = set(x for x in city_list if city_list.count(x) > 1)
    if duplicates:
        print('The following cities are in the Salespeople Info file multiple times:\n%s'
              % ', '.join(map(str, duplicates)) + '\nPlease remove the extras and try again.'
              '\n*Program Terminated*')
        return

    print('Looking up salespeople...')

    # Strip the root off of the filepath and leave just the filename.
    filename = os.path.basename(filepath)

    # Load the Digikey Insight file.
    insight_file = pd.read_excel(filepath, None)
    insight_file = insight_file[list(insight_file)[0]].fillna('')

    # -------------------------------------------
    # Clean up and match the new Digikey LI file.
    # -------------------------------------------
    # Switch the datetime objects over to strings.
    for col in list(insight_file):
        try:
            insight_file[col] = insight_file[col].dt.strftime('%Y-%m-%d')
        except AttributeError:
            pass

    # Get the column list and input new columns.
    col_names = list(insight_file)
    if 'Sales' not in col_names:
        col_names[4:4] = ['Sales']
    col_group = ('Must Contact', 'End Product', 'How Contacted', 'Information for Digikey')
    if any(i for i in col_group if i not in col_names):
        col_names[6:6] = [i for i in col_group if i not in col_names]
    if 'Invoiced Dollars' not in col_names:
        col_names[19:19] = ['Invoiced Dollars']
    if 'City on Acct List' not in col_names:
        col_names[25:25] = ['City on Acct List']
    col_names.extend(['TAARCOM Comments', 'New T-Cust'])
    # Remove potential duplicated columns.
    col_names = pd.Series(col_names).drop_duplicates().tolist()

    # Calculate the Invoiced Dollars.
    try:
        qty = pd.to_numeric(insight_file['Qty Shipped'], errors='coerce')
        price = pd.to_numeric(insight_file['Unit Price'], errors='coerce')
        insight_file['Invoiced Dollars'] = qty*price
        insight_file['Invoiced Dollars'].fillna('', inplace=True)
    except KeyError:
        print('Error calculating Invoiced Dollars.\n'
              'Please make sure Qty Shipped and Unit Price columns are in the report.\n'
              '(Also check that the top line of the file contains the column names).\n'
              '*Program Terminated*')
        return

    # Remove the 'Send' column, if present.
    try:
        col_names.remove('Send')
    except ValueError:
        pass

    if 'Root Customer..' not in col_names:
        print('Did not find a column named "Root Customer.."\n'
              'Please make sure this column exists and try again.\n'
              'Note: also check that row 1 of the file is the column headers.'
              '\n*Program Terminated*')
        return

    # ----------------------------------------------------------------------
    # Go through each entry in the Insight file and look for a sales match.
    # ----------------------------------------------------------------------
    for row in insight_file.index:
        # Check for individuals and CMs and note them in comments.
        if 'contract' in str(insight_file.loc[row, 'Root Customer Class']).lower().rstrip():
            insight_file.loc[row, 'TAARCOM Comments'] = 'Contract Manufacturer'
        if 'individual' in str(insight_file.loc[row, 'Root Customer Class']).lower().rstrip():
            insight_file.loc[row, 'TAARCOM Comments'] = 'Individual'
            city = str(insight_file.loc[row, 'Customer City']).upper().rstrip()
            # Check for matches to city and assign salesperson.
            for person in salespeople_info.index:
                cities = str(salespeople_info['Territory Cities'][person]).upper().rstrip().split(', ')
                if city in cities:
                    insight_file.loc[row, 'Sales'] = salespeople_info.loc[person, 'Sales Initials']
            # Done, so move to next line in file.
            continue
        cust = str(insight_file.loc[row, 'Root Customer..']).lower().rstrip()
        # Check for customer match in account list.
        acct_list_match = acct_list[acct_list['ProperName'].astype(str).str.lower() == cust]
        if cust and len(acct_list_match) == 1:
            # Check if the city is different from our account list.
            acct_list_city = str(acct_list_match['CITY'].iloc[0]).upper().split(', ')
            if str(insight_file.loc[row, 'Customer City']).upper().rstrip() not in acct_list_city:
                if len(acct_list_city) > 1:
                    acct_list_city = ', '.join(acct_list_city)
                insight_file.loc[row, 'City on Acct List'] = acct_list_city
            # Copy over salesperson.
            insight_file.loc[row, 'Sales'] = acct_list_match['SLS'].iloc[0]
        else:
            # Look for match in root_cust_map file.
            sales_match = root_cust_map['Root Customer'].astype(str).str.lower() == cust
            match = root_cust_map[sales_match]
            if cust and len(match) == 1:
                # Match to salesperson if exactly one match is found.
                insight_file.loc[row, 'Sales'] = match['Salesperson'].iloc[0]
            elif len(match) > 1:
                print('Multiple entries found in rootCustomerMappings for %s!' % cust)
            else:
                # Record that the customer is new.
                insight_file.loc[row, 'New T-Cust'] = 'Y'
                # Look up based on city and fill in.
                city = insight_file.loc[row, 'Customer City'].upper().rstrip()
                for person in salespeople_info.index:
                    cities = salespeople_info['Territory Cities'][person].upper().split(', ')
                    if city in cities:
                        insight_file.loc[row, 'Sales'] = salespeople_info.loc[person, 'Sales Initials']

        # Convert applicable entries to numeric.
        for col in list(insight_file):
            insight_file.loc[row, col] = pd.to_numeric(insight_file.loc[row, col], errors='ignore')

    # Reorder columns and fill NaNs.
    insight_file = insight_file.loc[:, col_names].fillna('')

    # Try saving the files, exit with error if any file is currently open.
    output_dir = 'C:/Users/kerry/Documents/disty data/Digikey/'
    if not os.path.exists(output_dir):
        print('Output directory %s not found!\nUsing current directory for saving files.' % output_dir)
        output_dir = os.getcwd()
    fname1 = os.path.join(output_dir, filename[:-5] + ' With Salespeople.xlsx')
    if saveError(fname1):
        print('---\nOne or more files are currently open in Excel!\n'
              'Please close the files and try again.\n*Program Terminated*')
        return

    # Write the Digikey Insight file, which now contains salespeople.
    writer1 = pd.ExcelWriter(fname1, engine='xlsxwriter', datetime_format='mm/dd/yyyy')
    insight_file.to_excel(writer1, sheet_name='Data', index=False)
    # Format in Excel.
    tableFormat(insight_file, 'Data', writer1)

    # Save the files.
    writer1.save()

    print('---\nSalespeople successfully looked up!\n'
          'New file saved as:\n ' + fname1 + '\n+Program Complete+')
