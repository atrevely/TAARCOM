import pandas as pd
import os
import sys
import re
import shutil
from dateutil.parser import parse
import win32com.client
import pythoncom

# Set the numerical columns.
num_cols = ['Quantity', 'Ext. Cost', 'Invoiced Dollars', 'Paid-On Revenue', 'Actual Comm Paid',
            'Unit Cost', 'Unit Price', 'CM Split', 'Year', 'Sales Commission',
            'Split Percentage', 'Commission Rate', 'Gross Rev Reduction', 'Shared Rev Tier Rate']


class PivotTables:
    """A class that builds pivot tables in Excel."""

    def __init__(self):
        """Start up the Excel instance."""
        pythoncom.CoInitialize()
        try:
            self.excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
        except AttributeError:
            # Need to shut down the old module, or something. Windows sucks.
            # Also, need to grab modules as list to prevent dictionary iteration error.
            module_list = [m.__name__ for m in list(sys.modules.values())]
            for module in module_list:
                if re.match(r'win32com\.gen_py\..+', module):
                    del sys.modules[module]
            shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
            # Now try again.
            self.excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
        self.win32c = win32com.client.constants

    def create_pivot_table(self, excel_file, data_sheet_name, pivot_sheet_name,
                           row_fields, col_field, data_field, page_field=None):
        """Creates a pivot table in the provided file using the specified fields."""
        # Create the workbook and add the report sheet as the rightmost tab.
        wb = self.excel.Workbooks.Open(excel_file)
        if wb is None:
            print('Error loading file :' + os.path.basename(excel_file)
                  + '/nMake sure the file is closed and try again.')
            return
        wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        pivot_sheet = wb.Worksheets(wb.Sheets.Count)
        pivot_sheet.Name = pivot_sheet_name
        data_sheet = wb.Worksheets(data_sheet_name)
        # Grab the report data by selecting the current region.
        data_range = data_sheet.Range('A1').CurrentRegion
        pivot_range = pivot_sheet.Range('A1')
        if not isinstance(row_fields, list):
            row_fields = [row_fields]
        # Create the pivot table and deploy it on the sheet.
        try:
            piv_cache = wb.PivotCaches().Create(SourceType=self.win32c.xlDatabase,
                                                SourceData=data_range,
                                                Version=self.win32c.xlPivotTableVersion14)
            piv_table = piv_cache.CreatePivotTable(TableDestination=pivot_range,
                                                   TableName=data_sheet_name,
                                                   DefaultVersion=self.win32c.xlPivotTableVersion14)
            # Drop the row fields into the pivot table.
            for index, row in enumerate(row_fields):
                piv_table.PivotFields(row).Orientation = self.win32c.xlRowField
                piv_table.PivotFields(row).Position = index + 1
            # Add the column field.
            piv_table.PivotFields(col_field).Orientation = self.win32c.xlColumnField
            if page_field:
                piv_table.PivotFields(page_field).Orientation = self.win32c.xlPageField
            # Add the data field.
            piv_data_field = piv_table.AddDataField(piv_table.PivotFields(data_field),
                                                    'Sum of ' + data_field, self.win32c.xlSum)
            piv_data_field.NumberFormat = '$#,##0'
        except Exception:
            print('Pivot table could not be created in file: ' + excel_file)
        wb.Close(SaveChanges=1)


def table_format(sheet_data, sheet_name, workbook):
    """Formats the Excel output as a table with correct column formatting."""
    # Nothing to format, so return.
    if sheet_data.shape[0] == 0:
        return
    sheet = workbook.sheets[sheet_name]
    sheet.freeze_panes(1, 0)
    # Set default document format.
    doc_format = workbook.book.add_format({'font': 'Calibri', 'font_size': 11})
    # Currency format ($XX.XX).
    acct_format = workbook.book.add_format({'font': 'Calibri', 'font_size': 11, 'num_format': 7})
    # Comma format (XX,XXX).
    comma_format = workbook.book.add_format({'font': 'Calibri', 'font_size': 11, 'num_format': 3})
    # Percent format, one decimal (XX.X%).
    pct_format = workbook.book.add_format({'font': 'Calibri', 'font_size': 11, 'num_format': '0.0%'})
    # Date format (YYYY-MM-DD).
    date_format = workbook.book.add_format({'font': 'Calibri', 'font_size': 11, 'num_format': 14})
    # Format and fit each column.
    for index, col in enumerate(sheet_data.columns):
        # Match the correct formatting to each column.
        acct_cols = ['Unit Price', 'Paid-On Revenue', 'Actual Comm Paid', 'Total NDS', 'Post-Split NDS',
                     'Cust Revenue YTD', 'Ext. Cost', 'Unit Cost', 'Total Commissions',
                     'Sales Commission', 'Invoiced Dollars', 'CM Sales Comm', 'Design Sales Comm']
        pct_cols = ['Split Percentage', 'Commission Rate', 'Gross Rev Reduction', 'Shared Rev Tier Rate',
                    'True Comm %', 'Comm Pct']
        core_cols = ['CM Sales', 'Design Sales', 'T-End Cust', 'T-Name', 'CM', 'Invoice Date']
        date_cols = ['Invoice Date', 'Paid Date', 'Sales Report Date', 'Date Added']
        hide_cols = ['Quarter Shipped', 'Month', 'Year', 'Reported Distributor', 'PO Number', 'Sales Order #',
                     'Unit Cost', 'Unit Price', 'Comm Source', 'On/Offshore', 'INF Comm Type', 'PL',
                     'Division', 'Gross Rev Reduction', 'Project', 'Product Category', 'Shared Rev Tier Rate',
                     'Cust Revenue YTD', 'Total NDS', 'Post-Split NDS', 'Cust Part Number', 'End Market',
                     'Q Number', 'CM Split']
        if col in acct_cols:
            formatting = acct_format
        elif col in pct_cols:
            formatting = pct_format
        elif col in date_cols:
            formatting = date_format
        elif col == 'Quantity':
            formatting = comma_format
        elif col in ['Invoice Number', 'Part Number']:
            # We're going to do some work in order to keep leading zeros.
            for row in sheet_data.index:
                inv_len = len(str(sheet_data.loc[row, col]))
                # Figure out how many places the number goes to.
                num_padding = '0'*inv_len
                inv_num = pd.to_numeric(sheet_data.loc[row, col], errors='ignore')
                # The only way to continually preserve leading zeros is by
                # adding the apostrophe label tag in front.
                if len(num_padding) > len(str(inv_num)):
                    inv_num = "'" + str(inv_num)
                inv_format = workbook.book.add_format({'font': 'Calibri', 'font_size': 11,
                                                       'num_format': num_padding})
                try:
                    sheet.write_number(row+1, index, inv_num, inv_format)
                except TypeError:
                    pass
            # Move to the next column.
            continue
        else:
            formatting = doc_format
        # Set column width and formatting.
        try:
            max_width = max(len(str(val)) for val in sheet_data[col].values)
        except ValueError:
            max_width = 0
        # Expand/collapse important columns for RC/ENF.
        if col in hide_cols and sheet_name != 'Master Data':
            max_width = 0
        elif col in core_cols:
            max_width = max(max_width, len(col), 10)
        # Don't let the columns get too wide.
        max_width = min(max_width, 50)
        # Extra space for '$'/'%' in accounting/percent format.
        if (col in acct_cols or col in pct_cols) and col not in hide_cols:
            max_width += 2
        sheet.set_column(index, index, max_width+0.8, formatting)
    # Set the auto-filter for the sheet.
    sheet.autofilter(0, 0, sheet_data.shape[0], sheet_data.shape[1]-1)


def save_error(*excel_files):
    """Check Excel files and return True if any file is open."""
    for file in excel_files:
        try:
            open(file, 'r+')
        except FileNotFoundError:
            pass
        except PermissionError:
            return True
    return False


def form_date(input_date):
    """Attempts to format a string as a date, otherwise ignores it."""
    try:
        output_date = parse(str(input_date)).date()
        return output_date
    except (ValueError, OverflowError):
        return input_date


def tab_save_prep(writer, data, sheet_name):
    """Prepares a file for being saved."""
    # Make sure desired columns are numeric, and replace zeros in non-commission columns with blanks.
    for col in num_cols:
        try:
            if col not in ['Actual Comm Paid', 'Sales Commission']:
                fill = ''
                data[col].replace(0, '', inplace=True)
            else:
                fill = 0
            data[col] = pd.to_numeric(data[col], errors='coerce').fillna(fill)
        except KeyError:
            pass
    # Convert individual numbers to numeric in rest of columns.
    mixed_cols = [col for col in list(data) if col not in num_cols]
    skip_cols = ['Invoice Number', 'Part Number', 'Principal']
    mixed_cols = [i for i in mixed_cols if i not in skip_cols]
    for col in mixed_cols:
        data[col] = pd.to_numeric(data[col], errors='ignore')
    date_cols = ['Invoice Date', 'Date Added', 'Paid Date']
    # Format the dates correctly where possible.
    for col in date_cols:
        try:
            data[col] = data[col].map(lambda x: form_date(x))
        except KeyError:
            pass
    data.to_excel(writer, sheet_name=sheet_name, index=False)
    # Do the Excel formatting.
    table_format(sheet_data=data, sheet_name=sheet_name, workbook=writer)
