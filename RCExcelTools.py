import pandas as pd
from dateutil.parser import parse


def tableFormat(sheetData, sheetName, wbook):
    """Formats the Excel output as a table with correct column formatting."""
    # Nothing to format, so return.
    if sheetData.shape[0] == 0:
        return
    sheet = wbook.sheets[sheetName]
    # Set default document format.
    docFormat = wbook.book.add_format({'font': 'Calibri',
                                       'font_size': 11})
    # Currency format ($XX.XX).
    acctFormat = wbook.book.add_format({'font': 'Calibri',
                                        'font_size': 11,
                                        'num_format': 7})
    # Comma format (XX,XXX).
    commaFormat = wbook.book.add_format({'font': 'Calibri',
                                         'font_size': 11,
                                         'num_format': 3})
    # Percent format, one decimal (XX.X%).
    pctFormat = wbook.book.add_format({'font': 'Calibri',
                                       'font_size': 11,
                                       'num_format': '0.0%'})
    # Date format (YYYY-MM-DD).
    dateFormat = wbook.book.add_format({'font': 'Calibri',
                                        'font_size': 11,
                                        'num_format': 14})
    # Format and fit each column.
    index = 0
    for col in sheetData.columns:
        # Match the correct formatting to each column.
        acctCols = ['Unit Price', 'Paid-On Revenue', 'Actual Comm Paid',
                    'Total NDS', 'Post-Split NDS', 'Cust Revenue YTD',
                    'Ext. Cost', 'Unit Cost', 'Total Commissions',
                    'Sales Commission', 'Invoiced Dollars']
        pctCols = ['Split Percentage', 'Commission Rate',
                   'Gross Rev Reduction', 'Shared Rev Tier Rate',
                   'True Comm %', 'Comm Pct']
        coreCols = ['CM Sales', 'Design Sales', 'T-End Cust', 'T-Name',
                    'CM', 'Invoice Date']
        dateCols = ['Invoice Date', 'Paid Date', 'Sales Report Date',
                    'Date Added']
        if col in acctCols:
            formatting = acctFormat
        elif col in pctCols:
            formatting = pctFormat
        elif col in dateCols:
            formatting = dateFormat
        elif col == 'Quantity':
            formatting = commaFormat
        elif col in ['Invoice Number', 'Part Number']:
            # We're going to do some work in order to keep leading zeros.
            for row in sheetData.index:
                invLen = len(str(sheetData.loc[row, col]))
                # Figure out how many places the number goes to.
                numPadding = '0'*invLen
                invNum = pd.to_numeric(sheetData.loc[row, col],
                                       errors='ignore')
                # The only way to coninually preserve leading zeros is by
                # adding the apostrophe label tag in front.
                if len(numPadding) > len(str(invNum)):
                    invNum = "'" + str(invNum)
                invFormat = wbook.book.add_format({'font': 'Calibri',
                                                   'font_size': 11,
                                                   'num_format': numPadding})
                try:
                    sheet.write_number(row+1, index, invNum, invFormat)
                except TypeError:
                    pass
            # Move to the next column, as we're now done formatting
            # the Invoice/Part Numbers.
            index += 1
            continue
        else:
            formatting = docFormat
        # Set column width and formatting.
        try:
            maxWidth = max(len(str(val)) for val in sheetData[col].values)
        except ValueError:
            maxWidth = 0
        # If column is one that always gets filled in, then keep it expanded.
        if col in coreCols:
            maxWidth = max(maxWidth, len(col), 10)
        # Don't let the columns get too wide.
        maxWidth = min(maxWidth, 50)
        # Extra space for '$'/'%' in accounting/percent format.
        if col in acctCols or col in pctCols:
            maxWidth += 2
        sheet.set_column(index, index, maxWidth+0.8, formatting)
        index += 1
    # Set the autofilter for the sheet.
    sheet.autofilter(0, 0, sheetData.shape[0], sheetData.shape[1]-1)


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


def formDate(inputDate):
    """Attemps to format a string as a date, otherwise ignores it."""
    try:
        outputDate = parse(str(inputDate)).date()
        return outputDate
    except (ValueError, OverflowError):
        return inputDate
