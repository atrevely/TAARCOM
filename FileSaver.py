import pandas as pd
from RCExcelTools import table_format, save_error


def save_excel_file(filename, tab_data, tab_names):
    """Save a file as an Excel spreadsheet."""
    if save_error(filename):
        print('---\nThe following file is currently open in Excel:\n %s' % filename
              + '\nPlease close the file and try again.')
        return None
    if not isinstance(tab_data, list):
        tab_data = [tab_data]
    if not isinstance(tab_names, list):
        tab_names = [tab_names]
    assert len(tab_data) == len(tab_names), 'Mismatch in size of tab data and tab names.'
    # Add each tab to the document.
    with pd.ExcelWriter(filename, datetime_format='mm/dd/yyyy') as writer:
        for data, sheet_name in zip(tab_data, tab_names):
            data.to_excel(writer, sheet_name=sheet_name, index=False)
            table_format(sheet_data=data, sheet_name=sheet_name, workbook=writer)
