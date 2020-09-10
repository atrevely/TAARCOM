import pandas as pd
from RCExcelTools import table_format, save_error


def prepare_save_file(filename, tab_data, tab_names):
    """Prepare a file for saving, returning its writer object."""
    writer = pd.ExcelWriter(filename, engine='xlsxwriter', datetime_format='mm/dd/yyyy')
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
    for data, sheet_name in zip(tab_data, tab_names):
        data.to_excel(writer, sheet_name=sheet_name, index=False)
        table_format(filename, sheet_name, writer)
    return writer


def save_files(writer_ojects):
    """Attempt to save files, aborting if errors are encountered."""
    if not all(writer_ojects):
        return False
    if not isinstance(writer_ojects, list):
        writer_ojects = [writer_ojects]
    for writer in writer_ojects:
        writer.save()
    return True
