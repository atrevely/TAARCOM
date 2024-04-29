import os
import pandas as pd
from FileLoader import load_run_com, load_com_master
from RCExcelTools import tab_save_prep, save_error

# Set the directory for the data input/output.
if os.path.exists('Z:\\'):
    data_dir = 'Z:\\MK Working Commissions'
else:
    data_dir = os.getcwd()


def main(run_com_path):
    """Merge a file (generally a Running Commissions or Quarterly Commissions) into the Commission Master
    by matching the unique IDs.
    """
    # Load up the file to be merged.
    running_com, files_processed = load_run_com(file_path=run_com_path)
    com_mast, master_files = load_com_master()
    if any([com_mast.empty, running_com.empty]):
        print('*Program Terminated*')
        return

    print('Merging file by UID...')
    for row in running_com.index:
        try:
            id_match_loc = com_mast[com_mast['Unique ID'] == running_com.loc[row, 'Unique ID']].index.tolist()
            if len(id_match_loc) == 0:
                print('WARNING! No match found for unique ID %s.' % running_com.loc[row, 'Unique ID'])
            elif len(id_match_loc) > 1:
                print('WARNING! Multiple matches found for unique ID %s.' % running_com.loc[row, 'Unique ID'])
            else:
                id_match_loc = id_match_loc[0]
                # Replace the target entry with the fixed/updated one.
                com_mast.loc[id_match_loc, :] = running_com.loc[row, list(running_com)]
        except ValueError:
            print('Error reading Running Com Index!\nMake sure all values are numeric.\n'
                  '*Program Terminated*')
            return

    filename = os.path.join(data_dir, 'Commissions Master.xlsx')
    writer = pd.ExcelWriter(filename, engine='xlsxwriter', datetime_format='mm/dd/yyyy')
    tab_save_prep(writer=writer, data=com_mast, sheet_name='Master')
    tab_save_prep(writer=writer, data=master_files, sheet_name='Files Processed')
    writer.save()
    print('+ Merge Complete +')
