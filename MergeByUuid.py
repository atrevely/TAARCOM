import os
import logging
import GenerateMasterUtils as Utils
from FileIO import load_run_com, load_com_master, save_excel_file

logger = logging.getLogger(__name__)


def main(run_com_path):
    """Merge a file (generally a Running Commissions or Quarterly Commissions) into the Commission Master
    by matching the unique IDs.
    """
    # Load up the file to be merged.
    running_com, files_processed = load_run_com(file_path=run_com_path)
    commission_master, master_files = load_com_master()
    if any([commission_master.empty, running_com.empty]):
        logger.info('*Program Terminated*')
        return

    logger.info('Merging file by UID...')
    for row in running_com.index:
        try:
            id_match_loc = commission_master[commission_master['Unique ID'] == running_com.loc[row, 'Unique ID']]
            id_match_loc = id_match_loc.index.tolist()
            if len(id_match_loc) == 0:
                logger.warning(f'WARNING! No match found for unique ID {running_com.loc[row, 'Unique ID']}.')
            elif len(id_match_loc) > 1:
                logger.warning(f'WARNING! Multiple matches found for unique ID {running_com.loc[row, 'Unique ID']}.')
            else:
                id_match_loc = id_match_loc[0]
                # Replace the target entry with the fixed/updated one.
                commission_master.loc[id_match_loc, :] = running_com.loc[row, list(running_com)]
        except ValueError:
            logger.error('Error reading Running Com Index! Make sure all values are numeric.\n'
                         '*Program Terminated*')
            return

    filename = os.path.join(Utils.DIRECTORIES.get('COMM_WORKING_DIR'), 'Commissions Master.xlsx')
    save_excel_file(filename=filename, tab_data=[commission_master, files_processed],
                    tab_names=['Master', 'Files Processed'])

    logger.info('+Merge Complete+')
