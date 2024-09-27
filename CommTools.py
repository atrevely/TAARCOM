import datetime
import os
import logging
import GenerateMasterUtils as Utils
from FileIO import load_run_com, load_lookup_master, save_excel_file

logger = logging.getLogger(__name__)


def extract_lookups(path_to_running_com):
    """Scans a Running Commissions file for new Lookup Master entries and copies them over."""
    running_com, files_processed = load_run_com(file_path=path_to_running_com)
    lookup_master = load_lookup_master()

    # ------------------------------------------------------------------------
    # Go through each line of the finished Running Commissions and use them to
    # update the Lookup Master.
    # ------------------------------------------------------------------------
    # Don't copy over INDIVIDUAL, MISC, or ALLOWANCE.
    no_copy_cols = ['INDIVIDUAL', 'UNKNOWN', 'ALLOWANCE']
    pared_ID = [i for i in running_com.index
                if not any(j in running_com.loc[i, 'T-End Cust'].upper() for j in no_copy_cols)]

    for row in pared_ID:
        # First match reported customer.
        reported_customer = str(running_com.loc[row, 'Reported Customer']).lower()
        POS_customter = lookup_master['Reported Customer'].map(lambda x: str(x).lower())
        customer_matches = reported_customer == POS_customter
        # Now match part number.
        part_number = str(running_com.loc[row, 'Part Number']).lower()
        PPN = lookup_master['Part Number'].map(lambda x: str(x).lower())
        part_number_matches = PPN == part_number
        full_matches = lookup_master[part_number_matches & customer_matches]

        # Figure out if this entry is a duplicate of any existing entry.
        duplicate = False
        for match_ID in full_matches.index:
            match_cols = ['CM Sales', 'Design Sales', 'CM', 'T-Name', 'T-End Cust']
            duplicate = all(full_matches.loc[match_ID, i] == running_com.loc[row, i]
                            for i in match_cols)
            if duplicate:
                break

        # If it's not an exact duplicate, add it to the Lookup Master.
        if not duplicate:
            lookup_cols = ['CM Sales', 'Design Sales', 'CM Split', 'CM', 'T-Name', 'T-End Cust',
                           'Reported Customer', 'Principal', 'Part Number', 'City']
            new_lookup = running_com.loc[row, lookup_cols]
            new_lookup['Date Added'] = datetime.datetime.now().date()
            new_lookup['Last Used'] = datetime.datetime.now().date()
            # Not really a better way to do this it seems.
            lookup_master.loc[lookup_master.index.argmax() + 1] = new_lookup

    # Save the Lookup Master.
    filepath = os.path.join(Utils.DIRECTORIES.get('COMM_LOOKUPS_DIR'), 'Lookup Master - Current.xlsx')
    save_excel_file(filename=filepath, tab_data=lookup_master, tab_names='Lookup')
