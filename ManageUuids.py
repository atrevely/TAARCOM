from uuid import uuid4
import os

from FileLoader import load_com_master
from FileSaver import prepare_save_file, save_files


def tag_comm_master(master_location):
    """Tag all lines in Commissions Master that don't have a unique id."""
    master_comm, master_files = load_com_master()
    output_file = os.path.join(master_location, 'Commissions Master_BACKUP.xlsx')
    print('Saving backup')
    writer = prepare_save_file(filename=output_file, tab_data=[master_comm, master_files],
                               tab_names=['Master', 'Files Processed'])
    save_files(writer)
    for i in master_comm.index:
        master_comm.loc[i, 'Unique ID'] = uuid4()

    print('Adding unique IDs.')
    writer2 = prepare_save_file(filename=output_file.replace('_BACKUP', ''),
                                tab_data=[master_comm, master_files],
                                tab_names=['Master', 'Files Processed'])
    save_files(writer2)


if __name__ == '__main__':
    tag_comm_master('Z:\\MK Working Commissions')
