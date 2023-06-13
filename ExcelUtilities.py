
import os
import pandas as pd
from xlrd import XLRDError


def load_lookup_file(file_name, sheet_name):
    """Loads the specified sheet from the lookup file to a dataframe

    :param: file_name: name of the lookup file
    :param: sheet_name: name of the main sheet we pull data from
    :return: dataframe with sheet data
    """

    # Connect to established lookup directory
    look_dir = ""
    filepath = look_dir + file_name

    if os.path.exists(filepath):
        try:
            sheet_data = pd.read_excel(filepath, sheet_name).fillna("")
        except XLRDError:
            print("..Error reading sheet name for " + file_name + "!\n"
                  "..Please make sure the main tab is named \"" + sheet_name + "\".\n"
                  "*Program Terminated*")
            return
    else:
        print("..No " + file_name + " file found!\n"
              "..Please make sure " + file_name + " is in the directory.\n"
              "*Program Terminated*")
        return

    return sheet_data


