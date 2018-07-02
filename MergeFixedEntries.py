import pandas as pd
import time


# Load up the Entries Need Fixing file.
fixList = pd.read_excel('Entries Need Fixing.xlsx', 'Data').fillna('')

# Grab the lines that have been fixed.
fixedEntries = fixList.loc[fixList['T-Name'] != '']
