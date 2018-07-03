import pandas as pd


# Load up the Entries Need Fixing file.
fixList = pd.read_excel('Entries Need Fixing.xlsx', 'Data').fillna('')

# Load up the current Running Commissions file.
runningCom = pd.read_excel('Running Commissions.xlsx', 'Master').fillna('')

# Grab the lines that have been fixed.
dateFixed = fixList['Invoice Date'] != ''
endCustFixed = fixList['Reported End Customer'] != ''
distFixed = fixList['Corrected Distributor'] != ''

fixedEntries = fixList[[x and y and z for x, y, z in zip(dateFixed,
                                                         endCustFixed,
                                                         distFixed)]]

