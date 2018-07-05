import pandas as pd
import time


# Load up the Entries Need Fixing file.
fixList = pd.read_excel('Entries Need Fixing.xlsx', 'Data').fillna('')

# Load up the current Running Commissions file.
runningCom = pd.read_excel('Running Commissions.xlsx', 'Master').fillna('')

# Grab the lines that have been fixed.
dateFixed = fixList['Invoice Date'] != ''
endCustFixed = fixList['T-End Cust'] != ''
distFixed = fixList['Corrected Distributor'] != ''

fixedEntries = fixList[[x and y and z for x, y, z in zip(dateFixed,
                                                         endCustFixed,
                                                         distFixed)]]
fixedEntries.reset_index(inplace=True, drop=True)

for entry in range(len(fixedEntries)):
    # Replace the Running Commissions entry with the fixed one.
    runningCom.loc[fixedEntries.loc[entry, 'Running Com Index'], :] = fixedEntries.loc[entry, :]

    # Delete the fixed entry from the Needs Fixing file.
    fixList.drop(fixList[fixList['Running Com Index'] == fixedEntries.loc[entry, 'Running Com Index']].index, inplace=True)

# Save the Running Commissions file.
writer = pd.ExcelWriter('Running Commissions '
                        + time.strftime('%Y-%m-%d-%H%M')
                        + '.xlsx', engine='xlsxwriter')
runningCom.to_excel(writer, sheet_name='Master', index=False)
try:
    writer.save()
except IOError:
    print('---')
    print('Running Commissions is open in Excel!')
    print('Please close the file and try again.')
    print('***')
    return
writer.save()

# Save the Needs Fixing file.
writer = pd.ExcelWriter('Entries Need Fixing.xlsx', engine='xlsxwriter')
fixList.to_excel(writer, sheet_name='Data', index=False)
try:
    writer.save()
except IOError:
    print('---')
    print('Entries Need Fixing is open in Excel!')
    print('Please close the file and try again.')
    print('***')
    return
writer.save()

