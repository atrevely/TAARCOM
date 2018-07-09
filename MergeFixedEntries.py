import pandas as pd
import time
from dateutil.parser import parse
import calendar
import math


# The main function.
def main():
    # Load up the Entries Need Fixing file.
    fixList = pd.read_excel('Entries Need Fixing.xlsx', 'Data').fillna('')

    # Load up the current Running Commissions file.
    runningCom = pd.read_excel('Running Commissions.xlsx', 'Master').fillna('')
    filesProcessed = pd.read_excel('Running Commissions.xlsx',
                                   'Files Processed').fillna('')

    # Load up the Master Lookup.
    masterLookup = pd.read_excel('Lookup Master 6-27-18.xlsx').fillna('')

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

        # Try parsing the date.
        dateError = 0
        try:
            parse(fixedEntries.loc[entry, 'Invoice Date'])
        except ValueError:
            dateError = 1
        except TypeError:
            # Check if Pandas read it in as a Timestamp object.
            # If so, turn it back into a string.
            if isinstance(fixedEntries.loc[entry, 'Invoice Date'], pd.Timestamp):
                fixedEntries.loc[entry, 'Invoice Date'] = str(fixedEntries.loc[entry, 'Invoice Date'])
            else:
                dateError = 1
        # If no error found in date, finish filling out the fixed entry.
        if not dateError:
            dateParsed = parse(fixedEntries.loc[entry, 'Invoice Date'])
            # Cast date format into mm/dd/yyyy.
            fixedEntries.loc[entry, 'Invoice Date'] = dateParsed.strftime('%m/%d/%Y')
            # Fill in quarter/year/month data.
            fixedEntries.loc[entry, 'Year'] = dateParsed.year
            fixedEntries.loc[entry, 'Month'] = calendar.month_name[dateParsed.month][0:3]
            fixedEntries.loc[entry, 'Quarter'] = str(dateParsed.year) + 'Q' + str(math.ceil(dateParsed.month/3))
            # Delete the fixed entry from the Needs Fixing file.
            fixList.drop(fixList[fixList['Running Com Index'] == fixedEntries.loc[entry, 'Running Com Index']].index, inplace=True)

            # Append entry to Lookup Master, if applicable.
            if not fixedEntries.loc[entry, 'Lookup Master Matches']:
                masterLookup = masterLookup.append(fixedEntries.loc[entry, list(masterLookup)],
                                                   ignore_index=True).fillna('')
                # Record the date that the new entry was added to Lookup Master.
                masterLookup.loc[len(masterLookup)-1, 'Date Added'] =  time.strftime('%m/%d/%Y')

    # DROP DUPLICATES IN LOOKUP MASTER AT THIS POINT?

    # Save the Running Commissions file.
    writer1 = pd.ExcelWriter('Running Commissions '
                             + time.strftime('%Y-%m-%d-%H%M')
                             + '.xlsx', engine='xlsxwriter')
    runningCom.to_excel(writer1, sheet_name='Master', index=False)
    filesProcessed.to_excel(writer1, sheet_name='Files Processed', index=False)

    # Save the Needs Fixing file.
    writer2 = pd.ExcelWriter('Entries Need Fixing.xlsx', engine='xlsxwriter')
    fixList.to_excel(writer2, sheet_name='Data', index=False)

    # Save the Lookup Master
    writer3 = pd.ExcelWriter('Lookup Master ' + time.strftime('%Y-%m-%d-%H%M')
                             + '.xlsx', engine='xlsxwriter')
    masterLookup.to_excel(writer3, sheet_name='Lookup', index=False)

    try:
        writer1.save()
    except IOError:
        print('---')
        print('Running Commissions is open in Excel!')
        print('Please close the file and try again.')
        print('***')
        return
    try:
        writer2.save()
    except IOError:
        print('---')
        print('Lookup Master is open in Excel!')
        print('Please close the file and try again.')
        print('***')
        return
    try:
        writer3.save()
    except IOError:
        print('---')
        print('Entries Need Fixing is open in Excel!')
        print('Please close the file and try again.')
        print('***')
        return

    # If no errors, save the files.
    writer1.save()
    writer2.save()
    writer3.save()
