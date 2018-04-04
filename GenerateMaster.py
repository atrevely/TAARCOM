import pandas as pd
import glob as glb
import numpy as np


# The main function.
def main():

    # Get the master dataframe ready for the new data.
    # %%
    # Check for an existing master list.
    oldMaster = glb.glob('oldMaster*')

    # Check if there's more than zero or one master list supplied. If there's
    # more than one, print an error and quit.
    if len(oldMaster) > 1:
        print('Too many master lists supplied! Only room for one (or zero).')
        return

    # Check to see if we've supplied an existing master list to append to,
    # otherwise start a new one.
    if oldMaster != []:
        finalData = pd.read_excel(oldMaster)
    else:
        print("I didn't find an existing master list. Starting a new one.")
        # These are our names for the data in the master list.
        finalData = pd.DataFrame(columns=['Invoice Number',
                                          'Invoice Date',
                                          'POS Customer',
                                          'Distributor',
                                          'PPN',
                                          'City',
                                          'State',
                                          'Zip',
                                          'Invoice Amount',
                                          'Commission Rate',
                                          'Commission Paid'])

    # Get the new files ready for processing.
    # They should be in the following folder: ...
    # %%
    # Glob the names of the .xlsx or .xls files that we want to process.
    filenames = glb.glob('*.xls*')

    # Remove the master from the filename list so we don't append it to itself.
    if oldMaster in filenames:
        del filenames[filenames.index(oldMaster)]

    # If we didn't find anything new, let us know and quit.
    if filenames == []:
        print("I didn't find any new files to append!")
        return

    # Read in each new file with Pandas and store them as dictionaries.
    # Each dictionary has a dataframe for each sheet in the file.
    inputData = [pd.read_excel(filename, sheetname=None) for filename in filenames]

    # Create the lookup tables for each column.
    # %%
    # Start with an empty lookup table for each column of data.
    lookupTable = pd.DataFrame(columns=list(finalData))

    # Add the lists of keywords for each data column that we want to be able to
    # find in the new sheets.
    lookupTable.at[0, 'Invoice Number'] = ['Invoice Number',
                                           'Invoice']

    lookupTable.at[0, 'Invoice Date'] = ['Date',
                                         'Billing Date',
                                         'Trans Date',
                                         'Invoice Date',
                                         'POS_ShpDate']

    lookupTable.at[0, 'POS Customer'] = ['Ship To Name',
                                         'Customer Name',
                                         'Cust Name',
                                         'Customer',
                                         'Source Customer Name']

    lookupTable.at[0, 'Distributor'] = ['Sold To Name',
                                        'Distri',
                                        'Disti',
                                        'Distributor']
    
    lookupTable.at[0, 'PPN'] = ['Material Number',
                                'Product',
                                'Item',
                                'PtNo',
                                'Databook Product']
    
    # Go through each file, grab the new data, and put it in the master list.
    # %%
    # Iterate through each file that we're appending to the master list.
    for filename in filenames:
        # Grab the next file from the dictionary.
        newData = inputData[filename]
        # Iterate over each column of data that we want to append.
        for dataName in list(finalData):
            

main()