import pandas as pd
import os


# The main function.
def main(filepath):
    """Appends new Digikey Insight files to the Digikey Insight Master.

    Arguments:
    filepath -- The filepath to the new Digikey Insight file.
    """
    # Load the Digikey Insights Master file.
    if os.path.exists('Digikey Insight Master.xlsx'):
        insMast = pd.read_excel('Digikey Insight Master.xlsx',
                                'Master').fillna('')
        filesProcessed = pd.read_excel('Digikey Insight Master.xlsx',
                                       'Files Processed').fillna('')
    else:
        print('---\n'
              'No Insight Master file found!\n'
              'Please make sure Digikey Insight Master is in the directory.\n'
              '***')
        return

    # Load the Root Customer Mappings file.
    if os.path.exists('rootCustomerMappings.xlsx'):
        rootCustMap = pd.read_excel('rootCustomerMappings.xlsx',
                                    'Sales Lookup').fillna('')
    else:
        print('---\n'
              'No Root Customer Mappings file found!\n'
              'Please make sure rootCustomerMappings.xlsx'
              'is in the directory.\n'
              '***')
        return

    # Strip the root off of the filepath and leave just the filename.
    filename = os.path.basename(filepath)
    if filename in filesProcessed['Filename']:
        # Let us know the file is a duplicte.
        print('---\n'
              'The selected Insight file is already in the Insight Master!\n'
              '***')
        return
    # Load the Insight file.
    insFile = pd.read_excel(filepath, None)
    insFile = insFile[list(insFile)[0]].fillna('')

    # Go through each entry in the Insight file and look for a sales match.
    for row in range(len(insFile)):
        salesMatch = insFile.loc[row, 'Root Customer..'] == rootCustMap['Root Customer']
        match = rootCustMap[salesMatch]
        if len(match) == 1:
            # Match to salesperson if exactly one match is found.
            insFile.loc[row, 'Sales'] = match['Salesperson'].iloc[0]
            
        if 'Contract' in insFile[row, 'Root Customer Class']:
            insFile[row, 'TAARCOM Comments'] = 'Contract Manufacturer'
        if 'Individual' in insFile[row, 'Root Customer Class']:
            insFile[row, 'TAARCOM Comments'] = 'Individual'
            
    # Check to see if column names match.

