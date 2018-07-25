import pandas as pd
import os


# The main function.
def main(filepath):
    """Docstring.
    
    """
    # Load the Digikey Insights Master file.
    if os.path.exists('distributorLookup.xlsx'):
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

    # Strip the root off of the filepath and leave just the filename.
    filename = os.path.basename(filepath)
    if filename in filesProcessed['Filename']:
        # Let us know the file is a duplicte.
        print('---\n'
              'The Insight file is already in the Insight Master!\n'
              '***')
        return