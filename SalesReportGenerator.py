import pandas as pd


# Load up the Running Master.
runningMaster = pd.read_excel('Running Master INF FY2017.xlsx','Master').fillna('')

# Grab all of the salespeople initials.
salespeople = list(set().union(runningMaster['CM Sales'].unique(),
                       runningMaster['Design Sales'].unique()))
del salespeople[salespeople == '']

# Go through each salesperson and pull their data.
for person in salespeople:
    # Find sales entries for the salesperson.
    CM = runningMaster['CM Sales'] == person
    Design = runningMaster['Design Sales'] == person
    # Grab entries that are CM Sales only.
    CMSales = runningMaster[[x and not y for x, y in zip(CM, Design)]]
    # Grab entries that are Design Sales only.
    designSales = runningMaster[[not x and y for x, y in zip(CM, Design)]]
    # Grab CM + Design Sales entries.
    dualSales = runningMaster[[x and y for x, y in zip(CM, Design)]]

