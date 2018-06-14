import pandas as pd


# Load up the Running Master.
runningMaster = pd.read_excel('Running Master INF FY2017.xlsx','Master').fillna('')

# Grab all of the salespeople initials.
salespeople = list(set().union(runningMaster['CM Sales'].unique(),
                       runningMaster['Design Sales'].unique()))
del salespeople[salespeople == '']

# Go through each salesperson and pull their data.
for person in salespeople:
    # Grab CM sales.
    CM = runningMaster['CM Sales'] == person
    # Grab design sales.
    Design = runningMaster['Design Sales'] == person
    # Find matches to CM and/or Design sales.
    salesData = runningMaster[[x or y for x, y in zip(CM, Design)]]
    