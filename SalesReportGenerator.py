import pandas as pd


# Load up the Running Master.
runningMaster = pd.read_excel('Running Master INF FY2017.xlsx',
                              'Master').fillna('')

# Grab all of the salespeople initials.
salespeople = list(set().union(runningMaster['CM Sales'].unique(),
                               runningMaster['Design Sales'].unique()))
del salespeople[salespeople == '']

# SELECT SUBSET OF DATA BASED ON WHETHER IT HAS BEEN REPORTED YET.

# Go through each salesperson and pull their data.
for person in salespeople:
    # Find sales entries for the salesperson.
    CM = runningMaster['CM Sales'] == person
    Design = runningMaster['Design Sales'] == person
    
    # Grab entries that are CM Sales for this salesperson.
    CMSales = runningMaster[[x and not y for x, y in zip(CM, Design)]]
    # Determine share of sales.
    CMOnly = CMSales[CMSales['Design Sales'] == '']
    CMOnly['Sales Percent'] = 100
    CMWithDesign = CMSales[CMSales['Design Sales'] != '']
    CMWithDesign['Sales Percent'] = CMWithDesign['CM Split']

    # Grab entries that are Design Sales only.
    designSales = runningMaster[[not x and y for x, y in zip(CM, Design)]]
    # Determine share of sales.
    designOnly = designSales[designSales['CM Sales'] == '']
    designOnly['Sales Percent'] = 100
    designWithCM = designSales[designSales['CM Sales'] != '']
    designWithCM['Sales Percent'] = 100 - designWithCM['CM Split']

    # Grab CM + Design Sales entries.
    dualSales = runningMaster[[x and y for x, y in zip(CM, Design)]]
    dualSales['Sales Percent'] = 100
    
    # Set report columns.
    reportCols = ['Salesperson', 'Sales Percent', 'T-End Cust', 'T-Name',
                  'CM', 'Reported Customer', 'Reported End Customer',
                  'Principal', 'Corrected Distributor', 'Invoice Number',
                  'Part Number','Quantity', 'Unit Price', 'Invoiced Dollars',
                  'Paid-On Revenue', 'Split Percentage', 'Commission Rate',
                  'Actual Comm Paid', 'Invoice Date', 'On/Offshore', 'City',
                  'Cust Part Number']

    # Start creating report.
    finalReport = pd.DataFrame(columns=reportCols)
    # Columns appended from Running Commissions.
    colAppend = [val for val in reportCols if val in list(dualSales)]
    # Appends dual sales.
    finalReport = finalReport.append(dualSales[colAppend], ignore_index=True)
    
    
    
    # Fill in salesperson initials.
    finalReport.loc[, 'Salesperson'] = person