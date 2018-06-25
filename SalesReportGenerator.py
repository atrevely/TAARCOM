import pandas as pd
import time


# Load up the Running Master.
runningMaster = pd.read_excel('Running Master INF FY2017.xlsx',
                              'Master').fillna('')

# Grab all of the salespeople initials.
salespeople = list(set().union(runningMaster['CM Sales'].unique(),
                               runningMaster['Design Sales'].unique()))
del salespeople[salespeople == '']

# Select data that has not been reported yet.
unreportedCommissions = runningMaster[runningMaster['Sales Report Date'] == '']

# Go through each salesperson and pull their data.
for person in salespeople:
    # Find sales entries for the salesperson.
    CM = unreportedCommissions['CM Sales'] == person
    Design = unreportedCommissions['Design Sales'] == person

    # Grab entries that are CM Sales for this salesperson.
    CMSales = unreportedCommissions[[x and not y for x, y in zip(CM, Design)]]
    # Determine share of sales.
    CMOnly = CMSales[CMSales['Design Sales'] == '']
    CMOnly['Sales Percent'] = 100
    CMWithDesign = CMSales[CMSales['Design Sales'] != '']
    CMWithDesign['Sales Percent'] = CMWithDesign['CM Split']

    # Grab entries that are Design Sales only.
    designSales = unreportedCommissions[[not x and y for x, y in zip(CM, Design)]]
    # Determine share of sales.
    designOnly = designSales[designSales['CM Sales'] == '']
    designOnly['Sales Percent'] = 100
    designWithCM = designSales[designSales['CM Sales'] != '']
    designWithCM['Sales Percent'] = 100 - designWithCM['CM Split']

    # Grab CM + Design Sales entries.
    dualSales = unreportedCommissions[[x and y for x, y in zip(CM, Design)]]
    dualSales['Sales Percent'] = 100

    # Set report columns.
    reportCols = ['Salesperson', 'Sales Percent', 'T-End Cust', 'T-Name',
                  'CM', 'Reported Customer', 'Reported End Customer',
                  'Principal', 'Corrected Distributor', 'Invoice Number',
                  'Part Number', 'Quantity', 'Unit Price', 'Invoiced Dollars',
                  'Paid-On Revenue', 'Split Percentage', 'Commission Rate',
                  'Actual Comm Paid', 'Invoice Date', 'On/Offshore', 'City',
                  'Cust Part Number']

    # Start creating report.
    finalReport = pd.DataFrame(columns=reportCols)
    # Columns appended from Running Commissions.
    colAppend = [val for val in reportCols if val in list(dualSales)]
    # Append the data.
    finalReport = finalReport.append([CMOnly[colAppend],
                                      CMWithDesign[colAppend],
                                      designOnly[colAppend],
                                      designWithCM[colAppend],
                                      dualSales[colAppend]],
                                     ignore_index=True)

    # Fill in salesperson initials.
    finalReport['Salesperson'] = person

    # Write report to file.
    writer = pd.ExcelWriter(person + ' Sales Report '
                            + time.strftime('%Y-%m-%d')
                            + '.xlsx', engine='xlsxwriter')
    finalReport.to_excel(writer, sheet_name='Report Data', index=False)
    writer.save()

# Fill in the Sales Report Date.
runningMaster.loc[runningMaster['Sales Report Date'] == '',
                  'Sales Report Date'] = time.strftime('%m/%d/%Y')