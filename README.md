# TAARCOM
Code under development for TAARCOM, Inc. All code in this repository was written by Alexander J. Trevelyan for sole use by TAARCOM, Inc.

GenerateMasterGUI.py -- Launches a GUI window that interfaces with GenerateMaster.py, MergeFixedEntries.py, and SalesReportGenerator.py in order to process and append new invoice files to the Running Commissions database, as well as manage corrected entries and produce reports for salespeople. New files are selected, then a Principal is detected in order to activate special processes that are unique to each Principal's file structure. Clicking on the 'Process Files to Running Commissions' button will execute GenerateMaster.py in a new thread, temporarily locking the GUI and updating the user (in the text window) as the file(s) are processed. Errors that require attention (such as collisions in column names or missing required fields) are captured and output to the user in the text window.

GenerateMaster.py -- Takes in raw invoice file(s) and a selected Principal, as well as databases for looking up salespeople (Lookup Master) for each transaction, and a table holding the mappings from the raw file(s) to the Running Commissions database (fieldMappings). Appends raw data to the proper column in the Running Commissions database, while calculating any missing data (depending on the Principal) and checking for errors in commission dollars.

MergeFixedEntries.py -- Takes an Entries Need Fixing File and detects which entries are ready for migration to the main database, then updates the database with the new information.

SalesReportGenerator.py -- Takes a database and generates reports for each salesperson using the data that is tagged as being associated with them. Also generates an overall report for the period for use by the manager(s) of the data and/or sales team.

DigikeyInsightsGUI.py -- Launches a GUI window for use in managing the Digikey Insights Reports, which alert the sales team to potential new business opportunities.
