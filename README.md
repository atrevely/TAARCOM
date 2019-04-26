# TAARCOM
Code under development for TAARCOM, Inc. All code in this repository was written by Alexander J. Trevelyan for sole use by TAARCOM, Inc.

GenerateMasterGUI.py -- Launches the main GUI window that interfaces with GenerateMaster.py, MergeFixedEntries.py, and SalesReportGenerator.py in order to process and append new invoice files to the Commissions Master database, as well as manage monthly commissions data and produce monthly reports for salespeople.

GenerateMaster.py -- Processes new (raw) data files, filling in any missing information from the detected principal, looking up end customers and salespeople, then building the monthly Running Commissions file along with the Entries Need Fixing file for line items that require attention. 

MergeFixedEntries.py -- Takes an Entries Need Fixing File and detects which entries are ready for migration to the associated Running Commissions, then updates the Running Commissions.

SalesReportGenerator.py -- Generates the monthly commission and revenue reports for each salesperson using a finished Running Commissions file along with the Commissions Master data. When finished, migrates the Running Commissions data into Commissions Master.

DigikeyInsightsGUI.py -- Launches a GUI window for use in managing the Digikey Insights Reports.
