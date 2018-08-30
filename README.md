# TAARCOM
Code under development for TAARCOM, Inc.

GenerateMasterGUI.py -- Launches a GUI window that interfaces with GenerateMaster.py in order to process and append new files to the Running Commissions file. New files are selected, then a Principal for the file(s) must be chosen in order to activate processes that are unique to each Principal's file structure. Clicking on the 'Process Files to Running Commissions' button will execute GenerateMaster.py in a new thread, temporarily locking the GUI and updating the user as the file(s) are processed.

GenerateMaster.py -- Takes in raw invoice file(s) and a selected Principal, as well as databases for looking up salespeople (Lookup Master) for each transaction, and a table holding the mappings from the raw file(s) to the Running Commissions database (fieldMappings). Appends raw data to the proper column in the Running Commissions database, while calculating any missing data (depending on the Principal) and checking for errors in commission dollars.
