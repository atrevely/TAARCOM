import pandas as pd
import time
import datetime
from dateutil.parser import parse
import calendar
import math
import os.path
from FileLoader import load_entries_need_fixing, load_run_com, load_lookup_master
from RCExcelTools import table_format, save_error, form_date

# Set the directory for the data input/output.
if os.path.exists('Z:\\'):
    out_dir = 'Z:\\MK Working Commissions'
    look_dir = 'Z:\\Commissions Lookup'
else:
    out_dir = os.getcwd()
    look_dir = os.getcwd()


def main(run_com_path):
    """Replaces incomplete entries in Running Commissions with final versions.

    Entries in Running Commissions which need attention are copied to the
    Entries Need Fixing file. This function merges fixed entries in the Need
    Fixing file into the Running Commissions file by overwriting the existing
    (bad) entry with the fixed one, then removing it from the Needs Fixing
    file.

    Additionally, this function maintains the Lookup Master by adding new
    entries when needed, and quarantining old entries that have not been
    used in 2+ years.
    """
    # Load up the necessary files.
    running_com, files_processed = load_run_com(run_com_path)
    com_date = run_com_path[-20:]
    entries_need_fixing = load_entries_need_fixing(out_dir + '\\Entries Need Fixing ' + com_date)
    lookup_master = load_lookup_master(look_dir)
    # Track commission dollars.
    try:
        comm = pd.to_numeric(running_com['Actual Comm Paid'], errors='raise').fillna(0)
        tot_com = sum(comm)
    except ValueError:
        print('Non-numeric entry detected in Actual Comm Paid!\n'
              'Check the Actual Comm Paid column for bad data and try again.\n'
              '*Program Teminated*')
        return

    # ------------------------------
    # Load the Quarantined Lookups.
    # ------------------------------
    if os.path.exists(look_dir + '\\Quarantined Lookups.xlsx'):
        quarantined = pd.read_excel(look_dir + '\\Quarantined Lookups.xlsx').fillna('')
    else:
        print('No Quarantied Lookups file found!\n'
              'Please make sure Quarantined Lookups.xlsx is in the directory.\n'
              '*Program Teminated*')
        return

    # ------------------------------------------
    # Get the data that's ready to be migrated.
    # ------------------------------------------
    # Grab the lines that have an End Customer.
    fixed_end_cust = entries_need_fixing[entries_need_fixing['T-End Cust'] != '']
    # Grab entries where salespeople are filled in.
    cm_sales = fixed_end_cust['CM Sales'].map(lambda x: len(x.strip()) == 2)
    design_sales = fixed_end_cust['Design Sales'].map(lambda x: len(x.strip()) == 2)
    fixed = fixed_end_cust[[x or y for x, y in zip(cm_sales, design_sales)]]
    # Return if there's nothing fixed.
    if fixed.shape[0] == 0:
        print('No new fixed entries detected.\n'
              'Entries need a T-End Cust, Salespeople, and an Invoice Date '
              'in order to be eligible for migration to Running Commissions.\n'
              '*Program Teminated*')
        return

    print('Writing fixed entries...')
    # Go through each entry that's fixed and replace it in Running Commissions.
    for row in fixed.index:
        # Fill in the Sales Commission info.
        sales_com = 0.45 * fixed.loc[row, 'Actual Comm Paid']
        fixed.loc[row, 'Sales Commission'] = sales_com
        if fixed.loc[row, 'CM Sales']:
            # Grab split with default to 20.
            split = fixed.loc[row, 'CM Split'] or 20
        else:
            # No CM Sales, so no split.
            split = 0
        fixed.loc[row, 'CM Sales Comm'] = split * sales_com / 100
        fixed.loc[row, 'Design Sales Comm'] = (100 - split) * sales_com / 100
        # -------------------------------
        # Make sure the date makes sense.
        # -------------------------------
        date_error = False
        date_given = fixed.loc[row, 'Invoice Date']
        # Check if the date is read in as a float/int, and convert to string.
        if isinstance(date_given, (float, int)):
            date_given = str(int(date_given))
        # Check if Pandas read it in as a Timestamp object.
        # If so, turn it back into a string (a bit roundabout, oh well).
        elif isinstance(date_given, (pd.Timestamp,  datetime.datetime, datetime.date)):
            date_given = str(date_given)
        # Try parsing the date.
        try:
            date = parse(date_given).date()
            # Make sure the date actually makes sense.
            cuurent_year = int(time.strftime('%Y'))
            if cuurent_year - date.year not in [0, 1]:
                date_error = True
            else:
                # Cast date format into mm/dd/yyyy.
                fixed.loc[row, 'Invoice Date'] = date
                # Fill in quarter/year/month data.
                fixed.loc[row, 'Year'] = date.year
                fixed.loc[row, 'Month'] = calendar.month_name[date.month][0:3]
                qtr = str(math.ceil(date.month/3))
                fixed.loc[row, 'Quarter Shipped'] = (str(date.year) + 'Q' + qtr)
        except (ValueError, TypeError):
            # The date isn't recognized by the parser.
            date_error = True
        except KeyError:
            print('There is no Invoice Date column in Entries Need Fixing!\n'
                  'Please check to make sure an Invoice Date column exists.\n'
                  'Note: Spelling, whitespace, and capitalization matter.\n'
                  '---')
            date_error = True
        # ---------------------------------------------------------------
        # If no error found in date, finish filling out the fixed entry.
        # ---------------------------------------------------------------
        if not date_error:
            # Check for match in commission dollars.
            try:
                id_match_loc = running_com[running_com['Unique ID'] == fixed.loc[row, 'Unique ID']].index.tolist()
                if len(id_match_loc) == 0:
                    print('WARNING! No match found for unique ID %s.' % fixed.loc[row, 'Unique ID'])
                elif len(id_match_loc) > 1:
                    print('WARNING! Multiple matches found for unique ID %s.' % fixed.loc[row, 'Unique ID'])
                id_match_loc = id_match_loc[0]
            except ValueError:
                print('Error reading Running Com Index!\nMake sure all values are numeric.\n'
                      '*Program Teminated*')
                return
            enf_comm = fixed.loc[row, 'Actual Comm Paid']
            rc_comm = running_com.loc[id_match_loc, 'Actual Comm Paid']
            if rc_comm == enf_comm:
                # Replace the Running Commissions entry with the fixed one.
                running_com.loc[id_match_loc, :] = fixed.loc[row, list(running_com)]
            else:
                print('Mismatch in commission dollars found in Entries Need Fixing on row '
                      + str(row + 2) + '\n Check to make sure lines were not '
                      'deleted from the Running Commissions.\n*Program Terminated*')
                return
            # Delete the fixed entry from the Needs Fixing file.
            entries_need_fixing.drop(row, inplace=True)

    # Make sure all the dates are formatted correctly.
    running_com['Invoice Date'] = running_com['Invoice Date'].map(lambda x: form_date(x))
    entries_need_fixing['Invoice Date'] = entries_need_fixing['Invoice Date'].map(lambda x: form_date(x))
    lookup_master['Last Used'] = lookup_master['Last Used'].map(lambda x: form_date(x))
    lookup_master['Date Added'] = lookup_master['Date Added'].map(lambda x: form_date(x))
    # Go through each column and convert applicable entries to numeric.
    cols = list(running_com)
    # Invoice number sometimes has leading zeros we'd like to keep.
    cols.remove('Invoice Number')
    # The INF gets read in as infinity, so skip the principal column.
    cols.remove('Principal')
    for col in cols:
        running_com[col] = pd.to_numeric(running_com[col], errors='ignore')
        entries_need_fixing[col] = pd.to_numeric(entries_need_fixing[col], errors='ignore')
    # Check to make sure commission dollars still match.
    comm = pd.to_numeric(running_com['Actual Comm Paid'],
                         errors='coerce').fillna(0)
    if sum(comm) != tot_com:
        print('Commission dollars do not match after fixing entries!\n'
              'Make sure Entries Need fixing aligns properly with Running Commissions.\n'
              'This error was potentially caused by adding or removing rows in either file.\n'
              '*Program Terminated*')
        return
    # Re-index the fix list and drop nans in Lookup Master.
    entries_need_fixing.reset_index(drop=True, inplace=True)
    lookup_master.fillna('', inplace=True)
    # Check for entries that are too old and quarantine them.
    twoYearsAgo = datetime.datetime.today() - datetime.timedelta(days=720)
    try:
        lastUsed = lookup_master['Last Used'].map(lambda x: pd.Timestamp(x))
        lastUsed = lastUsed.map(lambda x: x.strftime('%Y%m%d'))
    except (AttributeError, ValueError):
        print('Error reading one or more dates in the Lookup Master!\n'
              'Make sure the Last Used column is all MM/DD/YYYY format.\n---')
        return
    dateCutoff = lastUsed < twoYearsAgo.strftime('%Y%m%d')
    oldEntries = lookup_master[dateCutoff].reset_index(drop=True)
    lookup_master = lookup_master[~dateCutoff].reset_index(drop=True)
    if oldEntries.shape[0] > 0:
        # Record the date we quarantined the entries.
        oldEntries.loc[:, 'Date Quarantined'] = datetime.datetime.now().date()
        # Add deprecated entries to the quarantine.
        quarantined = quarantined.append(oldEntries,  ignore_index=True,
                                         sort=False)
        # Notify us of changes.
        print(str(len(oldEntries)) + ' entries quarantied for being more than 2 years old.\n---')

    # Check if the files we're going to save are open already.
    fname1 = out_dir + '\\Running Commissions ' + com_date
    fname2 = out_dir + '\\Entries Need Fixing ' + com_date
    fname3 = look_dir + '\\Lookup Master - Current.xlsx'
    fname4 = look_dir + '\\Quarantined Lookups.xlsx'
    if save_error(fname1, fname2, fname3, fname4):
        print('---\nOne or more files are currently open in Excel!\n'
              'Please close the files and try again.\n'
              '*Program Teminated*')
        return

    # Write the Running Commissions file.
    writer1 = pd.ExcelWriter(fname1, engine='xlsxwriter', datetime_format='mm/dd/yyyy')
    running_com.to_excel(writer1, sheet_name='Master', index=False)
    files_processed.to_excel(writer1, sheet_name='Files Processed', index=False)
    # Format as table in Excel.
    table_format(running_com, 'Master', writer1)
    table_format(files_processed, 'Files Processed', writer1)

    # Write the Needs Fixing file.
    writer2 = pd.ExcelWriter(fname2, engine='xlsxwriter', datetime_format='mm/dd/yyyy')
    entries_need_fixing.to_excel(writer2, sheet_name='Data', index=False)
    # Format as table in Excel.
    table_format(entries_need_fixing, 'Data', writer2)

    # Write the Lookup Master file.
    writer3 = pd.ExcelWriter(fname3, engine='xlsxwriter', datetime_format='mm/dd/yyyy')
    lookup_master.to_excel(writer3, sheet_name='Lookup', index=False)
    # Format as table in Excel.
    table_format(lookup_master, 'Lookup', writer3)

    # Write the Quarantined Lookups file.
    writer4 = pd.ExcelWriter(fname4, engine='xlsxwriter', datetime_format='mm/dd/yyyy')
    quarantined.to_excel(writer4, sheet_name='Lookup', index=False)
    # Format as table in Excel.
    table_format(quarantined, 'Lookup', writer4)

    # Save the files.
    writer1.save()
    writer2.save()
    writer3.save()
    writer4.save()

    print('Fixed entries migrated successfully!\n'
          '+Program Complete+')
