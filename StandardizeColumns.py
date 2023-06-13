

import ExcelUtilities
import pandas as pd


def main(raw_df, principal):
    """Takes raw file and maps data to standard columns"""

    # -----------------------------
    #  Read in field mappings file
    # -----------------------------

    field_mappings = ExcelUtilities.load_lookup_file(file_name="FieldMappings.xlsx", sheet_name="PrincipalFields")

    # ----------------------
    #  Create standard file
    # ----------------------

    # Field mappings header includes principal, header start location, and data start location before listing the
    #   header columns
    # Filter field mappings columns to header columns to define header for the standard file
    header_start_idx = 6
    standard_columns = field_mappings.columns[header_start_idx:]

    # Create pandas DataFrame for the standard file
    standard_df = pd.DataFrame(columns=standard_columns)

    # ---------------------------------------------------
    #  Map raw file columns to matching standard columns
    # ---------------------------------------------------

    # Find row which contains column names used by this principal
    principal_columns = field_mappings[field_mappings['Principal'] == principal].values[header_start_idx:]



