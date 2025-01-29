import os
import pandas as pd
from HelperFunc import HelperFunctions


class ProjectCheckerSSP:
    """Import Checks """

    # Check Nr.6
    @staticmethod
    def check_object_text_with_status_oem_zu_lieferant_r(df, compare_df,
                                                           file_path, compare_file_path):
        """
        Compares the 'ReqIF.Text' attribute with 'Object Text' attribute from a compare file.
        If 'Object Text' differs from 'ReqIF.Text', ensure 'Status OEM zu Lieferant R' is 'zu bewerten'.
        Handles cases where the identifier is either 'ReqIF.ForeignID' or 'Object ID'.
        Logs findings if the condition is not met.
        """
        findings = []
        # Determine the identifier column dynamically
        identifier_col = 'ReqIF.ForeignID' if 'ReqIF.ForeignID' in df.columns else 'Object ID'
        compare_identifier_col = 'ForeignID' if 'ForeignID' in compare_df.columns else 'Object ID'

        required_columns = ['ReqIF.Text', identifier_col,
                            'Status OEM zu Lieferant R']
        compare_required_columns = ['Object Text', compare_identifier_col]

        # Check for missing columns in both DataFrames
        missing_columns = [col for col in required_columns if
                           col not in df.columns]
        missing_compare_columns = [col for col in compare_required_columns if
                                   col not in compare_df.columns]

        # # Ensure required columns exist in both DataFrames
        # required_columns = ['ReqIF.Text', 'ReqIF.ForeignID',
        #                     'Status OEM zu Lieferant R', 'Object Text',
        #                     'ForeignID']
        # missing_columns = [col for col in required_columns[:3] if
        #                    col not in df.columns]
        # missing_compare_columns = [col for col in required_columns[3:] if
        #                            col not in compare_df.columns]

        if missing_columns:
            check_name = __class__.check_object_text_with_status_hersteller_bosch_ppx.__name__
            print(
                f"Warning: Missing columns in the DataFrame: {missing_columns}, "
                f"in File: {file_path}.\nSkipping check: {check_name}")
            return findings

        if missing_compare_columns:
            check_name = __class__.check_object_text_with_status_hersteller_bosch_ppx.__name__
            print(
                f"Warning: Missing columns in the compare file: {missing_compare_columns}.\nSkipping check: {check_name}")
            return findings

        # Create a dictionary for quick lookup of 'Object Text' from compare file
        compare_dict = compare_df.set_index(compare_identifier_col)[
            'Object Text'].to_dict()

        # Iterate through rows in the main DataFrame
        for index, row in df.iterrows():
            object_id = row[identifier_col]
            object_text = row['ReqIF.Text']
            oem_status = row.get('Status OEM zu Lieferant R', None)
            if pd.isna(oem_status):
                oem_status = "Empty"

            # Skip rows with missing 'Object ID'
            if pd.isna(object_id):
                continue

            # Check if the 'Object ID' exists in the compare file
            if object_id in compare_dict:
                compare_text = compare_dict[object_id]

                # Convert to string and strip whitespace
                object_text_str = str(object_text) if not pd.isna(
                    object_text) else ""
                compare_text_str = str(compare_text) if not pd.isna(
                    compare_text) else ""
                object_text_str = object_text_str.strip()
                compare_text_str = compare_text_str.strip()

                # Skip only if both texts are empty
                if not object_text_str and not compare_text_str:
                    continue

                # Normalize both object_text and compare_text
                normalized_object_text = HelperFunctions.normalize_text(
                    object_text_str)
                normalized_compare_text = HelperFunctions.normalize_text(
                    compare_text_str)
                if normalized_object_text != normalized_compare_text:
                    if oem_status not in ['zu bewerten,']:
                        findings.append({
                            'Row': index + 2,  # Adjust for Excel row numbering
                            'Attribute': 'ReqIF.Text, Status OEM zu Lieferant R',
                            'Issue': (
                                f"'ReqIF.Text' differs from 'Object Text' but 'Status OEM zu Lieferant R' is not 'zu bewerten'."
                            ),
                            'Value': (
                                f"{identifier_col}: {object_id}\n\n"
                                f"---------------\n"
                                f"       Customer File Name: {os.path.basename(file_path)}\n"
                                f"       Customer File Object Text: {object_text_str}\n"
                                f"---------------\n"
                                f"       Bosch File Name: {os.path.basename(compare_file_path)}\n"
                                f"       Bosch File Object Text: {compare_text_str}\n"
                                f"---------------\n"
                                f"       Status OEM zu Lieferant R: {oem_status}"
                            )
                        })

        return findings