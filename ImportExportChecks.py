import os
import pandas as pd
import shutil
from ReportGenerator import ReportGenerator


class CheckConfiguration:
    """Holds configuration and constants for checks."""
    IMPORT_CHECK = 0
    EXPORT_CHECK = 1

    PROJECT = {
        "PPE_MLBW": "PPE/MLBW",
        "SSP": "SSP"
    }

    IMPORT_FOLDERS = {
        IMPORT_CHECK: r"D:\AUDI\Import_Reqif2Excel_Converted",
        EXPORT_CHECK: r"D:\AUDI\Export_Reqif2Excel_Converted"
    }

    REPORT_FOLDER = os.path.join(os.getcwd(), "report")


class DataValidator:
    """Import Checks """

    # Check Nr.1
    @staticmethod
    def check_empty_object_id_with_forbidden_cr_status(df, file_path):
        """
        Checks if 'Object ID' is empty and 'CR-Status_Bosch_PPx' has forbidden values.
        Returns findings as a list of dictionaries.
        """
        findings = []
        # Check for required columns
        required_columns = ['Object ID', 'CR-Status_Bosch_PPx']
        missing_columns = [col for col in required_columns if
                           col not in df.columns]
        if missing_columns:
            check_name = __class__.check_empty_object_id_with_forbidden_cr_status.__name__
            print(
                f"Warning: Missing columns in the DataFrame: {missing_columns}, "
                f"in File: {file_path}.\nSkipping check: {check_name}")
            return findings

        forbidden_status = ['014,', '013,', '100,']
        for index, row in df.iterrows():
            if pd.isna(row['Object ID']) and row[
                'CR-Status_Bosch_PPx'] in forbidden_status:
                object_id = "Empty"
                findings.append({
                    'Row': index + 2,
                    # Excel rows start at 1; +2 accounts for header row
                    'Attribute': 'Object ID, CR-Status_Bosch_PPx',
                    'Issue': "Empty 'Object ID' with forbidden 'CR-Status_Bosch_PPx' value",
                    'Value': f"Object ID: {object_id}, CR-Status_Bosch_PPx: {row['CR-Status_Bosch_PPx']}"
                })
        return findings

    # Check Nr.2
    @staticmethod
    def check_cr_status_bosch_ppx_conditions(df, file_path):
        """
        Checks if 'CR-Status_Bosch_PPx' is '---', 'CR-ID_Bosch_PPx' is not empty,
        and 'BRS-1Box_Status_Hersteller_Bosch_PPx' is not 'verworfen'.
        Returns findings as a list of dictionaries.
        """
        findings = []
        # Check for required columns
        required_columns = ['CR-Status_Bosch_PPx', 'CR-ID_Bosch_PPx',
                            'BRS-1Box_Status_Hersteller_Bosch_PPx']
        missing_columns = [col for col in required_columns if
                           col not in df.columns]
        if missing_columns:
            check_name = __class__.check_cr_status_bosch_ppx_conditions.__name__
            print(
                f"Warning: Missing columns in the DataFrame: {missing_columns}, "
                f"in File: {file_path}.\nSkipping check: {check_name}")
            return findings

        for index, row in df.iterrows():
            if (row['CR-Status_Bosch_PPx'] == "---" and
                    not pd.isna(row['CR-ID_Bosch_PPx']) and
                    row[
                        'BRS-1Box_Status_Hersteller_Bosch_PPx'] != "verworfen"):
                findings.append({
                    'Row': index + 2,
                    # Adjust for Excel row (index + 2 to account for header row)
                    'Attribute': 'CR-Status_Bosch_PPx, CR-ID_Bosch_PPx, BRS-1Box_Status_Hersteller_Bosch_PPx',
                    'Issue': (
                        "'CR-Status_Bosch_PPx' is '---' while 'CR-ID_Bosch_PPx' is not empty "
                        "and 'BRS-1Box_Status_Hersteller_Bosch_PPx' is not 'verworfen'"),
                    'Value': (
                        f"CR-Status_Bosch_PPx: {row['CR-Status_Bosch_PPx']}, "
                        f"CR-ID_Bosch_PPx: {row['CR-ID_Bosch_PPx']}, "
                        f"BRS-1Box_Status_Hersteller_Bosch_PPx: {row['BRS-1Box_Status_Hersteller_Bosch_PPx']}")
                })
        return findings

    # Check Nr.3
    @staticmethod
    def check_anlaufkonfiguration_empty(df, file_path):
        """
        Checks if 'Anlaufkonfiguration_01', 'Anlaufkonfiguration_02', 'Anlaufkonfiguration_03'
        are empty where 'Object ID' is not empty.
        Returns findings as a list of dictionaries.
        """
        findings = []
        # Check for required columns
        required_columns = ['Object ID', 'Anlaufkonfiguration_01',
                            'Anlaufkonfiguration_02',
                            'Anlaufkonfiguration_03']
        missing_columns = [col for col in required_columns if
                           col not in df.columns]
        if missing_columns:
            check_name = __class__.check_anlaufkonfiguration_empty.__name__
            print(
                f"Warning: Missing columns in the DataFrame: {missing_columns}, "
                f"in File: {file_path}.\nSkipping check: {check_name}")
            return findings

        # Iterate through rows and check conditions
        for index, row in df.iterrows():
            if not pd.isna(
                    row['Object ID']):  # Check if 'Object ID' is not empty
                empty_columns = [col for col in required_columns[1:] if
                                 pd.isna(row[col])]
                if empty_columns:
                    findings.append({
                        'Row': index + 2,
                        # Adjust for Excel row (index + 2 to account for header row)
                        'Attribute': ', '.join(empty_columns),
                        'Issue': (
                            f"{', '.join(empty_columns)} is empty while 'Object ID' is not empty."),
                        'Value': (f"Object ID: {row['Object ID']}, "
                                  f"Empty Columns: {', '.join(empty_columns)}")
                    })
        return findings

    # Check Nr.4
    @staticmethod
    def check_cr_id_empty_for_brs_hersteller_status(df, file_path):
        """
        Checks if 'CR-ID_Bosch_PPx' is empty for any
        'BRS-1Box_Status_Hersteller_Bosch_PPx' status.
        Returns findings as a list of dictionaries.
        """
        findings = []
        # Check for required columns
        required_columns = ['CR-ID_Bosch_PPx',
                            'BRS-1Box_Status_Hersteller_Bosch_PPx']
        missing_columns = [col for col in required_columns if
                           col not in df.columns]
        if missing_columns:
            check_name = __class__.check_cr_id_empty_for_brs_hersteller_status.__name__
            print(
                f"Warning: Missing columns in the DataFrame: {missing_columns}, "
                f"in File: {file_path}.\nSkipping check: {check_name}")
            return findings

        # Iterate through rows and check conditions
        for index, row in df.iterrows():
            if pd.isna(row[
                           'CR-ID_Bosch_PPx']):  # Check if 'CR-ID_Bosch_PPx' is empty
                findings.append({
                    'Row': index + 2,
                    # Adjust for Excel row (index + 2 to account for header row)
                    'Attribute': 'CR-ID_Bosch_PPx, BRS-1Box_Status_Hersteller_Bosch_PPx',
                    'Issue': ("'CR-ID_Bosch_PPx' is empty while "
                              "'BRS-1Box_Status_Hersteller_Bosch_PPx' has a value."),
                    'Value': (f"CR-ID_Bosch_PPx: {row['CR-ID_Bosch_PPx']}, "
                              f"BRS-1Box_Status_Hersteller_Bosch_PPx: {row['BRS-1Box_Status_Hersteller_Bosch_PPx']}")
                })
        return findings

    # Check Nr.6
    @staticmethod
    def check_object_text_with_status_hersteller_bosch_ppx(df, compare_df,
                                                           file_path, compare_file_path):
        """
        Compares the 'Object Text' attribute based on 'Object ID' with a compare file.
        If 'Object Text' differs, ensure 'BRS-1Box_Status_Hersteller_Bosch_PPx' is 'neu/geändert'.
        Optionally ignores spaces in the 'Object Text' for comparison.
        Logs findings if the condition is not met.
        """
        findings = []
        # Ensure required columns exist in both DataFrames
        required_columns = ['Object ID', 'Object Text',
                            'BRS-1Box_Status_Hersteller_Bosch_PPx']
        missing_columns = [col for col in required_columns if
                           col not in df.columns]
        missing_compare_columns = [col for col in required_columns[:2] if
                                   col not in compare_df.columns]

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
        compare_dict = compare_df.set_index('Object ID')[
            'Object Text'].to_dict()

        # Iterate through rows in the main DataFrame
        for index, row in df.iterrows():
            object_id = row['Object ID']
            object_text = row['Object Text']
            brs_status = row.get('BRS-1Box_Status_Hersteller_Bosch_PPx', None)

            # Skip rows with missing 'Object ID'
            if pd.isna(object_id):
                continue

            # Check if the 'Object ID' exists in the compare file
            if object_id in compare_dict:
                compare_text = compare_dict[object_id]

                # Normalize both object_text and compare_text
                normalized_object_text = HelperFunctions.normalize_text(
                    object_text)
                normalized_compare_text = HelperFunctions.normalize_text(
                    compare_text)
                if normalized_object_text != normalized_compare_text:
                    if brs_status not in ['neu/geändert,']:
                        findings.append({
                            'Row': index + 2,  # Adjust for Excel row numbering
                            'Attribute': 'Object Text, BRS-1Box_Status_Hersteller_Bosch_PPx',
                            'Issue': (
                                f"'Object Text' differs but 'BRS-1Box_Status_Hersteller_Bosch_PPx' is not 'neu/geändert'."
                            ),
                            'Value': (
                                f"Object ID: {object_id}\n\n"
                                f"---------------\n"
                                f"       Customer File Name: {os.path.basename(file_path)}\n"
                                f"       Customer File Object Text: {object_text}\n"
                                f"---------------\n"
                                f"       Bosch File Name: {os.path.basename(compare_file_path)}\n"
                                f"       Bosch File Object Text: {compare_text}\n"
                                f"---------------\n"
                                f"       BRS-1Box_Status_Hersteller_Bosch_PPx: {brs_status}"
                            )
                        })

        return findings

    # Check Nr.7
    @staticmethod
    def check_object_text_with_rb_as_status(df, compare_df, file_path, compare_file_path):
        """
        Compares 'Object Text' in the main file with the compare file based on 'Object ID'.
        If 'Object Text' differs, ensure 'RB_AS_Status' is not 'accepted', 'no_req', or 'canceled_closed'.
        Logs findings if the condition is not met.
        """
        findings = []
        # Ensure required columns exist in both DataFrames
        required_columns = ['Object ID', 'Object Text', 'RB_AS_Status']
        missing_columns = [col for col in required_columns[:2] if
                           col not in df.columns]
        missing_compare_columns = [col for col in required_columns if
                                   col not in compare_df.columns]

        if missing_columns:
            check_name = __class__.check_object_text_with_rb_as_status.__name__
            print(
                f"Warning: Missing columns in the DataFrame: {missing_columns}, "
                f"in File: {file_path}.\nSkipping check: {check_name}")
            return findings

        if missing_compare_columns:
            check_name = __class__.check_object_text_with_rb_as_status.__name__
            print(
                f"Warning: Missing columns in the compare file: {missing_compare_columns}.\nSkipping check: {check_name}")
            return findings

        # Create a dictionary for quick lookup of 'Object Text' from main file(gernerated from reqif)
        compare_dict = df.set_index('Object ID')[
            'Object Text'].to_dict()

        # Iterate through rows in the compare file DataFrame
        for index, row in compare_df.iterrows():
            object_id = row['Object ID']
            # here object_text is from the compare CCB file
            object_text = row['Object Text']
            rb_as_status = row.get('RB_AS_Status', None)
            # Debugging
            if rb_as_status is None:
                print(
                    f"Warning: 'RB_AS_Status' is None for Object ID: {object_id}")

                # Skip rows with missing 'Object ID'
            if pd.isna(object_id):
                continue

            # Check if the 'Object ID' exists in the compare file
            if object_id in compare_dict:
                # here the compare text is from generated reqif file
                compare_text = compare_dict[object_id]
                # Normalize both object_text and compare_text
                normalized_object_text = HelperFunctions.normalize_text(
                    object_text)
                normalized_compare_text = HelperFunctions.normalize_text(
                    compare_text)

                # If 'Object Text' differs, check 'RB_AS_Status'
                if normalized_object_text != normalized_compare_text:
                    print(f"rb_as_status: {rb_as_status}")
                    if rb_as_status in ['accepted', 'no_req',
                                        'canceled_closed']:
                        findings.append({
                            'Row': index + 2,  # Adjust for Excel row numbering
                            'Attribute': 'Object Text, RB_AS_Status',
                            'Issue': (
                                f"'Object Text' differs but 'RB_AS_Status' is one of the prohibited values "
                                f"('accepted', 'no_req', 'canceled_closed')."
                            ),
                            'Value': (
                                  f"Object ID: {object_id}\n"
                                  f"---------------\n"
                                  f"       Bosch File Name: { os.path.basename(compare_file_path)}\n"
                                  f"       Bosch File Object Text: {object_text}\n"
                                  f"---------------\n"
                                  f"       Customer File Name: {os.path.basename(file_path)}\n"
                                  f"       Customer File Object Text: {compare_text}\n"
                                  f"---------------\n"
                                  f"       RB_AS_Status: {rb_as_status}"
                            )
                        })

        return findings

    """ Export Checks"""

    # Check Nr.1
    @staticmethod
    def check_cr_id_with_typ_and_brs_1box_status_zulieferer_bosch_ppx(df,
                                                                      file_path):
        """
        Checks if 'CR-ID_Bosch_PPx' is not empty and 'Typ' is 'Anforderung',
        then 'BRS-1Box_Status_Zulieferer_Bosch_PPx' must be 'akzeptiert' or 'abgelehnt'.
        Returns findings as a list of dictionaries.
        """
        findings = []
        # Check for required columns
        required_columns = ['CR-ID_Bosch_PPx', 'Typ',
                            'BRS-1Box_Status_Zulieferer_Bosch_PPx']
        missing_columns = [col for col in required_columns if
                           col not in df.columns]
        if missing_columns:
            print(
                f"Warning: Missing columns in the DataFrame: {missing_columns},"
                f" in File: {file_path}")
            return findings

        for index, row in df.iterrows():
            if not pd.isna(row['CR-ID_Bosch_PPx']) and \
                    row['Typ'] == "Anforderung,":
                if row['BRS-1Box_Status_Zulieferer_Bosch_PPx'] \
                        not in ["akzeptiert", "abgelehnt"]:
                    findings.append({
                        'Row': index + 2,
                        'Attribute': 'CR-ID_Bosch_PPx, Typ, 1Box_Status_Zulieferer_Bosch_PPx',
                        'Issue': (
                            "'CR-ID_Bosch_PPx' is not empty and 'Typ' is 'Anforderung', "
                            "but 'BRS-1Box_Status_Zulieferer_Bosch_PPx' is not 'akzeptiert' or 'abgelehnt'"),
                        'Value': (
                            f"CR-ID_Bosch_PPx: {row['CR-ID_Bosch_PPx']}, "
                            f"Typ: {row['Typ'].rstrip(',')}, BRS-1Box_Status_Zulieferer_Bosch_PPx: {row['BRS-1Box_Status_Zulieferer_Bosch_PPx']}")
                    })
        return findings

    # Check Nr.2
    def check_typ_with_brs_1box_status_zulieferer_bosch_ppx(df, file_path):
        """
        Checks if 'Typ' is 'Überschrift' or 'Information', then 'BRS-1Box_Status_Zulieferer_Bosch_PPx' must be 'n/a'.
        Returns findings as a list of dictionaries.
        """
        findings = []
        required_columns = ['Typ', 'BRS-1Box_Status_Zulieferer_Bosch_PPx']
        missing_columns = [col for col in required_columns if
                           col not in df.columns]

        if missing_columns:
            print(
                f"Warning: Missing columns in the DataFrame: {missing_columns},"
                f" in File: {file_path}")
            return findings

        for index, row in df.iterrows():
            if row['Typ'] in ["Überschrift,", "Information,"]:
                value = str(
                    row['BRS-1Box_Status_Zulieferer_Bosch_PPx']).lower()
                if value != "n/a":
                    findings.append({
                        'Row': index + 2,
                        'Attribute': 'Typ, BRS-1Box_Status_Zulieferer_Bosch_PPx',
                        'Issue': ("'Typ' is 'Überschrift' or 'Information', "
                                  "but 'BRS-1Box_Status_Zulieferer_Bosch_PPx' is not 'n/a'"),
                        'Value': f"Typ: {row['Typ'].rstrip(',')}, BRS-1Box_Status_Zulieferer_Bosch_PPx: {value}"
                    })
        return findings




class ChecksProcessor:
    """Main processor for Excel file Checks."""

    def __init__(self, project_type, check_type, excel_folder, compare_file=None):
        self.project = project_type
        self.check_type = check_type
        self.report_folder = CheckConfiguration.REPORT_FOLDER
        self.folder_path = excel_folder
        self.compare_file = compare_file
        self.compare_df = None  # Dataframe to hold compare file data

        # if compare_file is provided, read it into a DataFrame
        if self.compare_file:
            try:
                self.compare_df = pd.read_excel(self.compare_file,
                                                keep_default_na=False,
                                                na_values=[''])
                print(
                    f"Compare file '{self.compare_file}' loaded successfully.")
            except Exception as e:
                print(f"Error loading compare file '{self.compare_file}': {e}")
                self.compare_df = None

    def process_folder(self):
        """Process all Excel files in the specified folder."""
        # Delete existing report folder
        self._delete_folder(self.report_folder)
        os.makedirs(self.report_folder, exist_ok=True)

        reports = []
        for file_name in os.listdir(self.folder_path):
            if file_name.endswith('.xlsx'):
                file_path = os.path.join(self.folder_path, file_name)
                report = self._process_file(file_path)
                reports.append(report)

        return reports

    def _process_file(self, file_path):
        """Process a single Excel file."""
        # Read Excel file with special handling of missing values:
        #   - keep_default_na=False: Preserves strings like 'n/a', 'N/A', 'NA' as actual strings instead of converting them to NaN
        #   - na_values=['']: Only treats completely empty cells as missing values (NaN)
        # Read Excel file: preserve 'n/a' as string (keep_default_na=False) and only treat empty cells as NaN (na_values=[''])
        df = pd.read_excel(file_path, keep_default_na=False, na_values=[''])
        findings = []

        # Select Project
        if self.project == CheckConfiguration.PROJECT["PPE_MLBW"]:
            # Select checks based on type
            # Import check AUDI ==> BOSCH
            if self.check_type == CheckConfiguration.IMPORT_CHECK:
                findings = (
                        DataValidator.check_empty_object_id_with_forbidden_cr_status(
                            df, file_path) +
                        DataValidator.check_cr_status_bosch_ppx_conditions(
                            df, file_path) +
                        DataValidator.check_anlaufkonfiguration_empty(
                            df, file_path) +
                        DataValidator.check_cr_id_empty_for_brs_hersteller_status(
                            df, file_path)
                )
                if self.compare_df is not None:

                    findings += DataValidator.check_object_text_with_status_hersteller_bosch_ppx(
                        df, self.compare_df, file_path, self.compare_file)

                    # Execute check check_object_text_with_rb_as_status and create a separate report
                    rb_as_status_findings = DataValidator.check_object_text_with_rb_as_status(
                        df, self.compare_df, file_path, self.compare_file)

                    # if rb_as_status_findings:
                        # Generate a separate report for this check
                    ReportGenerator.generate_report(
                        self.compare_file,
                        self.report_folder,
                        rb_as_status_findings
                    )

            else:
                # Export check BOSCH ==> AUDI
                findings = (
                        DataValidator.check_cr_id_with_typ_and_brs_1box_status_zulieferer_bosch_ppx(
                            df, file_path)
                        +
                        DataValidator.check_typ_with_brs_1box_status_zulieferer_bosch_ppx(
                            df, file_path)
                )
        elif self.project == CheckConfiguration.PROJECT["SSP"]:
            print(f"######## CURRENTLY NO CHECKS DEFINED ####################################")


        # Generate report
        return ReportGenerator.generate_report(file_path, self.report_folder,
                                               findings)

    def _delete_folder(self, folder_path):
        """Delete a folder and its contents."""
        try:
            shutil.rmtree(folder_path, ignore_errors=True)
        except Exception as e:
            print(f"Error deleting '{folder_path}': {str(e)}")


class HelperFunctions:

    @staticmethod
    def normalize_text(text, ignore_spaces_and_semicolons=True):
        """
        Normalize the given text by removing spaces, semicolons, quotes, and tab characters.

        :param text: The text to normalize.
        :param ignore_spaces_and_semicolons: Whether to remove spaces and semicolons.
        :return: The normalized text.
        """
        if not isinstance(text, str):  # Ensure text is a string
            text = str(text)  # Convert it to a string if it's not

        if ignore_spaces_and_semicolons:
            text = text.replace(" ", "")  # Remove spaces
            text = text.replace("\t", "")  # Remove tab characters
            text = text.replace(";", "")  # Remove semicolons
            text = text.replace("'", "")  # Remove single quotes
            text = text.replace('"', "")  # Remove double quotes

        # Normalize all types of whitespace (e.g., newlines, extra spaces)
        text = text.strip()  # Remove leading/trailing whitespace

        return text

def main():
    # Set the check type: 0 for Import Check, 1 for Export Check
    check_type = CheckConfiguration.IMPORT_CHECK  # Change to EXPORT_CHECK if needed
    compare_file = r"D:\AUDI\comparefile\CCB_Tracking_PPE.xlsx"

    processor = ChecksProcessor(check_type,
                                CheckConfiguration.IMPORT_FOLDERS[check_type],
                                compare_file)
    reports = processor.process_folder()

    print(
        f"Processed {len(reports)} files. Reports are stored in {CheckConfiguration.REPORT_FOLDER}")


if __name__ == "__main__":
    main()
