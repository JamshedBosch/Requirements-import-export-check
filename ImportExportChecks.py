import os
import pandas as pd
import shutil


class CheckConfiguration:
    """Holds configuration and constants for checks."""
    IMPORT_CHECK = 0
    EXPORT_CHECK = 1

    IMPORT_FOLDERS = {
        IMPORT_CHECK: r"D:\AUDI\Import_Reqif2Excel_Converted",
        EXPORT_CHECK: r"D:\AUDI\Export_Reqif2Excel_Converted"
    }

    REPORT_FOLDER = os.path.join(os.getcwd(), "report")


class DataValidator:
    """Import Checks """

    # Check Nr.1
    @staticmethod
    def check_empty_object_id_with_forbidden_cr_status(df):
        """
        Checks if 'Object ID' is empty and 'CR-Status_Bosch_PPx' has forbidden values.
        Returns findings as a list of dictionaries.
        """
        findings = []
        forbidden_status = ['014,', '013,', '100,']
        for index, row in df.iterrows():
            if pd.isna(row['Object ID']) and row['CR-Status_Bosch_PPx'] in forbidden_status:
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
    def check_cr_status_bosch_ppx_conditions(df):
        """
        Checks if 'CR-Status_Bosch_PPx' is '---', 'CR-ID_Bosch_PPx' is not empty,
        and 'BRS-1Box_Status_Hersteller_Bosch_PPx' is not 'verworfen'.
        Returns findings as a list of dictionaries.
        """
        findings = []
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

    """ Export Checks"""
    # Check Nr.1
    @staticmethod
    def check_cr_id_with_typ_and_brs_1box_status_zulieferer_bosch_ppx(df):
        """
        Checks if 'CR-ID_Bosch_PPx' is not empty and 'Typ' is 'Anforderung',
        then 'BRS-1Box_Status_Zulieferer_Bosch_PPx' must be 'akzeptiert' or 'abgelehnt'.
        Returns findings as a list of dictionaries.
        """
        findings = []
        if 'BRS-1Box_Status_Zulieferer_Bosch_PPx' not in df.columns:
            print("Warning: 'BRS-1Box_Status_Zulieferer_Bosch_PPx' column not found in the file.")
            return findings

        for index, row in df.iterrows():
            if not pd.isna(row['CR-ID_Bosch_PPx']) and \
                    row['Typ'] == "Anforderung,":
                if row['BRS-1Box_Status_Zulieferer_Bosch_PPx'] \
                        not in ["akzeptiert", "abgelehnt"]:
                    findings.append({
                        'Row': index + 2,
                        'Attribute': 'CR-ID_Bosch_PPx, Typ, 1Box_Status_Zulieferer_Bosch_PPx',
                        'Issue': ("'CR-ID_Bosch_PPx' is not empty and 'Typ' is 'Anforderung', "
                                  "but 'BRS-1Box_Status_Zulieferer_Bosch_PPx' is not 'akzeptiert' or 'abgelehnt'"),
                        'Value': (f"CR-ID_Bosch_PPx: {row['CR-ID_Bosch_PPx']}, "
                                  f"Typ: {row['Typ'].rstrip(',')}, BRS-1Box_Status_Zulieferer_Bosch_PPx: {row['BRS-1Box_Status_Zulieferer_Bosch_PPx']}")
                    })
        return findings

    # Check Nr.2
    def check_typ_with_brs_1box_status_zulieferer_bosch_ppx(df):
        """
        Checks if 'Typ' is 'Überschrift' or 'Information', then 'BRS-1Box_Status_Zulieferer_Bosch_PPx' must be 'n/a'.
        Returns findings as a list of dictionaries.
        """
        findings = []
        if 'BRS-1Box_Status_Zulieferer_Bosch_PPx' not in df.columns:
            print("Warning: 'BRS-1Box_Status_Zulieferer_Bosch_PPx' column not found in the file.")
            return findings
        for index, row in df.iterrows():
            if row['Typ'] in ["Überschrift,", "Information,"]:
                if row['BRS-1Box_Status_Zulieferer_Bosch_PPx'] != "n/a":
                    findings.append({
                        'Row': index + 2,
                        # Adjust for Excel row (index + 2 to account for header row)
                        'Attribute': 'Typ, BRS-1Box_Status_Zulieferer_Bosch_PPx',
                        'Issue': ("'Typ' is 'Überschrift' or 'Information', "
                                  "but 'BRS-1Box_Status_Zulieferer_Bosch_PPx' is not 'n/a'"),
                        'Value': f"Typ: {row['Typ'].rstrip(',')}, BRS-1Box_Status_Zulieferer_Bosch_PPx: {row['BRS-1Box_Status_Zulieferer_Bosch_PPx']}"
                    })
        return findings


class ReportGenerator:
    """Generates reports from validation findings."""

    @staticmethod
    def generate_report(file_path, report_folder, findings):
        """Generate a text report for findings."""
        report_file = os.path.join(report_folder,
                                   f"{os.path.basename(file_path).replace('.xlsx', '')}_report.txt")

        with open(report_file, 'w') as f:
            f.write(f"Report for file: {os.path.basename(file_path)}\n")
            f.write("=" * 100 + "\n\n")

            if findings:
                f.write(f"Issues found: {len(findings)}\n")
                f.write("-" * 100 + "\n")
                for finding in findings:
                    f.write(f"Row: {finding['Row']}\n")
                    f.write(f"Attribute: {finding['Attribute']}\n")
                    f.write(f"Issue: {finding['Issue']}\n")
                    f.write(f"Value: {finding['Value']}\n")
                    f.write("-" * 100 + "\n")
            else:
                f.write("No issues found.\n")

        return report_file


class ChecksProcessor:
    """Main processor for REQIF file validation."""

    def __init__(self, check_type, excel_folder):
        self.check_type = check_type
        self.report_folder = CheckConfiguration.REPORT_FOLDER
        self.folder_path = excel_folder

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
        df = pd.read_excel(file_path)

        # Select checks based on type
        # Import check AUDI ==> BOSCH
        if self.check_type == CheckConfiguration.IMPORT_CHECK:
            findings = (
                    DataValidator.check_empty_object_id_with_forbidden_cr_status(
                        df) +
                    DataValidator.check_cr_status_bosch_ppx_conditions(df)
            )
        else:
            # Export check BOSCH ==> AUDI
            findings = (
                DataValidator.check_cr_id_with_typ_and_brs_1box_status_zulieferer_bosch_ppx(df)
                +
                DataValidator.check_typ_with_brs_1box_status_zulieferer_bosch_ppx(df)
            )

        # Generate report
        return ReportGenerator.generate_report(file_path, self.report_folder,
                                               findings)

    def _delete_folder(self, folder_path):
        """Delete a folder and its contents."""
        try:
            shutil.rmtree(folder_path, ignore_errors=True)
        except Exception as e:
            print(f"Error deleting '{folder_path}': {str(e)}")


def main():
    # Set the check type: 0 for Import Check, 1 for Export Check
    check_type = CheckConfiguration.IMPORT_CHECK  # Change to EXPORT_CHECK if needed

    processor = ChecksProcessor(check_type)
    reports = processor.process_folder()

    print(
        f"Processed {len(reports)} files. Reports are stored in {CheckConfiguration.REPORT_FOLDER}")


if __name__ == "__main__":
    main()