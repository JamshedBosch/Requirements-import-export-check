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

    REPORT_FOLDER = r"D:\AUDI\report"


class DataValidator:
    """Performs specific validation checks on DataFrame."""

    @staticmethod
    def check_empty_object_id_with_forbidden_cr_status(df):
        """Check for empty Object ID with forbidden CR-Status."""
        findings = []
        forbidden_status = ['014,', '013,', '100,']
        for index, row in df.iterrows():
            if pd.isna(row['Object ID']) and row[
                'CR-Status_Bosch_PPx'] in forbidden_status:
                findings.append({
                    'Row': index + 2,
                    'Attribute': 'Object ID, CR-Status_Bosch_PPx',
                    'Issue': "Empty 'Object ID' with forbidden 'CR-Status_Bosch_PPx' value",
                    'Value': f"Object ID: Empty, CR-Status_Bosch_PPx: {row['CR-Status_Bosch_PPx']}"
                })
        return findings

    @staticmethod
    def check_cr_status_bosch_ppx_conditions(df):
        """Check CR-Status conditions."""
        findings = []
        for index, row in df.iterrows():
            if (row['CR-Status_Bosch_PPx'] == "---" and
                    not pd.isna(row['CR-ID_Bosch_PPx']) and
                    row[
                        'BRS-1Box_Status_Hersteller_Bosch_PPx'] != "verworfen"):
                findings.append({
                    'Row': index + 2,
                    'Attribute': 'CR-Status related attributes',
                    'Issue': "'CR-Status_Bosch_PPx' issue with related attributes",
                    'Value': (
                        f"CR-Status_Bosch_PPx: {row['CR-Status_Bosch_PPx']}, "
                        f"CR-ID_Bosch_PPx: {row['CR-ID_Bosch_PPx']}")
                })
        return findings

    @staticmethod
    def check_cr_id_with_typ_and_rb_as_status(df):
        """Check Export-related conditions."""
        findings = []
        if 'RB_AS_Status' not in df.columns:
            return findings

        for index, row in df.iterrows():
            if not pd.isna(row['BRS-1Box_Status_Zulieferer_Bosch_PPx']) and \
                    row['Typ'] == "Anforderung,":
                if row['BRS-1Box_Status_Zulieferer_Bosch_PPx'] not in [
                    "akzeptiert", "abgelehnt"]:
                    findings.append({
                        'Row': index + 2,
                        'Attribute': 'Export-related attributes',
                        'Issue': "Invalid status for requirement type",
                        'Value': (f"Typ: {row['Typ'].rstrip(',')}, "
                                  f"Status: {row['BRS-1Box_Status_Zulieferer_Bosch_PPx']}")
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


class ReqIFProcessor:
    """Main processor for REQIF file validation."""

    def __init__(self, check_type):
        self.check_type = check_type
        self.report_folder = CheckConfiguration.REPORT_FOLDER
        self.folder_path = CheckConfiguration.IMPORT_FOLDERS.get(check_type)

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
        if self.check_type == CheckConfiguration.IMPORT_CHECK:
            findings = (
                    DataValidator.check_empty_object_id_with_forbidden_cr_status(
                        df) +
                    DataValidator.check_cr_status_bosch_ppx_conditions(df)
            )
        else:
            findings = (
                DataValidator.check_cr_id_with_typ_and_rb_as_status(df)
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

    processor = ReqIFProcessor(check_type)
    reports = processor.process_folder()

    print(
        f"Processed {len(reports)} files. Reports are stored in {CheckConfiguration.REPORT_FOLDER}")


if __name__ == "__main__":
    main()