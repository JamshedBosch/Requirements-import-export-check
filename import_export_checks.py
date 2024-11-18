import os
import pandas as pd
import shutil


# Define the check functions

"""
    Import Checks
"""

# Check Nr.1
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
                'Row': index + 2,  # Excel rows start at 1; +2 accounts for header row
                'Attribute': 'Object ID, CR-Status_Bosch_PPx',
                'Issue': "Empty 'Object ID' with forbidden 'CR-Status_Bosch_PPx' value",
                'Value': f"Object ID: {object_id}, CR-Status_Bosch_PPx: {row['CR-Status_Bosch_PPx']}"
            })
    return findings


# Check Nr.2
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
                row['BRS-1Box_Status_Hersteller_Bosch_PPx'] != "verworfen"):
            findings.append({
                'Row': index + 2,  # Adjust for Excel row (index + 2 to account for header row)
                'Attribute': 'CR-Status_Bosch_PPx, CR-ID_Bosch_PPx, BRS-1Box_Status_Hersteller_Bosch_PPx',
                'Issue': ("'CR-Status_Bosch_PPx' is '---' while 'CR-ID_Bosch_PPx' is not empty "
                          "and 'BRS-1Box_Status_Hersteller_Bosch_PPx' is not 'verworfen'"),
                'Value': (f"CR-Status_Bosch_PPx: {row['CR-Status_Bosch_PPx']}, "
                          f"CR-ID_Bosch_PPx: {row['CR-ID_Bosch_PPx']}, "
                          f"BRS-1Box_Status_Hersteller_Bosch_PPx: {row['BRS-1Box_Status_Hersteller_Bosch_PPx']}")
            })
    return findings


"""
    Export Checks
"""


# Check Nr.1
def check_cr_id_with_typ_and_rb_as_status(df):
    """
    Checks if 'CR-ID_Bosch_PPx' is not empty and 'Typ' is 'Anforderung',
    then 'BRS-1Box_Status_Zulieferer_Bosch_PPx' must be 'akzeptiert' or 'abgelehnt'.
    Returns findings as a list of dictionaries.
    """
    findings = []
    if 'RB_AS_Status' not in df.columns:
        print("Warning: 'RB_AS_Status' column not found in the file.")
        return findings
    for index, row in df.iterrows():
        if not pd.isna(row['BRS-1Box_Status_Zulieferer_Bosch_PPx']) and row['Typ'] == "Anforderung,":
            if row['BRS-1Box_Status_Zulieferer_Bosch_PPx'] not in ["akzeptiert", "abgelehnt"]:
                findings.append({
                    'Row': index + 2,  # Adjust for Excel row (index + 2 to account for header row)
                    'Attribute': 'CR-ID_Bosch_PPx, Typ, RB_AS_Status',
                    'Issue': ("'CR-ID_Bosch_PPx' is not empty and 'Typ' is 'Anforderung', "
                              "but 'RB_AS_Status' is not 'accepted' or 'rejected'"),
                    'Value': (f"CR-ID_Bosch_PPx: {row['CR-ID_Bosch_PPx']}, "
                              f"Typ: {row['Typ'].rstrip(',')}, BRS-1Box_Status_Zulieferer_Bosch_PPx: {row['BRS-1Box_Status_Zulieferer_Bosch_PPx']}")
                })
    return findings


# Check Nr.2
def check_typ_with_rb_as_status_no_req(df):
    """
    Checks if 'Typ' is 'Überschrift' or 'Information', then 'RB_AS_Status' must be 'no_req'.
    Returns findings as a list of dictionaries.
    """
    findings = []
    if 'RB_AS_Status' not in df.columns:
        print("Warning: 'RB_AS_Status' column not found in the file.")
        return findings
    for index, row in df.iterrows():
        if row['Typ'] in ["Überschrift,", "Information,"]:
            if row['RB_AS_Status'] != "no_req":
                findings.append({
                    'Row': index + 2,  # Adjust for Excel row (index + 2 to account for header row)
                    'Attribute': 'Typ, RB_AS_Status',
                    'Issue': ("'Typ' is 'Überschrift' or 'Information', "
                              "but 'RB_AS_Status' is not 'no_req'"),
                    'Value': f"Typ: {row['Typ'].rstrip(',')}, RB_AS_Status: {row['RB_AS_Status']}"
                })
    return findings

# Main function to collect all checks
def perform_checks(df, check_type):
    """
    Performs all checks on the DataFrame and returns findings.
    """
    findings = []

    # Add checks here,
    if check_type == 1:
        # Export Check
        findings.extend(check_cr_id_with_typ_and_rb_as_status(df))
        findings.extend(check_typ_with_rb_as_status_no_req(df))

    elif check_type == 0:
        # import check
        findings.extend(check_empty_object_id_with_forbidden_cr_status(df))
        findings.extend(check_cr_status_bosch_ppx_conditions(df))
    else:
        print("Please define the check type")

    return findings


# Function to generate a text report for each file
def generate_report(file_path, report_folder, check_type):
    # Load the Excel file
    df = pd.read_excel(file_path)

    # Run all checks and collect findings
    findings = perform_checks(df, check_type)

    # Prepare the report file path
    report_file = os.path.join(report_folder,
                               f"{os.path.basename(file_path).replace('.xlsx', '')}_report.txt")

    # Write findings to the text report with formatting
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

    print(f"Report generated: {report_file}")


def delete_folder(folder_path):
    try:
        shutil.rmtree(folder_path)
        print(f"Folder '{folder_path}' and all its contents have been successfully deleted.")
    except Exception as e:
        print(f"Error deleting '{folder_path}': {str(e)}")


# Main function to process files in the specified folder
def process_folder(folder_path, check_type):

    # report_folder = os.path.join(os.getcwd(), 'report')

    # Set the path for the 'report' folder under D:\AUDI\
    report_folder = r"D:\AUDI\report"

    # Delete the 'report' folder if exist
    delete_folder(report_folder)

    # Create 'report' folder in the current directory
    os.makedirs(report_folder, exist_ok=True)

    # Process each .xlsx file in the folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(folder_path, file_name)
            generate_report(file_path, report_folder, check_type)


def main():
    # Set the check type here: 0 for Import Check, 1 for Export Check
    check_type = 0  # Change to 1 for Export Check if needed

    # Set folder_path based on check_type
    if check_type == 0:
        folder_path = r"D:\AUDI\Import_Reqif2Excel_Converted"
    else:
        folder_path = r"D:\AUDI\Export_Reqif2Excel_Converted"

    # Check if the folder path exists
    if not os.path.isdir(folder_path):
        print("Invalid folder path. Please check and try again.")
        return

    # Process the folder and generate reports, specifying the type of check (import or export)
    process_folder(folder_path, check_type)
    print(
        "Processing completed. Reports are stored in the 'report' folder in the current directory.")


if __name__ == "__main__":
    main()