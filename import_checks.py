import os
import pandas as pd


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

# Main function to collect all checks
def perform_checks(df):
    """
    Performs all checks on the DataFrame and returns findings.
    """
    findings = []

    # Add checks here
    findings.extend(check_empty_object_id_with_forbidden_cr_status(df))
    findings.extend(check_cr_status_bosch_ppx_conditions(df))
    # findings.extend(check_mandatory_attributes_filled(df, ['Object ID',
    #                                                        'CR-Status_Bosch_PPx']))

    # You can add more checks as needed by calling different functions here

    return findings


# Function to generate a text report for each file
def generate_report(file_path, report_folder):
    # Load the Excel file
    df = pd.read_excel(file_path)

    # Run all checks and collect findings
    findings = perform_checks(df)

    # Prepare the report file path
    report_file = os.path.join(report_folder,
                               f"{os.path.basename(file_path).replace('.xlsx', '')}_report.txt")

    # Write findings to the text report with formatting
    with open(report_file, 'w') as f:
        f.write(f"Report for file: {os.path.basename(file_path)}\n")
        f.write("=" * 50 + "\n\n")

        if findings:
            f.write("Issues found:\n")
            f.write("-" * 50 + "\n")
            for finding in findings:
                f.write(f"Row: {finding['Row']}\n")
                f.write(f"Attribute: {finding['Attribute']}\n")
                f.write(f"Issue: {finding['Issue']}\n")
                f.write(f"Value: {finding['Value']}\n")
                f.write("-" * 50 + "\n")
        else:
            f.write("No issues found.\n")

    print(f"Text report generated: {report_file}")


# Main function to process files in the specified folder
def process_folder(folder_path):
    # Create 'report' folder in the current directory
    report_folder = os.path.join(os.getcwd(), 'report')
    os.makedirs(report_folder, exist_ok=True)

    # Process each .xlsx file in the folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(folder_path, file_name)
            generate_report(file_path, report_folder)


def main():
    # Get the folder path from the user
    # folder_path = input(
    #     "Enter the path to the folder containing .xlsx files: ").strip()

    folder_path = r"D:\AUDI\Reqif2Excel_Converted"

    # Check if the folder exists
    if not os.path.isdir(folder_path):
        print("Invalid folder path. Please check and try again.")
        return

    # Process the folder and generate reports
    process_folder(folder_path)
    print(
        "Processing completed. Reports are stored in the 'report' folder in the current directory.")


if __name__ == "__main__":
    main()