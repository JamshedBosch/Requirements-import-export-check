import os
import pandas as pd


# Define the check functions
def check_empty_object_id_with_forbidden_cr_status(df):
    """
    Checks if 'Object ID' is empty and 'CR-Status_Bosch_PPx' has forbidden values.
    Returns findings as a list of dictionaries.
    """
    findings = []
    forbidden_status = ['014,', '013,', '100,']
    for index, row in df.iterrows():
        if pd.isna(row['Object ID']) and row['CR-Status_Bosch_PPx'] in forbidden_status:
            findings.append({
                'Row': index + 2,  # Adjust for Excel row (index + 2 to account for header row starting at 1)
                'Issue': "Empty 'Object ID' with forbidden 'CR-Status_Bosch_PPx' value",
                'CR-Status_Bosch_PPx': row['CR-Status_Bosch_PPx']
            })
    return findings


# Function to generate a text report for each file
def generate_report(file_path, report_folder):
    # Load the Excel file
    df = pd.read_excel(file_path)

    # Run checks and collect findings
    findings = []
    findings.extend(check_empty_object_id_with_forbidden_cr_status(df))

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
                f.write(f"Issue: {finding['Issue']}\n")
                f.write(
                    f"CR-Status_Bosch_PPx: {finding['CR-Status_Bosch_PPx']}\n")
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