import os
import zipfile
import pandas as pd
import shutil
import glob
import os.path
import pyreqif.reqif
import pyreqif.rif
import sys
import pyreqif.xlsx


def extract_all_files(source_folder, destination_folder):
    for root, _, files in os.walk(source_folder):
        for file in files:
            if file.endswith('.zip') or file.endswith('.reqifz'):
                # construct the file path
                file_path = os.path.join(root, file)
                extract_zip_recursive(file_path, destination_folder)


def extract_zip_recursive(file_path, destination_folder):
    try:

        # Check if the file is a valid ZIP file
        if not zipfile.is_zipfile(file_path):
            print(f"Skipping invalid zip file: {file_path}")
            return

        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(destination_folder)
            for name in zip_ref.namelist():
                nested_zip_path = os.path.join(destination_folder, name)
                if nested_zip_path.endswith('.zip') or nested_zip_path.endswith('.reqifz'):
                    extract_zip_recursive(nested_zip_path, destination_folder)

    except zipfile.BadZipFile:
        print(f"Error: {file_path} is not a valid zip file.")
    except Exception as e:
        print(f"Unexpected error with file {file_path}: {e}")


def delete_folder(folder_path):
    try:
        shutil.rmtree(folder_path)
        print(f"Folder '{folder_path}' and all its contents have been successfully deleted.")
    except Exception as e:
        print(f"Error deleting '{folder_path}': {str(e)}")


def delete_files_except_extensions(folder_path, allowed_extensions):
    """
    Deletes all files and directories in the given folder path except files with the specified extensions.

    Args:
        folder_path (str): The folder path to clean up.
        allowed_extensions (list): List of allowed file extensions (e.g., ['reqif', 'xml']).
    """
    try:
        # Get all files and directories in the folder
        all_files = glob.glob(os.path.join(folder_path, "*"))

        # Loop through all items in the folder
        for file_path in all_files:
            # Check if it's a file and doesn't match the allowed extensions
            if os.path.isfile(file_path):
                if not any(file_path.endswith(f".{ext}") for ext in
                           allowed_extensions):
                    os.remove(file_path)
                    print(f"Deleted file: {file_path}")
            # Check if it's a directory
            elif os.path.isdir(file_path):
                # Optionally delete directories (uncomment if needed)
                # os.rmdir(file_path)  # Removes empty directories only
                print(f"Skipping directory: {file_path}")

        print(
            f"All files except {', '.join(allowed_extensions)} files in '{folder_path}' have been successfully deleted.")

    except Exception as e:
        print(
            f"Error deleting files except '{allowed_extensions}' in '{folder_path}': {str(e)}")


def get_files_with_extension(folder_path, file_extension):
    """
       Searches for files with the specified extension in the given folder path and its subfolders.

       Args:
           folder_path (str): The root folder path to search in.
           file_extension (str): The file extension to search for (e.g., 'txt', 'xml').

       Returns:
           list: A list of file paths matching the specified extension.
       """
    try:
        files_list = []
        # Traverse the folder and its subfolders
        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.endswith(
                        f".{file_extension}"):  # Match the file extension
                    files_list.append(os.path.join(root, file))

        # Optional: Output the found files
        print(
            f"Found {len(files_list)} '{file_extension}' files in '{folder_path}':")
        for file_path in files_list:
            print(file_path)
        
        return files_list
    
    except Exception as e:
        print(f"Error getting '{file_extension}' files in '{folder_path}': {str(e)}")
        return []


# Beispielaufruf
folder_path = 'path_to_your_folder'
file_extension = 'txt'  # Beispiel: Dateien mit der Endung '.txt' suchen
files_list = get_files_with_extension(folder_path, file_extension)
print("Files list:")
print(files_list)


def main():

    # Ping: Von Kunde --> Bosch
    # source_folder = r"\\bosch.com\dfsrb\DfsDE\DIV\CS\DE_CS$\Prj\Customer\DC\1_all\MBOS2Alpha\PING\PING16"  # Update this with the path to your folder containing the zip files
    # reqif_folder  = r"\\bosch.com\dfsrb\DfsDE\DIV\CS\DE_CS$\Prj\Customer\DC\1_all\MBOS2Alpha\PING\PING16\reqif"  # Update this with the path to store extracted REQIF files
    # excel_folder  = r"\\bosch.com\dfsrb\DfsDE\DIV\CS\DE_CS$\Prj\Customer\DC\1_all\MBOS2Alpha\PING\PING16\excel"  # Update this with the path to save Excel files

    # Audi Ping
    # Path containing the reqIf files (Zip Files)
    source_folder = r"D:\AUDI\LAHs_import_FROM_AUDI\2024-11-06_18-17-44-174_export"
    # Path containing the  extracted REQIF files
    reqif_folder = r"D:\AUDI\Reqif_Extracted"
    # Folder containg the converted excel files
    excel_folder = r"D:\AUDI\Reqif2Excel_Converted"

    # PONG: Von Bosch --> Kunde
    # source_folder = r"\\bosch.com\dfsrb\DfsDE\DIV\CS\DE_CS$\Prj\Customer\DC\1_all\MBOS2Alpha\PONG\PONG9"  # Update this with the path to your folder containing the zip files
    # reqif_folder  = r"\\bosch.com\dfsrb\DfsDE\DIV\CS\DE_CS$\Prj\Customer\DC\1_all\MBOS2Alpha\PONG\PONG9\reqif"  # Update this with the path to store extracted REQIF files
    # excel_folder  = r"\\bosch.com\dfsrb\DfsDE\DIV\CS\DE_CS$\Prj\Customer\DC\1_all\MBOS2Alpha\PONG\PONG9\excel"  # Update this with the path to save Excel files


    # delete folders
    delete_folder(reqif_folder)
    delete_folder(excel_folder)
    
    # Ensure destination folders exist
    os.makedirs(reqif_folder, exist_ok=True)
    os.makedirs(excel_folder, exist_ok=True)

    # Extract all recursively zipped files
    extract_all_files(source_folder, reqif_folder)
    
    # Delete all files except reqifs
    delete_files_except_extensions(reqif_folder, ['reqif', 'xml'])
    
    # convert to excel
    files = get_files_with_extension(reqif_folder, 'reqif')
                
    original_path = os.getcwd()
    os.chdir(excel_folder)        
    for file in files:
        base_filename = (file.split("\\")[-1].replace(".reqif",""))
        reqif_document = pyreqif.reqif.load(file)
        pyreqif.xlsx.dump(reqif_document, base_filename+"_local_conversion.xlsx")
    os.chdir(original_path)


if __name__ == "__main__":
    main()

