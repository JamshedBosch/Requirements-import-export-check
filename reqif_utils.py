import os
import shutil
import tempfile
import zipfile


class ReqIFProcessor:
    def __init__(self):
        self.customer_temp_dir = None
        self.own_temp_dir = None

    def extract_reqifz_files(self, customer_reqif_path, own_reqif_path):
        """
        Extracts .reqifz files to temporary directories in the current working directory
        and returns the paths to the extracted .reqif files.
        If the input paths are directories, it looks for .reqifz or .reqif files inside them.
        If the input paths are already .reqif files, they are returned as-is.
        """
        def find_reqif_file(directory):
            """Finds a .reqif or .reqifz file in the given directory."""
            for file_name in os.listdir(directory):
                if file_name.endswith('.reqif') or file_name.endswith('.reqifz'):
                    return os.path.join(directory, file_name)
            raise ValueError(f"No .reqif or .reqifz file found in {directory}")

        def extract_if_zip(file_path, temp_dir):
            """Extracts a .reqifz file to a temporary directory and returns the path to the .reqif file."""
            if file_path.endswith('.reqifz'):
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                # Find the extracted .reqif file
                for file_name in os.listdir(temp_dir):
                    if file_name.endswith('.reqif'):
                        return os.path.join(temp_dir, file_name)
                raise ValueError(f"No .reqif file found in {file_path}")
            elif file_path.endswith('.reqif'):
                return file_path
            else:
                raise ValueError(f"Unsupported file type: {file_path}")

        # Create temporary directories in the current working directory
        current_directory = os.getcwd()
        self.customer_temp_dir = os.path.join(current_directory, "temp_customer_reqif")
        self.own_temp_dir = os.path.join(current_directory, "temp_own_reqif")

        # Ensure the temporary directories do not already exist
        if os.path.exists(self.customer_temp_dir):
            shutil.rmtree(self.customer_temp_dir)
        if os.path.exists(self.own_temp_dir):
            shutil.rmtree(self.own_temp_dir)

        # Create the temporary directories
        os.makedirs(self.customer_temp_dir, exist_ok=True)
        os.makedirs(self.own_temp_dir, exist_ok=True)

        try:
            # Handle customer path
            if os.path.isdir(customer_reqif_path):
                customer_file = find_reqif_file(customer_reqif_path)
                customer_reqif_file = extract_if_zip(customer_file, self.customer_temp_dir)
            else:
                customer_reqif_file = extract_if_zip(customer_reqif_path, self.customer_temp_dir)

            # Handle own path
            if os.path.isdir(own_reqif_path):
                own_file = find_reqif_file(own_reqif_path)
                own_reqif_file = extract_if_zip(own_file, self.own_temp_dir)
            else:
                own_reqif_file = extract_if_zip(own_reqif_path, self.own_temp_dir)

            return customer_reqif_file, own_reqif_file
        except Exception as e:
            # Clean up temporary directories in case of an error
            print(f"Error extracting .reqifz files: {e}")
            self.cleanup_temp_dirs()
            raise e

    def cleanup_temp_dirs(self):
        """Cleans up temporary directories."""
        if self.customer_temp_dir and os.path.exists(self.customer_temp_dir):
            shutil.rmtree(self.customer_temp_dir, ignore_errors=True)
        if self.own_temp_dir and os.path.exists(self.own_temp_dir):
            shutil.rmtree(self.own_temp_dir, ignore_errors=True)