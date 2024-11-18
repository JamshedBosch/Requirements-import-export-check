from ReqIF2ExelConverter import ReqIF2ExcelProcessor
from ImportExportChecks import ReqIFProcessor, CheckConfiguration
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os


class ImportExportGui:
    def __init__(self, master):
        self.master = master
        master.title("Import Export Checker")
        master.geometry("600x400")

        # create  the menu bar
        menubar = tk.Menu(master)
        master.config(menu=menubar)

        # Create the File menu
        file_menu = tk.Menu(menubar)
        menubar.add_cascade(label="File", menu=file_menu)
        # file_menu.add_command(label="Open")
        # file_menu.add_command(label="Save")
        # file_menu.add_separator()
        file_menu.add_command(label="Exit", command=master.quit)

        # Create the Help menu
        help_menu = tk.Menu(menubar)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about_dialog)

        # Check Type Selection Frame
        self.check_type_frame = tk.Frame(master)
        self.check_type_frame.pack(side=tk.TOP, fill=tk.X, padx=20, pady=10)
        self.check_type_var = tk.IntVar(value=CheckConfiguration.IMPORT_CHECK)

        # set the label for Select type
        tk.Label(self.check_type_frame, text="Select Check Type:",
                 font=("Helvetica", 12)).grid(row=0, column=0, sticky="w")

        # create Radio buttons for import and export
        tk.Radiobutton(self.check_type_frame, text="Import",
                       variable=self.check_type_var,
                       value=CheckConfiguration.IMPORT_CHECK,
                       font=("Helvetica", 10)).grid(row=0, column=1, padx=10,
                                                    sticky="w")
        tk.Radiobutton(self.check_type_frame, text="Export",
                       variable=self.check_type_var,
                       value=CheckConfiguration.EXPORT_CHECK,
                       font=("Helvetica", 10)).grid(row=0, column=2, padx=10,
                                                    sticky="w")

        # Paths Frame
        self.path_frame = tk.Frame(master)
        self.path_frame.pack(side=tk.TOP, fill=tk.X, padx=20, pady=10)

        self.add_path_entry(self.path_frame, "ReqIF Folder Path:",
                            self.browse_reqif_path, 0)
        self.add_path_entry(self.path_frame, "Unzip Folder Path:",
                            self.browse_unzip_path, 1)
        self.add_path_entry(self.path_frame, "Excel Storage Path:",
                            self.browse_excel_path, 2)

        # Buttons Frame
        self.button_frame = tk.Frame(master)
        self.button_frame.pack(side=tk.TOP, fill=tk.X, padx=20, pady=20)

        self.convert_button = tk.Button(self.button_frame, text="Convert",
                                   command=self.convert_files,
                                   font=("Helvetica", 12), width=15, height=2)
        self.convert_button.pack(side=tk.LEFT, padx=20)

        self.execute_button = tk.Button(self.button_frame, text="Execute Checks",
                                   command=self.execute_checks,
                                   font=("Helvetica", 12), width=15, height=2, stat=tk.DISABLED)
        self.execute_button.pack(side=tk.LEFT, padx=20)

        # Status bar
        self.status_bar = tk.Label(master, text="", bd=1, relief=tk.SUNKEN,
                                   anchor=tk.W, font=("Helvetica", 10))
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def show_about_dialog(self):
        about_text = (
            "Import Export Checker\n\n"
            "Version: 1.0\n"
            "This tool allows users to perform Import and Export checks for ReqIF files.\n\n"
            "Features:\n"
            "- Convert ReqIF files to Excel format\n"
            "- Execute Import/Export checks\n"
            "\nInstructions for use will be added here."
        )
        messagebox.showinfo("About", about_text)

    def add_path_entry(self, parent, label_text, browse_command, row):
        """Helper function to add label, entry, and browse button."""
        label = tk.Label(parent, text=label_text, font=("Helvetica", 12))
        label.grid(row=row, column=0, sticky="w", padx=5, pady=10)

        # Hold the text entered(folder path) in the entry field.
        entry_var = tk.StringVar()
        entry = tk.Entry(parent, textvariable=entry_var, width=40,
                         font=("Helvetica", 10))
        entry.grid(row=row, column=1, padx=5, pady=10)

        browse_button = tk.Button(parent, text="Browse",
                                  command=browse_command,
                                  font=("Helvetica", 10), width=10)
        browse_button.grid(row=row, column=2, padx=5, pady=10)

        # Save reference to the entry variable
        if label_text == "ReqIF Folder Path:":
            self.reqif_path_var = entry_var
        elif label_text == "Unzip Folder Path:":
            self.unzip_path_var = entry_var
        elif label_text == "Excel Storage Path:":
            self.excel_path_var = entry_var

    def browse_reqif_path(self):
        self.reqif_path_var.set(filedialog.askdirectory())

    def browse_unzip_path(self):
        self.unzip_path_var.set(filedialog.askdirectory())

    def browse_excel_path(self):
        self.excel_path_var.set(filedialog.askdirectory())

    def convert_files(self):
        """Execute the conversion logic based on the selected radio button."""
        check_type = self.check_type_var.get()  # Get the selected radio button value
        operation_type = {CheckConfiguration.IMPORT_CHECK: "Import",
                          CheckConfiguration.EXPORT_CHECK: "Export"}

        # Ensure the check type is valid
        if check_type not in operation_type:
            self.update_status_bar(
                "No valid option selected. Please select Import or Export.")
            return

        # Shared logic for both Import and Export
        operation = operation_type[check_type]
        self.update_status_bar(f"Performing {operation} Conversion...")
        print(f"\n{operation} Checks Active")

        # Retrieve folder paths
        reqif_folder = self.reqif_path_var.get()
        unzip_folder = self.unzip_path_var.get()
        excel_folder = self.excel_path_var.get()

        # Display folder paths for debugging
        print(f"Reqif Folder: {reqif_folder}"
              f"\nUnzip folder: {unzip_folder}"
              f"\nExcel storage folder: {excel_folder}\n")

        # Create and process the ReqIF to Excel conversion
        processor = ReqIF2ExcelProcessor(
            source_folder=reqif_folder,
            reqif_folder=unzip_folder,
            excel_folder=excel_folder,
            check_type=check_type
        )
        processor.process()

        # Update the status bar after completion
        self.update_status_bar(
            f"{operation} Conversion completed successfully.")

        # Check if Excel files exist in the specified folder
        excel_path = self.excel_path_var.get()
        if os.path.isdir(excel_path) and any(
                file.endswith(".xlsx") or file.endswith(".xls") for file in
                os.listdir(excel_path)):
            self.execute_button.config(state=tk.NORMAL)
        else:
            self.update_status_bar(
                "No Excel files found in the specified path. Execute Checks disabled.")

    def execute_checks(self):
        check_type = self.check_type_var.get()
        processor = ReqIFProcessor(check_type)
        reports = processor.process_folder()
        self.update_status_bar(
            f"Processed {len(reports)} files. Check reports in {CheckConfiguration.REPORT_FOLDER}")

    def update_status_bar(self, message):
        self.status_bar.config(text=message)


def main():
    root = tk.Tk()
    app = ImportExportGui(root)
    root.mainloop()


if __name__ == "__main__":
    main()

