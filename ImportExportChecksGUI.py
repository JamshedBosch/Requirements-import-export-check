from ReqIF2ExelConverter import ReqIF2ExcelProcessor
from ImportExportChecks import ChecksProcessor, CheckConfiguration
from tkinter import filedialog, ttk, messagebox, PhotoImage
import os
import tkinter as tk
from tkinter import ttk

class ImportExportGui:
    def __init__(self, master):
        self.master = master
        master.title("Import Export Checker")
        master.geometry("600x400")
        master.resizable(False, False)

        # Apply custom styles
        style = ttk.Style()
        style.theme_use('default')

        # Set custom colors
        style.configure('TLabel', background='#f0f0f0', foreground='#333333')
        style.configure('TButton', background='#007bff', foreground='#ffffff', padding=8,
                        font=("Helvetica", 10), relief=tk.FLAT)
        style.map('TButton', background=[('active', '#0069d9')])
        style.configure('TRadiobutton', background='#f0f0f0', foreground='#333333')
        style.configure('TEntry', fieldbackground='#ffffff', foreground='#333333')
        style.configure('TFrame', background='#f0f0f0')

        # Configure Compare checkbox style to match window
        style.configure('TCheckbutton',
                        background='#f0f0f0',  # Match window background
                        foreground='#333333',  # Match text color
                        font=("Helvetica", 10))  # Match font
        style.map('TCheckbutton',
                  background=[('active', '#f0f0f0')],
                  # Keep background consistent when active
                  foreground=[('active',
                               '#333333')])  # Keep text color consistent when active

        # create the menu bar
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
        self.check_type_frame = ttk.Frame(master)
        self.check_type_frame.pack(side=tk.TOP, fill=tk.X, padx=20, pady=10)
        self.check_type_var = tk.IntVar(value=CheckConfiguration.IMPORT_CHECK)

        # Set the label for Select type
        ttk.Label(self.check_type_frame, text="Select Check Type:",
                  font=("Helvetica", 12)).grid(row=0, column=0, sticky="w")

        # Create Radio buttons for import and export
        ttk.Radiobutton(self.check_type_frame, text="Import",
                        variable=self.check_type_var,
                        value=CheckConfiguration.IMPORT_CHECK,
                        style='TRadiobutton').grid(row=0, column=1, padx=10,
                                                  sticky="w")
        ttk.Radiobutton(self.check_type_frame, text="Export",
                        variable=self.check_type_var,
                        value=CheckConfiguration.EXPORT_CHECK,
                        style='TRadiobutton').grid(row=0, column=2, padx=10,
                                                  sticky="w")

        # Add "Select compare file" label and checkbox
        ttk.Label(self.check_type_frame, text="Select compare file:",
                  font=("Helvetica", 12)).grid(row=1, column=0, sticky="w",
                                               pady=(15, 5))

        # Checkbox variable and button
        self.show_path_var = tk.BooleanVar()
        self.checkbox = ttk.Checkbutton(
            self.check_type_frame,
            text="",
            variable=self.show_path_var,
            command=self.toggle_reference_path,
            style='TCheckbutton'
        )
        self.checkbox.grid(row=1, column=1, sticky="w", pady=(15, 5), padx=(25, 0))

        # Paths Frame
        self.path_frame = ttk.Frame(master)
        self.path_frame.pack(side=tk.TOP, fill=tk.X, padx=20, pady=10)

        # Add path entries
        self.add_path_entry(self.path_frame, "ReqIF folder:", self.browse_reqif_path, 0)
        self.add_path_entry(self.path_frame, "Extract folder:", self.browse_unzip_path, 1)
        self.add_path_entry(self.path_frame, "Excel folder:", self.browse_excel_path, 2)

        # Reference path entry and browse button (initially hidden)
        self.ref_path_var = tk.StringVar()
        self.ref_path_label = ttk.Label(self.path_frame, text="Compare file:",
                                        font=("Helvetica", 12))
        self.ref_path_entry = ttk.Entry(
            self.path_frame,
            textvariable=self.ref_path_var,
            width=40,
            font=("Helvetica", 10)
        )
        self.ref_browse_button = tk.Button(
            self.path_frame,
            text="Browse",
            command=self.browse_reference_path,
            font=("Helvetica", 10),
            width=10  # Match width with other browse buttons
        )

        # Buttons Frame
        self.button_frame = ttk.Frame(master)
        self.button_frame.pack(side=tk.TOP, fill=tk.X, padx=20, pady=20)

        self.convert_button = ttk.Button(self.button_frame, text="Convert",
                                        command=self.convert_files,
                                        style='TButton')
        self.convert_button.pack(side=tk.LEFT, padx=20)

        self.execute_button = ttk.Button(self.button_frame, text="Execute Checks",
                                        command=self.execute_checks,
                                        style='TButton', stat=tk.DISABLED)
        self.execute_button.pack(side=tk.LEFT, padx=20)

        # Status bar
        self.status_bar = ttk.Label(master, text="", relief=tk.SUNKEN,
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
        if label_text == "ReqIF folder:":
            self.reqif_path_var = entry_var
        elif label_text == "Extract folder:":
            self.unzip_path_var = entry_var
        elif label_text == "Excel folder:":
            self.excel_path_var = entry_var

    def browse_reqif_path(self):
        self.reqif_path_var.set(filedialog.askdirectory())

    def browse_unzip_path(self):
        self.unzip_path_var.set(filedialog.askdirectory())

    def browse_excel_path(self):
        self.excel_path_var.set(filedialog.askdirectory())

    def browse_reference_path(self):
        """Open file dialog to select reference Excel file"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if file_path:
            self.ref_path_var.set(file_path)

    def operation_type(self):
        """Execute the conversion logic based on the selected radio button."""
        check_type = self.check_type_var.get()  # Get the selected radio button value
        operation_type = {CheckConfiguration.IMPORT_CHECK: "Import",
                          CheckConfiguration.EXPORT_CHECK: "Export"}

        # Ensure the check type is valid
        if check_type not in operation_type:
            self.update_status_bar(
                "No valid option selected. Please select Import or Export.")
            return
        else:
            return operation_type[check_type]

    def convert_files(self):
        check_type = self.operation_type()
        self.update_status_bar(f"Performing {check_type} Conversion...")
        print(f"\n{check_type} Checks Active")

        # Retrieve folder paths
        reqif_folder = self.reqif_path_var.get()
        unzip_folder = self.unzip_path_var.get()
        excel_folder = self.excel_path_var.get()

        # Display folder paths for debugging
        print(f"Reqif Folder: {reqif_folder}"
              f"\nExtract folder: {unzip_folder}"
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
            f"{check_type} Conversion completed successfully.")

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
        self.update_status_bar(
            f"{self.operation_type()} Checks processing started...")
        self.master.update()  # Updates the Tkinter GUI before continuing
        reference_file = self.ref_path_var.get() if self.show_path_var.get() else None
        print(f"Path of the refernce file is:  '{reference_file}'")

        processor = ChecksProcessor(check_type, self.excel_path_var.get())
        reports = processor.process_folder()
        self.update_status_bar(
            f"Processed {len(reports)} files. Check reports in {CheckConfiguration.REPORT_FOLDER}")

    def toggle_reference_path(self):
        """Show or hide the reference path entry and browse button based on checkbox state"""
        if self.show_path_var.get():
            self.ref_path_label.grid(row=3, column=0, sticky="w", pady=10)
            self.ref_path_entry.grid(row=3, column=1, padx=5, pady=10)
            self.ref_browse_button.grid(row=3, column=2, padx=5, pady=10)
        else:
            self.ref_path_label.grid_remove()
            self.ref_path_entry.grid_remove()
            self.ref_browse_button.grid_remove()
            self.ref_path_var.set("")

    def update_status_bar(self, message):
        self.status_bar.config(text=message)


def main():
    root = tk.Tk()
    app = ImportExportGui(root)
    root.mainloop()


if __name__ == "__main__":
    main()

