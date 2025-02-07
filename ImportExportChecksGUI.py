from ReqIF2ExelConverter import ReqIF2ExcelProcessor
from ImportExportChecksExcel import ChecksProcessorExcel, CheckConfiguration
from reqif_utils import ReqIFProcessor
from tkinter import filedialog, ttk, messagebox, PhotoImage
import os
import tkinter as tk
from tkinter import ttk
import sys


class ImportExportGui:
    def __init__(self, master):
        self.master = master
        master.title("Import Export Checker")
        icon_path = ImportExportGui.resource_path(os.path.join('icons', 'check.png'))
        img = PhotoImage(file=icon_path)
        master.iconphoto(False, img)
        master.geometry("600x500")
        master.resizable(False, False)

        # Apply custom styles
        style = ttk.Style()
        style.theme_use('default')

        # Set custom colors
        style.configure('TLabel', background='#f0f0f0', foreground='#333333')
        style.configure('TButton', background='#007bff', foreground='#ffffff',
                        padding=8,
                        font=("Helvetica", 10), relief=tk.FLAT)
        style.map('TButton', background=[('active', '#0069d9')])
        style.configure('TRadiobutton', background='#f0f0f0',
                        foreground='#333333')
        style.configure('TEntry', fieldbackground='#ffffff',
                        foreground='#333333')
        style.configure('TFrame', background='#f0f0f0')

        # Configure custom checkbox style with no focus indicators
        style.layout('NoFocus.TCheckbutton',
                     [('Checkbutton.padding', {'children':
                                                   [('Checkbutton.indicator',
                                                     {'side': 'left',
                                                      'sticky': ''}),
                                                    ('Checkbutton.focus',
                                                     {'children':
                                                         [(
                                                             'Checkbutton.label',
                                                             {
                                                                 'sticky': 'nswe'})],
                                                         'side': 'left',
                                                         'sticky': ''})],
                                               'sticky': 'nswe'})])

        style.configure('NoFocus.TCheckbutton',
                        background='#f0f0f0',
                        foreground='#333333',
                        focuscolor='#f0f0f0',
                        highlightthickness=0,
                        borderwidth=0)

        style.map('NoFocus.TCheckbutton',
                  background=[('active', '#f0f0f0')],
                  foreground=[('active', '#333333')],
                  focuscolor=[('active', '#f0f0f0')],
                  highlightcolor=[('focus', '#f0f0f0')],
                  relief=[('focus', 'flat')])

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


        # Project Selection Frame
        self.project_frame = ttk.Frame(master)
        self.project_frame.pack(side=tk.TOP, fill=tk.X, padx=20, pady=10)

        # Set the label for Select Project
        ttk.Label(self.project_frame, text="Select Project:",
                  font=("Helvetica", 12)).grid(row=0, column=0, sticky="w")

        # Create a dropdown list for project selection
        self.project_var = tk.StringVar(value="PPE/MLBW")
        self.project_dropdown = ttk.Combobox(self.project_frame, textvariable=self.project_var,
                                             postcommand=self.print_status,
                                             values=["PPE/MLBW", "SSP"], state="readonly", style='TCombobox')
        self.project_dropdown.grid(row=0, column=1, padx=10, sticky="w")
        print(f"project selected is: {self.project_var.get()}")

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
        # Comparison Type Selection Frame
        self.comparison_type_frame = ttk.Frame(master)
        self.comparison_type_frame.pack(side=tk.TOP, fill=tk.X, padx=20,
                                        pady=10)
        self.comparison_type_var = tk.StringVar(value="Excel")

        # Set the label for Comparison Type
        ttk.Label(self.comparison_type_frame, text="Comparison Type:",
                  font=("Helvetica", 12)).grid(row=0, column=0, sticky="w")

        # Create Radio buttons for Excel and ReqIF conversion
        ttk.Radiobutton(self.comparison_type_frame, text="Excel Based",
                        variable=self.comparison_type_var,
                        value="Excel",
                        command=self.toggle_conversion_type,
                        style='TRadiobutton').grid(row=0, column=1, padx=10,
                                                   sticky="w")
        ttk.Radiobutton(self.comparison_type_frame, text="ReqIF Based",
                        variable=self.comparison_type_var,
                        value="ReqIF",
                        command=self.toggle_conversion_type,
                        style='TRadiobutton').grid(row=0, column=2, padx=10,
                                                   sticky="w")
        # Paths Frame
        self.path_frame = ttk.Frame(master)
        self.path_frame.pack(side=tk.TOP, fill=tk.X, padx=20, pady=10)

        # Add path entries
        self.add_path_entry(self.path_frame, "ReqIF folder:",
                            self.browse_reqif_path, 0)
        self.add_path_entry(self.path_frame, "Extract folder:",
                            self.browse_unzip_path, 1)
        self.add_path_entry(self.path_frame, "Excel folder:",
                            self.browse_excel_path, 2)
        self.add_path_entry(self.path_frame, "Compare file:",
                            self.browse_reference_path, 3)

        # Buttons Frame
        self.button_frame = ttk.Frame(master)
        self.button_frame.pack(side=tk.TOP, fill=tk.X, padx=20, pady=20)

        self.convert_button = ttk.Button(self.button_frame, text="Convert",
                                         command=self.convert_files,
                                         style='TButton')

        self.execute_button = ttk.Button(self.button_frame,
                                         text="Execute Checks",
                                         command=self.execute_checks,
                                         style='TButton', stat=tk.DISABLED)

        self.execute_reqif_button = ttk.Button(self.button_frame,
                                                text="Execute",
                                                command=self.execute_reqif_checks,
                                                style='TButton')

        # Default view: Show Convert and Execute Checks buttons
        self.convert_button.pack(side=tk.LEFT, padx=20)
        self.execute_button.pack(side=tk.LEFT, padx=20)

        # Report Type Selection Frame
        self.report_type_frame = ttk.Frame(master)
        self.report_type_frame.pack(side=tk.TOP, fill=tk.X, padx=20, pady=10)

        # Report Type Label
        ttk.Label(self.report_type_frame, text="Report Type:",
                  font=("Helvetica", 12)).grid(row=0, column=0, sticky="w")

        # Report Type Variable
        self.report_type_var = tk.StringVar(value="HTML")  # Default: HTML

        # Excel Report Radio Button
        ttk.Radiobutton(self.report_type_frame, text="HTML",
                        variable=self.report_type_var,
                        value="HTML",
                        style='TRadiobutton').grid(row=0, column=1, padx=10,
                                                   sticky="w")

        # HTML Report Radio Button
        ttk.Radiobutton(self.report_type_frame, text="Excel",
                        variable=self.report_type_var,
                        value="Excel",
                        style='TRadiobutton').grid(row=0, column=2, padx=10,
                                                   sticky="w")

        # Status bar
        self.status_bar = ttk.Label(master, text="", relief=tk.SUNKEN,
                                    anchor=tk.W, font=("Helvetica", 10))
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Store the original labels and fields
        self.original_labels = []
        self.original_entries = []

        # Initialize the original labels and fields
        self.initialize_original_fields()

    def initialize_original_fields(self):
        """Initialize the original labels and fields for Excel Conversion."""
        self.original_labels = [
            "ReqIF folder:", "Extract folder:", "Excel folder:",
            "Compare file:"
        ]
        self.original_entries = [
            (self.browse_reqif_path, 0),
            (self.browse_unzip_path, 1),
            (self.browse_excel_path, 2),
            (self.browse_reference_path, 3)
        ]

    def resource_path(relative_path):
        """Get absolute path to resource, works for dev and for PyInstaller"""
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

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

        # If it's the Compare file field, set default text and add focus handlers
        if label_text == "Compare file:":
            entry_var.set("---- Optional ----")

            def on_entry_click(event):
                if entry_var.get() == "---- Optional ----":
                    entry_var.set("")

            def on_focus_out(event):
                if entry_var.get() == "":
                    entry_var.set("---- Optional ----")

            entry.bind('<FocusIn>', on_entry_click)
            entry.bind('<FocusOut>', on_focus_out)

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
        elif label_text == "Compare file:":
            self.ref_path_var = entry_var
        elif label_text == "Customer ReqIF:":
            self.cus_reqif_path_var = entry_var
        elif label_text == "Bosch ReqIF:":
            self.own_reqif_path_var = entry_var

    def browse_reqif_path(self):
        self.reqif_path_var.set(filedialog.askdirectory())

    def browse_unzip_path(self):
        self.unzip_path_var.set(filedialog.askdirectory())

    def browse_excel_path(self):
        self.excel_path_var.set(filedialog.askdirectory())

    def browse_reference_path(self):
        """Open file dialog to select reference Excel file"""
        self.ref_path_var.set(filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        ))

    def browse_cus_reqif_path(self):
        self.cus_reqif_path_var.set(filedialog.askdirectory())

    def browse_bosch_reqif_path(self):
        self.own_reqif_path_var.set(filedialog.askdirectory())

    def toggle_conversion_type(self):
        """Toggle the visibility of path entries based on the selected conversion type."""
        if self.comparison_type_var.get() == "Excel":
            # Clear the path frame
            for widget in self.path_frame.winfo_children():
                widget.destroy()

            # Restore the original labels and fields
            for label_text, (browse_command, row) in zip(self.original_labels,
                                                         self.original_entries):
                self.add_path_entry(self.path_frame, label_text,
                                    browse_command, row)

            # Hide the single Execute button and show Convert and Execute Checks buttons
            self.execute_reqif_button.pack_forget()
            self.convert_button.pack(side=tk.LEFT, padx=20)
            self.execute_button.pack(side=tk.LEFT, padx=20)
        else:
            for widget in self.path_frame.winfo_children():
                widget.destroy()

            # Add fields for ReqIF Conversion
            self.add_path_entry(self.path_frame, "Customer ReqIF:",
                                self.browse_cus_reqif_path, 0)
            self.add_path_entry(self.path_frame, "Bosch ReqIF:",
                                self.browse_bosch_reqif_path, 1)

            # Hide Convert and Execute Checks buttons and show the single Execute button
            self.convert_button.pack_forget()
            self.execute_button.pack_forget()
            self.execute_reqif_button.pack(side=tk.LEFT, padx=20)

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
        print(f"Project Type: {self.project_var.get()}")
        print(f"\n{check_type} Checks Active")

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
        project_type = self.project_var.get()
        print(f"execute_checks for Project: {project_type}")
        check_type = self.check_type_var.get()
        print(f"Checks type is: {check_type}")
        report_type = self.report_type_var.get()
        print(f"Report type is: {report_type}")

        self.update_status_bar(
            f"{self.operation_type()} Checks processing started...")
        self.master.update()  # Updates the Tkinter GUI before continuing
        reference_file = self.ref_path_var.get() if self.ref_path_var.get() != "---- Optional ----" else None
        print(f"Path of the refernce file is:  '{reference_file}'")

        processor = ChecksProcessorExcel(project_type, check_type, self.excel_path_var.get(),
                                         reference_file, report_type)
        reports = processor.process_folder()
        self.update_status_bar(
            f"Processed {len(reports)} files. Check reports in {CheckConfiguration.REPORT_FOLDER}")

    def execute_reqif_checks(self):
        """Execute checks for ReqIF Conversion."""
        customer_reqif_path = self.cus_reqif_path_var.get()
        own_reqif_path = self.own_reqif_path_var.get()

        if not customer_reqif_path or not own_reqif_path:
            self.update_status_bar(
                "Please select both Customer ReqIF and Own ReqIF.")
            return

        self.update_status_bar("Executing ReqIF checks...")
        print(f"Customer ReqIF: {customer_reqif_path}")
        print(f"Bosch ReqIF: {own_reqif_path}")

        # Create an instance of ReqIFProcessor
        reqif_processor = ReqIFProcessor()

        try:
            # Extract .reqifz files (if necessary) and get paths to .reqif files
            customer_reqif_file, own_reqif_file = reqif_processor.extract_reqifz_files(
                customer_reqif_path, own_reqif_path
            )
            print(f"Customer ReqIF file: {customer_reqif_file}")
            print(f"Own ReqIF file: {own_reqif_file}")

            # Perform the comparison of the .reqif files
            self.update_status_bar("Comparing ReqIF files...")
            self.compare_reqif_files(customer_reqif_file, own_reqif_file)

        except Exception as e:
            self.update_status_bar(f"Error during ReqIF checks: {str(e)}")
        finally:
            # Clean up temporary directories
            print("Deletion to be perfomed here")
            reqif_processor.cleanup_temp_dirs()

        self.update_status_bar("ReqIF checks completed successfully.")

    def compare_reqif_files(self, customer_reqif_file, own_reqif_file):
        """
        Compare the two .reqif files.
        Add your specific comparison logic here.
        """
        # Example: Print the paths of the files being compared
        print(f"Comparing Customer ReqIF: {customer_reqif_file}")
        print(f"Comparing Own ReqIF: {own_reqif_file}")

        # Add your comparison logic here
        # For example, parse the .reqif files and compare their contents
        # ...

        # ... (rest of the class remains the same)

    def update_status_bar(self, message):
        self.status_bar.config(text=message)

    def print_status(self):
        print(f"Status: {self.project_var.get()}")


def main():
    root = tk.Tk()
    app = ImportExportGui(root)
    root.mainloop()


if __name__ == "__main__":
    main()
