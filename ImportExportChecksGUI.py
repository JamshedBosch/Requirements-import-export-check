from ReqIF2ExelConverter import ReqIF2ExcelProcessor
from ImportExportChecks import ReqIFProcessor, CheckConfiguration
import tkinter as tk
from tkinter import filedialog, ttk

class ImportExportGui:
    def __init__(self, master):
        self.master = master
        master.title("Import Export Checker")
        master.geometry("400x600")

        # create  the menu bar
        menubar = tk.Menu(master)
        master.config(menu=menubar)

        # Create the File menu
        file_menu = tk.Menu(menubar)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open")
        file_menu.add_command(label="Save")
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=master.quit)

        # Create the Help menu
        help_menu = tk.Menu(menubar)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About")

        # Check Type Selection Frame
        self.check_type_frame = tk.Frame(master)
        self.check_type_frame.pack(side=tk.TOP, fill=tk.X, padx=20, pady=10)
        self.check_type_var = tk.IntVar(value=CheckConfiguration.IMPORT_CHECK)
        tk.Label(self.check_type_frame, text="Select Check Type:",
                 font=("Helvetica", 12)).pack(side=tk.LEFT)
        tk.Radiobutton(self.check_type_frame, text="Import Check",
                       variable=self.check_type_var,
                       value=CheckConfiguration.IMPORT_CHECK,
                       font=("Helvetica", 10)).pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(self.check_type_frame, text="Export Check",
                       variable=self.check_type_var,
                       value=CheckConfiguration.EXPORT_CHECK,
                       font=("Helvetica", 10)).pack(side=tk.LEFT, padx=10)

        # Paths Frame
        self.path_frame = tk.Frame(master)
        self.path_frame.pack(side=tk.TOP, fill=tk.X, padx=20, pady=10)

        self.reqif_path_var = tk.StringVar()
        reqif_path_label = tk.Label(self.path_frame, text="ReqIF Path:",
                                    font=("Helvetica", 12))
        reqif_path_label.pack(side=tk.TOP, padx=10, pady=5)
        reqif_path_entry = tk.Entry(self.path_frame,
                                    textvariable=self.reqif_path_var, width=40,
                                    font=("Helvetica", 10))
        reqif_path_entry.pack(side=tk.TOP, padx=10, pady=5)
        reqif_path_button = tk.Button(self.path_frame, text="Browse",
                                      command=self.browse_reqif_path,
                                      font=("Helvetica", 10))
        reqif_path_button.pack(side=tk.TOP, padx=10, pady=5)

        self.unzip_path_var = tk.StringVar()
        unzip_path_label = tk.Label(self.path_frame, text="Unzip Path:",
                                    font=("Helvetica", 12))
        unzip_path_label.pack(side=tk.TOP, padx=10, pady=5)
        unzip_path_entry = tk.Entry(self.path_frame,
                                    textvariable=self.unzip_path_var, width=40,
                                    font=("Helvetica", 10))
        unzip_path_entry.pack(side=tk.TOP, padx=10, pady=5)
        unzip_path_button = tk.Button(self.path_frame, text="Browse",
                                      command=self.browse_unzip_path,
                                      font=("Helvetica", 10))
        unzip_path_button.pack(side=tk.TOP, padx=10, pady=5)

        self.excel_path_var = tk.StringVar()
        excel_path_label = tk.Label(self.path_frame,
                                    text="Generated Excel Path:",
                                    font=("Helvetica", 12))
        excel_path_label.pack(side=tk.TOP, padx=10, pady=5)
        excel_path_entry = tk.Entry(self.path_frame,
                                    textvariable=self.excel_path_var, width=40,
                                    font=("Helvetica", 10))
        excel_path_entry.pack(side=tk.TOP, padx=10, pady=5)
        excel_path_button = tk.Button(self.path_frame, text="Browse",
                                      command=self.browse_excel_path,
                                      font=("Helvetica", 10))
        excel_path_button.pack(side=tk.TOP, padx=10, pady=5)

        # Buttons Frame
        self.button_frame = tk.Frame(master)
        self.button_frame.pack(side=tk.TOP, fill=tk.X, padx=20, pady=10)
        convert_button = tk.Button(self.button_frame, text="Convert",
                                   command=self.convert_files,
                                   font=("Helvetica", 12))
        convert_button.pack(side=tk.LEFT, padx=10)
        execute_button = tk.Button(self.button_frame, text="Execute Checks",
                                   command=self.execute_checks,
                                   font=("Helvetica", 12))
        execute_button.pack(side=tk.LEFT, padx=10)

        # Status bar
        self.status_bar = tk.Label(master, text="", bd=1, relief=tk.SUNKEN,
                                   anchor=tk.W, font=("Helvetica", 10))
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def browse_reqif_path(self):
        self.reqif_path_var.set(filedialog.askdirectory())

    def browse_unzip_path(self):
        self.unzip_path_var.set(filedialog.askdirectory())

    def browse_excel_path(self):
        self.excel_path_var.set(filedialog.askdirectory())

    def convert_files(self):
        # Implement file conversion logic here
        self.update_status_bar("Files converted successfully.")

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

