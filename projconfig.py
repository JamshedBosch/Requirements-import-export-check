import os

class CheckConfiguration:
    """Holds configuration and constants for checks."""
    IMPORT_CHECK = 0
    EXPORT_CHECK = 1

    PROJECT = {
        "PPE_MLBW": "PPE/MLBW",
        "SSP": "SSP"
    }

    REPORT_FOLDER = os.path.join(os.getcwd(), "report")


    IMPORT_FOLDERS = {
        IMPORT_CHECK: r"D:\AUDI\Import_Reqif2Excel_Converted",
        EXPORT_CHECK: r"D:\AUDI\Export_Reqif2Excel_Converted"
    }
