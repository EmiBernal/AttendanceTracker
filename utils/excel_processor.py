import pandas as pd
import numpy as np

class ExcelProcessor:
    def __init__(self, file):
        self.excel_file = pd.ExcelFile(file)
        self.validate_sheets()

    def validate_sheets(self):
        required_sheets = [
            "Summary",
            "Shifts",
            "Logs",
            "Exceptional"
        ]
        missing_sheets = [sheet for sheet in required_sheets if sheet not in self.excel_file.sheet_names]
        if missing_sheets:
            raise ValueError(f"Missing required sheets: {', '.join(missing_sheets)}")

    def process_attendance_summary(self):
        df = pd.read_excel(self.excel_file, sheet_name="Summary")
        return df.iloc[:, 0:14]  # Columns A-N

    def process_shift_table(self):
        df = pd.read_excel(self.excel_file, sheet_name="Shifts")
        return df.iloc[:, 0:34]  # Columns A-AH

    def process_records_list(self):
        df = pd.read_excel(self.excel_file, sheet_name="Logs", skiprows=4)
        return df

    def process_individual_reports(self):
        df = pd.read_excel(self.excel_file, sheet_name="Exceptional")
        return df.iloc[:, 0:12]  # Columns A-L