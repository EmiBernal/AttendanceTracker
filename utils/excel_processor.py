import pandas as pd
import numpy as np
from difflib import get_close_matches

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

    def _find_column_match(self, expected_name, available_columns):
        """Find the closest matching column name"""
        matches = get_close_matches(expected_name, available_columns, n=1, cutoff=0.6)
        if matches:
            return matches[0]
        return None

    def _normalize_columns(self, df):
        """Normalize column names by stripping whitespace and replacing spaces with underscores"""
        df.columns = df.columns.str.strip().str.replace(' ', '_').str.replace(':', '')
        return df

    def _validate_columns(self, df, required_columns, sheet_name):
        """Validate that all required columns are present in the DataFrame"""
        missing_columns = []
        column_mapping = {}

        for col in required_columns:
            if col not in df.columns:
                match = self._find_column_match(col, df.columns)
                if match:
                    column_mapping[match] = col
                else:
                    missing_columns.append(col)

        if missing_columns:
            print(f"Available columns in {sheet_name}:", df.columns.tolist())
            raise ValueError(f"Missing required columns in {sheet_name}: {', '.join(missing_columns)}")

        if column_mapping:
            df = df.rename(columns=column_mapping)

        return df

    def _convert_numeric_columns(self, df, numeric_columns):
        """Helper method to convert columns to numeric type"""
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(float)
        return df

    def process_attendance_summary(self):
        print("Reading Summary sheet...")
        df = pd.read_excel(self.excel_file, sheet_name="Summary", header=0)
        print("Original columns:", df.columns.tolist())

        df = self._normalize_columns(df)
        print("Normalized columns:", df.columns.tolist())

        # Expected column mapping
        columns = {
            'ID': 'Employee_ID',
            'Name': 'Employee_Name',
            'Department': 'Department',
            'Required': 'Required_Hours',
            'Actual': 'Actual_Hours',
            'Late_Count': 'Late_Count',
            'Late_Min': 'Late_Minutes',
            'Early_Count': 'Early_Departure_Count',
            'Early_Min': 'Early_Departure_Minutes',
            'Normal_OT': 'Normal_Overtime',
            'Special_OT': 'Special_Overtime',
            'Attendance': 'Attendance_Ratio',
            'Absences': 'Absences'
        }

        # Try to find and rename columns
        for expected, target in columns.items():
            if target not in df.columns:
                match = self._find_column_match(expected, df.columns)
                if match:
                    df = df.rename(columns={match: target})

        required_columns = [
            'Employee_ID', 'Employee_Name', 'Department',
            'Required_Hours', 'Actual_Hours'
        ]
        df = self._validate_columns(df, required_columns, "Summary")

        numeric_columns = [
            'Required_Hours', 'Actual_Hours', 
            'Late_Count', 'Late_Minutes',
            'Early_Departure_Count', 'Early_Departure_Minutes',
            'Normal_Overtime', 'Special_Overtime',
            'Attendance_Ratio', 'Absences'
        ]
        return self._convert_numeric_columns(df, numeric_columns)

    def process_shift_table(self):
        df = pd.read_excel(self.excel_file, sheet_name="Shifts", header=0)
        df = self._normalize_columns(df)

        base_columns = {
            'ID': 'Employee_ID',
            'Name': 'Employee_Name',
            'Dept': 'Department'
        }

        for expected, target in base_columns.items():
            if target not in df.columns:
                match = self._find_column_match(expected, df.columns)
                if match:
                    df = df.rename(columns={match: target})

        required_columns = ['Employee_ID', 'Employee_Name', 'Department']
        df = self._validate_columns(df, required_columns, "Shifts")

        # Add day columns (D to AH)
        for i in range(3, len(df.columns)):
            df = df.rename(columns={df.columns[i]: f'Day_{i-2}'})
        return df

    def process_records_list(self):
        df = pd.read_excel(self.excel_file, sheet_name="Logs", skiprows=4)
        records = []

        for i in range(0, len(df), 2):
            if i+1 >= len(df):
                break

            employee_name = df.iloc[i].iloc[0]
            schedule_row = df.iloc[i+1]

            for j in range(0, len(schedule_row), 4):
                if pd.notna(schedule_row.iloc[j]):
                    record = {
                        'Employee_Name': employee_name,
                        'Initial_Entry': pd.to_datetime(schedule_row.iloc[j], errors='coerce'),
                        'Midday_Exit': pd.to_datetime(schedule_row.iloc[j+1], errors='coerce') if j+1 < len(schedule_row) else None,
                        'Midday_Entry': pd.to_datetime(schedule_row.iloc[j+2], errors='coerce') if j+2 < len(schedule_row) else None,
                        'Final_Exit': pd.to_datetime(schedule_row.iloc[j+3], errors='coerce') if j+3 < len(schedule_row) else None
                    }
                    records.append(record)

        return pd.DataFrame(records)

    def process_individual_reports(self):
        df = pd.read_excel(self.excel_file, sheet_name="Exceptional", header=0)
        df = self._normalize_columns(df)

        columns = {
            'ID': 'Employee_ID',
            'Name': 'Employee_Name',
            'Dept': 'Department',
            'Date': 'Date',
            'AM_In': 'AM_Entry',
            'AM_Out': 'AM_Exit',
            'PM_In': 'PM_Entry',
            'PM_Out': 'PM_Exit',
            'Late': 'Late_Minutes',
            'Early': 'Early_Leave_Minutes',
            'Total': 'Total_Minutes',
            'Comments': 'Comments'
        }

        for expected, target in columns.items():
            if target not in df.columns:
                match = self._find_column_match(expected, df.columns)
                if match:
                    df = df.rename(columns={match: target})

        required_columns = [
            'Employee_ID', 'Employee_Name', 'Department',
            'Late_Minutes', 'Early_Leave_Minutes', 'Total_Minutes'
        ]
        df = self._validate_columns(df, required_columns, "Individual Reports")

        numeric_columns = ['Late_Minutes', 'Early_Leave_Minutes', 'Total_Minutes']
        return self._convert_numeric_columns(df, numeric_columns)