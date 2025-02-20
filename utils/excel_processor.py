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

    def _normalize_columns(self, df):
        """Normalize column names by stripping whitespace and replacing spaces with underscores"""
        df.columns = df.columns.str.strip().str.replace(' ', '_')
        return df

    def _validate_columns(self, df, required_columns, sheet_name):
        """Validate that all required columns are present in the DataFrame"""
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"Available columns in {sheet_name}:", df.columns.tolist())
            raise ValueError(f"Missing required columns in {sheet_name}: {', '.join(missing_columns)}")

    def _convert_numeric_columns(self, df, numeric_columns):
        """Helper method to convert columns to numeric type"""
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(float)
        return df

    def process_attendance_summary(self):
        # Print available columns for debugging
        df = pd.read_excel(self.excel_file, sheet_name="Summary", header=0)
        print("Original columns in Summary sheet:", df.columns.tolist())

        df = self._normalize_columns(df)
        print("Normalized columns in Summary sheet:", df.columns.tolist())

        # Define the expected mapping based on the actual column names
        columns = {
            df.columns[0]: 'Employee_ID',
            df.columns[1]: 'Employee_Name',
            df.columns[2]: 'Department',
            df.columns[3]: 'Required_Hours',
            df.columns[4]: 'Actual_Hours',
            df.columns[5]: 'Late_Count',
            df.columns[6]: 'Late_Minutes',
            df.columns[7]: 'Early_Departure_Count',
            df.columns[8]: 'Early_Departure_Minutes',
            df.columns[9]: 'Normal_Overtime',
            df.columns[10]: 'Special_Overtime',
            df.columns[11]: 'Attendance_Ratio',
            df.columns[13]: 'Absences'
        }
        df = df.rename(columns=columns)

        # Validate required columns
        required_columns = [
            'Employee_ID', 'Employee_Name', 'Department',
            'Required_Hours', 'Actual_Hours'
        ]
        self._validate_columns(df, required_columns, "Summary")

        # Convert numeric columns
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
            df.columns[0]: 'Employee_ID',
            df.columns[1]: 'Employee_Name',
            df.columns[2]: 'Department'
        }
        df = df.rename(columns=base_columns)

        # Validate required columns
        required_columns = ['Employee_ID', 'Employee_Name', 'Department']
        self._validate_columns(df, required_columns, "Shifts")

        # Add day columns (D to AH)
        for i in range(3, len(df.columns)):
            df = df.rename(columns={df.columns[i]: f'Day_{i-2}'})
        return df

    def process_records_list(self):
        df = pd.read_excel(self.excel_file, sheet_name="Logs", skiprows=4)
        records = []

        # Process every two rows (name and schedule)
        for i in range(0, len(df), 2):
            if i+1 >= len(df):
                break

            employee_name = df.iloc[i].iloc[0]  # First cell contains employee name
            schedule_row = df.iloc[i+1]

            # Process each time entry (4 entries per day)
            for j in range(0, len(schedule_row), 4):
                if pd.notna(schedule_row.iloc[j]):
                    record = {
                        'Employee_Name': employee_name,
                        'Initial_Entry': schedule_row.iloc[j],
                        'Midday_Exit': schedule_row.iloc[j+1] if j+1 < len(schedule_row) else None,
                        'Midday_Entry': schedule_row.iloc[j+2] if j+2 < len(schedule_row) else None,
                        'Final_Exit': schedule_row.iloc[j+3] if j+3 < len(schedule_row) else None
                    }
                    records.append(record)

        return pd.DataFrame(records)

    def process_individual_reports(self):
        df = pd.read_excel(self.excel_file, sheet_name="Exceptional", header=0)
        df = self._normalize_columns(df)

        columns = {
            df.columns[0]: 'Employee_ID',
            df.columns[1]: 'Employee_Name',
            df.columns[2]: 'Department',
            df.columns[3]: 'Date',
            df.columns[4]: 'AM_Entry',
            df.columns[5]: 'AM_Exit',
            df.columns[6]: 'PM_Entry',
            df.columns[7]: 'PM_Exit',
            df.columns[8]: 'Late_Minutes',
            df.columns[9]: 'Early_Leave_Minutes',
            df.columns[10]: 'Total_Minutes',
            df.columns[11]: 'Comments'
        }
        df = df.rename(columns=columns)

        # Validate required columns
        required_columns = [
            'Employee_ID', 'Employee_Name', 'Department',
            'Late_Minutes', 'Early_Leave_Minutes', 'Total_Minutes'
        ]
        self._validate_columns(df, required_columns, "Individual Reports")

        # Convert numeric columns
        numeric_columns = ['Late_Minutes', 'Early_Leave_Minutes', 'Total_Minutes']
        return self._convert_numeric_columns(df, numeric_columns)