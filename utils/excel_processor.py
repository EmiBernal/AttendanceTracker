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
        # Rename columns according to the structure
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
        return df.rename(columns=columns)

    def process_shift_table(self):
        df = pd.read_excel(self.excel_file, sheet_name="Shifts")
        base_columns = {
            df.columns[0]: 'Employee_ID',
            df.columns[1]: 'Employee_Name',
            df.columns[2]: 'Department'
        }
        df = df.rename(columns=base_columns)
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

            # Process each time entry (4 entries per day: initial entry, midday exit, midday entry, final exit)
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
        df = pd.read_excel(self.excel_file, sheet_name="Exceptional")
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
        return df.rename(columns=columns)