import pandas as pd
import numpy as np
from difflib import get_close_matches


class ExcelProcessor:

    def __init__(self, file):
        self.excel_file = pd.ExcelFile(file)
        self.validate_sheets()

    def validate_sheets(self):
        required_sheets = ["Summary", "Shifts", "Logs", "Exceptional"]
        missing_sheets = [
            sheet for sheet in required_sheets
            if sheet not in self.excel_file.sheet_names
        ]
        if missing_sheets:
            raise ValueError(
                f"Missing required sheets: {', '.join(missing_sheets)}")

    def _find_column_match(self, expected_name, available_columns):
        """Find the closest matching column name"""
        matches = get_close_matches(expected_name,
                                    available_columns,
                                    n=1,
                                    cutoff=0.6)
        return matches[0] if matches else None

    def _normalize_columns(self, df):
        """Normalize column names by stripping whitespace and replacing spaces with underscores"""
        if isinstance(df.columns, pd.MultiIndex):
            # Convertir a string cada elemento antes de unir
            df.columns = [
                ' '.join(map(str, col)).strip() for col in df.columns
            ]

        df.columns = df.columns.str.strip().str.replace(' ', '_').str.replace(
            ':', '')
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
            raise ValueError(
                f"Missing required columns in {sheet_name}: {', '.join(missing_columns)}"
            )

        if column_mapping:
            df = df.rename(columns=column_mapping)

        return df

    def _convert_numeric_columns(self, df, numeric_columns):
        """Helper method to convert columns to numeric type"""
        for col in numeric_columns:
            if col in df.columns:  # Verificamos si la columna existe
                try:
                    print(f"Converting column: {col}")
                    # Verificamos si es una serie antes de convertir
                    if isinstance(df[col], pd.Series):
                        # Si el tipo es 'object' (cadenas), intentamos convertir a numérico
                        df[col] = pd.to_numeric(
                            df[col], errors='coerce'
                        )  # Convertimos a numérico, invalidos serán NaN
                        df[col] = df[col].fillna(0).astype(
                            float
                        )  # Rellenamos NaN con 0 y convertimos a float
                    else:
                        print(f"Skipping non-Series column: {col}")
                except Exception as e:
                    print(f"Error processing column {col}: {e}")
            else:
                print(f"Column {col} is missing.")
        return df

    def process_attendance_summary(self):
        print("Reading Summary sheet...")
        # Cambiar header a 2, 3 (especifica que los encabezados están en las filas 3 y 4)
        df = pd.read_excel(self.excel_file,
                           sheet_name="Summary",
                           header=[2, 3])
        print("Original columns:", df.columns.tolist())

        df = self._normalize_columns(df)
        print("Normalized columns:", df.columns.tolist())

        # Mapeo de nombres correctos
        column_mapping = {
            'No._Unnamed_0_level_1': 'Employee_ID',
            'Name_Unnamed_1_level_1': 'Employee_Name',
            'Department_Unnamed_2_level_1': 'Department',
            'Work_Hrs._Required': 'Required_Hours',
            'Work_Hrs._Actual': 'Actual_Hours',
            'Late_Times': 'Late_Minutes',
            'Late__Min.': 'Late_Minutes',
            'Early_Leave_Times': 'Early_Departure_Minutes',
            'Early_Leave__Min.': 'Early_Departure_Minutes',
            'Overtime_Regular': 'Overtime_Regular',
            'Overtime_Special': 'Overtime_Special',
            'Attend_(Required/Actual)_Unnamed_11_level_1': 'Attendance_Status',
            'Business_Trip_Unnamed_12_level_1': 'Business_Trip',
            'Absence_Unnamed_13_level_1': 'Absence',
            'On_Leave_Unnamed_14_level_1': 'On_Leave',
            'Bonus_Note': 'Bonus_Note',
            'Bonus_OT': 'Bonus_OT',
            'Bonus_All_owance': 'Bonus_Allowance',
            'Deduction_Late_Early_Leave': 'Deduction_Late_Early_Leave',
            'Deduction_On_Leave': 'Deduction_On_Leave',
            'Deduction_Other\\nDeduction': 'Deduction_Other_Deduction',
            'Actual_Pay_Unnamed_21_level_1': 'Actual_Pay',
            'Remark_Unnamed_22_level_1': 'Remark'
        }

        df = df.rename(columns=column_mapping)

        required_columns = [
            'Employee_ID', 'Employee_Name', 'Department', 'Required_Hours',
            'Actual_Hours', 'Late_Minutes', 'Early_Departure_Minutes'
        ]
        df = self._validate_columns(df, required_columns, "Summary")

        numeric_columns = [
            'Required_Hours', 'Actual_Hours', 'Late_Minutes',
            'Early_Departure_Minutes'
        ]
        return self._convert_numeric_columns(df, numeric_columns)

    def process_shift_table(self):
        print("Reading Shifts sheet...")
        df = pd.read_excel(self.excel_file, sheet_name="Shifts",
                           header=[2, 3])  # Leer desde fila 3 y 4
        print("Original columns:", df.columns.tolist())

        df = self._normalize_columns(df)
        print("Normalized columns:", df.columns.tolist())

        # Mapeo de nombres correctos
        column_mapping = {
            'No._Unnamed_0_level_1': 'Employee_ID',
            'Name_Unnamed_1_level_1': 'Employee_Name',
            'Department_Unnamed_2_level_1': 'Department'
        }

        df = df.rename(columns=column_mapping)

        required_columns = ['Employee_ID', 'Employee_Name', 'Department']
        df = self._validate_columns(df, required_columns, "Shifts")

        # Renombrar las columnas de los días
        for i in range(3, len(df.columns)):
            df = df.rename(columns={df.columns[i]: f'Day_{i-2}'})

        return df

    def process_records_list(self):
        df = pd.read_excel(self.excel_file, sheet_name="Logs", skiprows=4)

        records = []
        for i in range(0, len(df), 2):
            if i + 1 >= len(df):
                break

            employee_id = df.iloc[i].iloc[0]  # "No."
            employee_name = df.iloc[i].iloc[1]  # "Name"
            schedule_row = df.iloc[i + 1]

            for j in range(0, len(schedule_row), 4):
                if pd.notna(schedule_row.iloc[j]):
                    record = {
                        'Employee_ID':
                        employee_id,
                        'Employee_Name':
                        employee_name,
                        'Initial_Entry':
                        pd.to_datetime(schedule_row.iloc[j], errors='coerce'),
                        'Midday_Exit':
                        pd.to_datetime(schedule_row.iloc[j + 1],
                                       errors='coerce') if j +
                        1 < len(schedule_row) else None,
                        'Midday_Entry':
                        pd.to_datetime(schedule_row.iloc[j + 2],
                                       errors='coerce') if j +
                        2 < len(schedule_row) else None,
                        'Final_Exit':
                        pd.to_datetime(schedule_row.iloc[j + 3],
                                       errors='coerce') if j +
                        3 < len(schedule_row) else None
                    }
                    records.append(record)

        return pd.DataFrame(records)

    def process_individual_reports(self):
        print("Reading Exceptional sheet...")
        df = pd.read_excel(self.excel_file,
                           sheet_name="Exceptional",
                           header=[2, 3])  # Leer encabezados desde fila 3 y 4
        print("Original columns:", df.columns.tolist())

        df = self._normalize_columns(df)
        print("Normalized columns:", df.columns.tolist())

        # Mapeo de nombres correctos
        column_mapping = {
            'No._Unnamed_0_level_1': 'Employee_ID',
            'Name_Unnamed_1_level_1': 'Employee_Name',
            'Department_Unnamed_2_level_1': 'Department',
            'Date_Unnamed_3_level_1': 'Date',
            'Late_in_(mm)_Unnamed_8_level_1': 'Late_In_Minutes',
            'Early_Leave_(mm)_Unnamed_9_level_1': 'Early_Leave_Minutes',
            'Total_(mm)_Unnamed_10_level_1': 'Total_Minutes',
            'Remark_Unnamed_11_level_1': 'Remark'
        }

        df = df.rename(columns=column_mapping)

        required_columns = [
            'Employee_ID', 'Employee_Name', 'Department', 'Date', 'AM_In',
            'AM_Out', 'PM_In', 'PM_Out', 'Late_In_Minutes',
            'Early_Leave_Minutes', 'Total_Minutes', 'Remark'
        ]
        df = self._validate_columns(df, required_columns, "Exceptional")

        return df
