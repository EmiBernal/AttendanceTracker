import pandas as pd
import numpy as np
from difflib import get_close_matches
from datetime import datetime, timedelta

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
            df.columns = [
                '_'.join(str(level) for level in col).strip()
                for col in df.columns.values
            ]
        df.columns = df.columns.str.strip().str.replace(' ', '_').str.lower()
        return df

    def process_attendance_summary(self):
        print("Reading Summary sheet...")
        df = pd.read_excel(self.excel_file, sheet_name="Summary", header=[2, 3])

        # Normalize column names
        df = self._normalize_columns(df)

        # Map standard column names
        column_mapping = {
            'no._unnamed:_0_level_1': 'employee_id',
            'name_unnamed:_1_level_1': 'employee_name',
            'department_unnamed:_2_level_1': 'department',
            'work_hrs._required': 'required_hours',
            'work_hrs._actual': 'actual_hours',
            'late_times': 'late_count',
            'late__min.': 'late_minutes',
            'early_leave_times': 'early_departure_count',
            'early_leave__min.': 'early_departure_minutes',
            'overtime_regular': 'overtime_regular',
            'overtime_special': 'overtime_special',
            'attend_(required/actual)_unnamed:_11_level_1': 'attendance_ratio',
            'absence_unnamed:_13_level_1': 'absences',
            'on_leave_unnamed:_14_level_1': 'on_leave'
        }

        df = df.rename(columns=column_mapping)

        # Convert numeric columns
        numeric_columns = [
            'required_hours', 'actual_hours', 'late_minutes',
            'early_departure_minutes', 'overtime_regular', 'overtime_special',
            'late_count', 'early_departure_count', 'absences', 'on_leave'
        ]

        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        return df

    def process_shift_table(self):
        print("Reading Shifts sheet...")
        df = pd.read_excel(self.excel_file, sheet_name="Shifts",
                           header=[2, 3])  # Leer desde fila 3 y 4
        print("Original columns:", df.columns.tolist())

        df = self._normalize_columns(df)
        print("Normalized columns:", df.columns.tolist())

        # Mapeo de nombres correctos
        column_mapping = {
            'no._unnamed:_0_level_1': 'employee_id',
            'name_unnamed:_1_level_1': 'employee_name',
            'department_unnamed:_2_level_1': 'department'
        }

        df = df.rename(columns=column_mapping)

        required_columns = ['employee_id', 'employee_name', 'department']
        df = self._validate_columns(df, required_columns, "Shifts")

        # Renombrar las columnas de los días
        for i in range(3, len(df.columns)):
            df = df.rename(columns={df.columns[i]: f'day_{i-2}'})

        return df

    def process_records_list(self):
        df = pd.read_excel(self.excel_file, sheet_name="Logs", skiprows=4)

        records = []
        for i in range(0, len(df), 2):
            if i + 1 >= len(df):
                break

            employee_name = df.iloc[i].iloc[1]  # Name is in column B
            schedule_row = df.iloc[i + 1]

            # Process each set of 4 columns (entry, exit, entry, exit)
            for j in range(2, len(schedule_row), 4):
                record = {
                    'employee_name': employee_name,
                    'date': df.columns[j].strftime('%Y-%m-%d') if isinstance(df.columns[j], datetime) else None,
                    'initial_entry': schedule_row.iloc[j],
                    'midday_exit': schedule_row.iloc[j + 1] if j + 1 < len(schedule_row) else None,
                    'midday_entry': schedule_row.iloc[j + 2] if j + 2 < len(schedule_row) else None,
                    'final_exit': schedule_row.iloc[j + 3] if j + 3 < len(schedule_row) else None
                }

                # Calculate lunch duration if both midday times exist
                if pd.notna(record['midday_exit']) and pd.notna(record['midday_entry']):
                    lunch_duration = (pd.to_datetime(record['midday_entry']) - 
                                   pd.to_datetime(record['midday_exit'])).total_seconds() / 60
                    record['lunch_duration'] = lunch_duration
                    record['extended_lunch'] = lunch_duration > 20 if not pd.isna(lunch_duration) else False
                else:
                    record['lunch_duration'] = None
                    record['extended_lunch'] = False

                # Check for missing records
                record['missing_records'] = (
                    pd.isna(record['initial_entry']) or
                    pd.isna(record['midday_exit']) or
                    pd.isna(record['midday_entry']) or
                    pd.isna(record['final_exit'])
                )

                records.append(record)

        return pd.DataFrame(records)


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
            'no._unnamed:_0_level_1': 'employee_id',
            'name_unnamed:_1_level_1': 'employee_name',
            'department_unnamed:_2_level_1': 'department',
            'date_unnamed:_3_level_1': 'date',
            'late_in_(mm)_unnamed:_8_level_1': 'late_in_minutes',
            'early_leave_(mm)_unnamed:_9_level_1': 'early_leave_minutes',
            'total_(mm)_unnamed:_10_level_1': 'total_minutes',
            'remark_unnamed:_11_level_1': 'remark'
        }

        df = df.rename(columns=column_mapping)

        required_columns = [
            'employee_id', 'employee_name', 'department', 'date', 'late_in_minutes',
            'early_leave_minutes', 'total_minutes', 'remark'
        ]
        df = self._validate_columns(df, required_columns, "Exceptional")

        return df

    def get_employee_stats(self, employee_name):
        summary = self.process_attendance_summary()
        records = self.process_records_list()

        employee_summary = summary[summary['employee_name'] == employee_name].iloc[0]
        employee_records = records[records['employee_name'] == employee_name]

        stats = {
            'name': employee_name,
            'department': employee_summary['department'],
            'required_hours': float(employee_summary['required_hours']),
            'actual_hours': float(employee_summary['actual_hours']),
            'late_days': int(employee_summary['late_count']),
            'late_minutes': float(employee_summary['late_minutes']),
            'early_departures': int(employee_summary['early_departure_count']),
            'early_minutes': float(employee_summary['early_departure_minutes']),
            'absences': int(employee_summary['absences']),
            'extended_lunch_days': len(employee_records[employee_records['extended_lunch']]),
            'missing_record_days': len(employee_records[employee_records['missing_records']]),
            'attendance_ratio': float(employee_summary['attendance_ratio'])
        }

        return stats