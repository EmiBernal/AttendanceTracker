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

    def _clean_column_name(self, col_name):
        """Clean column names by removing special characters and normalizing spaces"""
        if isinstance(col_name, tuple):
            # Join multi-level column names and clean
            col_name = '_'.join(str(level).strip() for level in col_name if pd.notna(level))
        return (str(col_name)
                .strip()
                .replace('\n', '_')
                .replace(':', '')
                .replace('.', '_')
                .replace(' ', '_')
                .replace('(', '')
                .replace(')', '')
                .replace('/', '_')
                .lower())

    def _normalize_columns(self, df):
        """Normalize column names handling multi-level headers"""
        if isinstance(df.columns, pd.MultiIndex):
            # Create a dictionary to map original column names to cleaned ones
            column_mapping = {}
            for col in df.columns:
                cleaned_name = self._clean_column_name(col)
                column_mapping[col] = cleaned_name

            # Rename columns using the mapping
            df.columns = [column_mapping[col] for col in df.columns]
        else:
            df.columns = [self._clean_column_name(col) for col in df.columns]

        return df

    def _find_column_match(self, expected_name, available_columns):
        """Find the closest matching column name"""
        matches = get_close_matches(expected_name,
                                    available_columns,
                                    n=1,
                                    cutoff=0.6)
        return matches[0] if matches else None

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

    def process_attendance_summary(self):
        print("Reading Summary sheet...")
        df = pd.read_excel(self.excel_file, sheet_name="Summary", header=[2, 3])

        # Mapear directamente las columnas del Excel
        column_mapping = {
            ('No.', 'Unnamed: 0_level_1'): 'employee_id',
            ('Name', 'Unnamed: 1_level_1'): 'employee_name',
            ('Department', 'Unnamed: 2_level_1'): 'department',
            ('Work Hrs.', 'Required'): 'required_hours',
            ('Work Hrs.', 'Actual'): 'actual_hours',
            ('Late', 'Times'): 'late_count',
            ('Late', ' Min.'): 'late_minutes',
            ('Early Leave', 'Times'): 'early_departure_count',
            ('Early Leave', ' Min.'): 'early_departure_minutes',
            ('Overtime', 'Regular'): 'overtime_regular',
            ('Overtime', 'Special'): 'overtime_special',
            ('Attend (Required/Actual)', 'Unnamed: 11_level_1'): 'attendance_ratio',
            ('Absence', 'Unnamed: 13_level_1'): 'absences'
        }

        # Renombrar columnas
        df = df.rename(columns=column_mapping)

        # Convertir columnas numéricas
        numeric_columns = [
            'required_hours', 'actual_hours', 'late_minutes',
            'early_departure_minutes', 'overtime_regular', 'overtime_special',
            'late_count', 'early_departure_count', 'absences'
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
            'no_unnamed_0_level_1': 'employee_id',
            'name_unnamed_1_level_1': 'employee_name',
            'department_unnamed_2_level_1': 'department'
        }

        df = df.rename(columns=column_mapping)

        required_columns = ['employee_id', 'employee_name', 'department']
        df = self._validate_columns(df, required_columns, "Shifts")

        # Renombrar las columnas de los días
        for i in range(3, len(df.columns)):
            df = df.rename(columns={df.columns[i]: f'Day_{i-2}'})

        return df

    def process_exceptional_records(self):
        """Procesa los registros excepcionales para calcular pausas de almuerzo prolongadas"""
        df = pd.read_excel(self.excel_file, sheet_name="Exceptional", header=[2, 3])

        # Mapear las columnas del Excel
        column_mapping = {
            ('No.', 'Unnamed: 0_level_1'): 'employee_id',
            ('Name', 'Unnamed: 1_level_1'): 'employee_name',
            ('Department', 'Unnamed: 2_level_1'): 'department',
            ('Date', 'Unnamed: 3_level_1'): 'date',
            ('AM', 'In'): 'am_in',
            ('AM', 'Out'): 'am_out',
            ('PM', 'In'): 'pm_in',
            ('PM', 'Out'): 'pm_out'
        }

        df = df.rename(columns=column_mapping)

        # Convertir columnas de tiempo a datetime
        time_columns = ['am_in', 'am_out', 'pm_in', 'pm_out']
        for col in time_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], format='%H:%M:%S', errors='coerce')

        # Calcular pausas de almuerzo prolongadas
        df['lunch_duration'] = None
        df['extended_lunch'] = False
        df['lunch_minutes_exceeded'] = 0

        for idx, row in df.iterrows():
            if pd.notna(row['am_out']) and pd.notna(row['pm_in']):
                am_out_time = row['am_out'].time()
                pm_in_time = row['pm_in'].time()

                # Verificar si la salida está entre 12:00 y 16:00
                if (datetime.strptime('12:00', '%H:%M').time() <= am_out_time <=
                    datetime.strptime('16:00', '%H:%M').time()):

                    lunch_duration = (row['pm_in'] - row['am_out']).total_seconds() / 60
                    df.at[idx, 'lunch_duration'] = lunch_duration

                    if lunch_duration > 20:
                        df.at[idx, 'extended_lunch'] = True
                        df.at[idx, 'lunch_minutes_exceeded'] = lunch_duration - 20

        return df


    def get_employee_stats(self, employee_name):
        """Obtiene estadísticas completas para un empleado"""
        summary = self.process_attendance_summary()
        exceptional = self.process_exceptional_records()

        employee_summary = summary[summary['employee_name'] == employee_name].iloc[0]
        employee_exceptional = exceptional[exceptional['employee_name'] == employee_name]

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
            'extended_lunch_days': len(employee_exceptional[employee_exceptional['extended_lunch']]),
            'total_lunch_minutes_exceeded': employee_exceptional['lunch_minutes_exceeded'].sum(),
            'attendance_ratio': float(employee_summary['attendance_ratio']) if pd.notna(employee_summary['attendance_ratio']) else 0
        }

        return stats