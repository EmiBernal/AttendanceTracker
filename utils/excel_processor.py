import pandas as pd
import numpy as np
from datetime import datetime, timedelta

class ExcelProcessor:
    def __init__(self, file):
        self.excel_file = pd.ExcelFile(file)
        self.WORK_START_TIME = datetime.strptime('7:50', '%H:%M').time()
        self.WORK_END_TIME = datetime.strptime('17:10', '%H:%M').time()

    def process_attendance_summary(self):
        """Procesa los datos de asistencia desde las hojas después de 'Exceptional'"""
        exceptional_index = self.excel_file.sheet_names.index('Exceptional')
        attendance_sheets = self.excel_file.sheet_names[exceptional_index:]

        all_records = []

        for sheet in attendance_sheets:
            df = pd.read_excel(self.excel_file, sheet_name=sheet)

            # Procesar cada grupo de 3 empleados
            employee_groups = [
                # Primer empleado
                {
                    'dept_cols': 'B:H',
                    'name_cols': 'J:N',
                    'date_row': 4,
                    'absence_col': 'A',
                    'late_col': 'I',
                    'late_mins_cols': ['J', 'K'],
                    'early_cols': ['L', 'M'],
                    'early_mins_col': 'N',
                    'time_cols': {'day': 'A', 'entry': 'B', 'lunch_out': 'D', 'lunch_in': 'G', 'exit': 'I'}
                },
                # Segundo empleado
                {
                    'dept_cols': 'Q:W',
                    'name_cols': 'Y:AC',
                    'date_row': 4,
                    'absence_col': 'P',
                    'late_col': 'X',
                    'late_mins_cols': ['Y', 'Z'],
                    'early_cols': ['AA', 'AB'],
                    'early_mins_col': 'AC',
                    'time_cols': {'day': 'P', 'entry': 'Q', 'lunch_out': 'S', 'lunch_in': 'V', 'exit': 'X'}
                },
                # Tercer empleado
                {
                    'dept_cols': 'AF:AL',
                    'name_cols': 'AN:AR',
                    'date_row': 4,
                    'absence_col': 'AE',
                    'late_col': 'AM',
                    'late_mins_cols': ['AN', 'AO'],
                    'early_cols': ['AP', 'AQ'],
                    'early_mins_col': 'AR',
                    'time_cols': {'day': 'AE', 'entry': 'AF', 'lunch_out': 'AH', 'lunch_in': 'AK', 'exit': 'AM'}
                }
            ]

            for group in employee_groups:
                try:
                    # Obtener información del empleado
                    department = str(df.iloc[2, df.columns.get_loc(group['dept_cols'].split(':')[0])]).strip()
                    name = str(df.iloc[2, df.columns.get_loc(group['name_cols'].split(':')[0])]).strip()

                    if pd.isna(name) or name == '':
                        continue

                    # Obtener estadísticas básicas
                    absences = float(df.iloc[6, df.columns.get_loc(group['absence_col'])]) if pd.notna(df.iloc[6, df.columns.get_loc(group['absence_col'])]) else 0
                    late_count = float(df.iloc[6, df.columns.get_loc(group['late_col'])]) if pd.notna(df.iloc[6, df.columns.get_loc(group['late_col'])]) else 0
                    late_minutes = sum(float(df.iloc[6, df.columns.get_loc(col)]) if pd.notna(df.iloc[6, df.columns.get_loc(col)]) else 0 for col in group['late_mins_cols'])
                    early_count = sum(float(df.iloc[6, df.columns.get_loc(col)]) if pd.notna(df.iloc[6, df.columns.get_loc(col)]) else 0 for col in group['early_cols'])
                    early_minutes = float(df.iloc[6, df.columns.get_loc(group['early_mins_col'])]) if pd.notna(df.iloc[6, df.columns.get_loc(group['early_mins_col'])]) else 0

                    # Calcular horas trabajadas
                    required_hours = 160  # Ejemplo: 8 horas * 20 días
                    actual_hours = required_hours - (absences * 8) - (late_minutes / 60) - (early_minutes / 60)

                    record = {
                        'employee_name': name,
                        'department': department,
                        'required_hours': required_hours,
                        'actual_hours': actual_hours,
                        'late_count': late_count,
                        'late_minutes': late_minutes,
                        'early_departure_count': early_count,
                        'early_departure_minutes': early_minutes,
                        'absences': absences,
                        'attendance_ratio': (required_hours - (absences * 8)) / required_hours if required_hours > 0 else 0
                    }

                    all_records.append(record)

                except Exception as e:
                    print(f"Error processing employee in sheet {sheet}: {str(e)}")
                    continue

        return pd.DataFrame(all_records)

    def get_employee_stats(self, employee_name):
        """Obtiene estadísticas para un empleado específico"""
        summary = self.process_attendance_summary()
        employee_summary = summary[summary['employee_name'] == employee_name].iloc[0]

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
            'missing_entry_days': 0,  # These will be calculated later
            'missing_exit_days': 0,   # These will be calculated later
            'attendance_ratio': float(employee_summary['attendance_ratio'])
        }

        return stats