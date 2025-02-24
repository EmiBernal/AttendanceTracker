import pandas as pd
import numpy as np
from datetime import datetime, timedelta

class ExcelProcessor:
    def __init__(self, file):
        self.excel_file = pd.ExcelFile(file)
        self.WORK_START_TIME = datetime.strptime('7:50', '%H:%M').time()
        self.WORK_END_TIME = datetime.strptime('17:10', '%H:%M').time()
        self.validate_sheets()

    def validate_sheets(self):
        """Validar que exista la hoja Exceptional"""
        if 'Exceptional' not in self.excel_file.sheet_names:
            raise ValueError("Missing required sheet: Exceptional")

    def process_attendance_summary(self):
        """Procesa los datos de asistencia desde todas las hojas"""
        all_records = []

        # Procesar todas las hojas después de Exceptional
        exceptional_index = self.excel_file.sheet_names.index('Exceptional')
        attendance_sheets = self.excel_file.sheet_names[exceptional_index:]

        for sheet in attendance_sheets:
            try:
                df = pd.read_excel(self.excel_file, sheet_name=sheet)

                # Definir los grupos de empleados y sus columnas correspondientes
                employee_groups = [
                    {
                        'dept_range': (2, 'B:H'),
                        'name_range': (2, 'J:N'),
                        'stats_row': 6,
                        'cols': {
                            'absences': 'A',
                            'late_count': 'I',
                            'late_minutes': ['J', 'K'],
                            'early_count': ['L', 'M'],
                            'early_minutes': 'N'
                        }
                    },
                    {
                        'dept_range': (2, 'Q:W'),
                        'name_range': (2, 'Y:AC'),
                        'stats_row': 6,
                        'cols': {
                            'absences': 'P',
                            'late_count': 'X',
                            'late_minutes': ['Y', 'Z'],
                            'early_count': ['AA', 'AB'],
                            'early_minutes': 'AC'
                        }
                    },
                    {
                        'dept_range': (2, 'AF:AL'),
                        'name_range': (2, 'AN:AR'),
                        'stats_row': 6,
                        'cols': {
                            'absences': 'AE',
                            'late_count': 'AM',
                            'late_minutes': ['AN', 'AO'],
                            'early_count': ['AP', 'AQ'],
                            'early_minutes': 'AR'
                        }
                    }
                ]

                # Procesar cada grupo de empleado
                for group in employee_groups:
                    try:
                        # Extraer nombre y departamento
                        dept_col = group['dept_range'][1].split(':')[0]
                        name_col = group['name_range'][1].split(':')[0]

                        department = str(df.iloc[group['dept_range'][0], df.columns.get_loc(dept_col)]).strip()
                        name = str(df.iloc[group['name_range'][0], df.columns.get_loc(name_col)]).strip()

                        if pd.isna(name) or name == '':
                            continue

                        # Extraer estadísticas
                        stats_row = group['stats_row']
                        cols = group['cols']

                        absences = float(df.iloc[stats_row, df.columns.get_loc(cols['absences'])]) if pd.notna(df.iloc[stats_row, df.columns.get_loc(cols['absences'])]) else 0
                        late_count = float(df.iloc[stats_row, df.columns.get_loc(cols['late_count'])]) if pd.notna(df.iloc[stats_row, df.columns.get_loc(cols['late_count'])]) else 0

                        # Sumar minutos de tardanza
                        late_minutes = sum(
                            float(df.iloc[stats_row, df.columns.get_loc(col)]) 
                            if pd.notna(df.iloc[stats_row, df.columns.get_loc(col)]) else 0 
                            for col in cols['late_minutes']
                        )

                        # Sumar conteo de salidas tempranas
                        early_count = sum(
                            float(df.iloc[stats_row, df.columns.get_loc(col)]) 
                            if pd.notna(df.iloc[stats_row, df.columns.get_loc(col)]) else 0 
                            for col in cols['early_count']
                        )

                        # Obtener minutos de salida temprana
                        early_minutes = float(df.iloc[stats_row, df.columns.get_loc(cols['early_minutes'])]) if pd.notna(df.iloc[stats_row, df.columns.get_loc(cols['early_minutes'])]) else 0

                        # Calcular horas trabajadas
                        required_hours = 160  # 8 horas * 20 días
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

            except Exception as e:
                print(f"Error processing sheet {sheet}: {str(e)}")
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
            'missing_entry_days': 0,  # Placeholder for future implementation
            'missing_exit_days': 0,   # Placeholder for future implementation
            'missing_lunch_days': 0,  # Placeholder for future implementation
            'attendance_ratio': float(employee_summary['attendance_ratio'])
        }

        return stats