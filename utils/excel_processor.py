import pandas as pd
import numpy as np
from datetime import datetime, timedelta

class ExcelProcessor:
    def __init__(self, file):
        self.excel_file = pd.ExcelFile(file)
        self.WORK_START_TIME = datetime.strptime('7:50', '%H:%M').time()
        self.WORK_END_TIME = datetime.strptime('17:10', '%H:%M').time()

    def process_attendance_summary(self):
        """Procesa los datos de asistencia desde la hoja summary"""
        try:
            # Leer la hoja summary
            df = pd.read_excel(self.excel_file, sheet_name="summary")

            # Procesar las filas 5 a 23 de la hoja summary
            empleados_df = df.iloc[4:23, [0, 1, 2, 3, 4, 5, 6, 7, 8, 11, 13]]

            # Renombrar las columnas
            empleados_df.columns = [
                'employee_id', 'employee_name', 'department', 'required_hours',
                'actual_hours', 'late_count', 'late_minutes', 'early_departure_count',
                'early_departure_minutes', 'attendance_ratio', 'absences'
            ]

            # Limpiar los datos
            empleados_df = empleados_df.dropna(subset=['employee_name'])
            empleados_df = empleados_df.fillna(0)

            # Convertir columnas numéricas
            numeric_cols = ['required_hours', 'actual_hours', 'late_count', 'late_minutes',
                          'early_departure_count', 'early_departure_minutes', 'absences']
            for col in numeric_cols:
                empleados_df[col] = pd.to_numeric(empleados_df[col], errors='coerce').fillna(0)

            # Procesar el ratio de asistencia
            def convert_ratio(value):
                try:
                    if isinstance(value, str) and '/' in value:
                        required, actual = map(float, value.split('/'))
                        return actual / required if required > 0 else 0
                    return float(value)
                except:
                    return 0.0

            empleados_df['attendance_ratio'] = empleados_df['attendance_ratio'].apply(convert_ratio)

            return empleados_df

        except Exception as e:
            print(f"Error procesando el archivo: {str(e)}")
            # Retornar DataFrame vacío con las columnas esperadas
            return pd.DataFrame(columns=[
                'employee_id', 'employee_name', 'department', 'required_hours',
                'actual_hours', 'late_count', 'late_minutes', 'early_departure_count',
                'early_departure_minutes', 'attendance_ratio', 'absences'
            ])

    def get_employee_stats(self, employee_name):
        """Obtiene estadísticas para un empleado específico"""
        summary = self.process_attendance_summary()
        print(f"Columns in summary: {summary.columns.tolist()}")
        print(f"Number of records: {len(summary)}")
        print(f"Available employees: {summary['employee_name'].tolist()}")

        if len(summary) == 0:
            raise ValueError("No employee records found in the Excel file")

        if 'employee_name' not in summary.columns:
            raise ValueError("Column 'employee_name' not found in processed data")

        employee_data = summary[summary['employee_name'] == employee_name]

        if len(employee_data) == 0:
            raise ValueError(f"Employee '{employee_name}' not found in records")

        employee_summary = employee_data.iloc[0]

        stats = {
            'name': employee_name,
            'department': str(employee_summary['department']),
            'required_hours': float(employee_summary['required_hours']),
            'actual_hours': float(employee_summary['actual_hours']),
            'late_days': int(employee_summary['late_count']),
            'late_minutes': float(employee_summary['late_minutes']),
            'early_departures': int(employee_summary['early_departure_count']),
            'early_minutes': float(employee_summary['early_departure_minutes']),
            'absences': int(employee_summary['absences']),
            'missing_entry_days': 0,  # Placeholder para implementación futura
            'missing_exit_days': 0,   # Placeholder para implementación futura
            'missing_lunch_days': 0,  # Placeholder para implementación futura
            'attendance_ratio': float(employee_summary['attendance_ratio'])
        }

        return stats