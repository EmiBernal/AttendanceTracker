import pandas as pd
import numpy as np
from difflib import get_close_matches
from datetime import datetime, timedelta

class ExcelProcessor:
    def __init__(self, file):
        self.excel_file = pd.ExcelFile(file)
        self.validate_sheets()
        self.WORK_START_TIME = datetime.strptime('7:50', '%H:%M').time()
        self.LUNCH_DURATION_LIMIT = 20  # minutos

    def validate_sheets(self):
        required_sheets = ["Summary", "Shifts", "Logs", "Exceptional"]
        missing_sheets = [
            sheet for sheet in required_sheets
            if sheet not in self.excel_file.sheet_names
        ]
        if missing_sheets:
            raise ValueError(
                f"Missing required sheets: {', '.join(missing_sheets)}")

    def process_attendance_summary(self):
        print("Reading Summary sheet...")
        df = pd.read_excel(self.excel_file, sheet_name="Summary", header=[2, 3])

        print("Original columns:", df.columns.tolist())

        # Mapear las columnas correctamente según el formato del archivo
        df.columns = [
            'employee_id' if 'No.' in str(col) else
            'employee_name' if 'Name' in str(col) else
            'department' if 'Department' in str(col) else
            'required_hours' if ('Work Hrs.' in str(col) and 'Required' in str(col)) else
            'actual_hours' if ('Work Hrs.' in str(col) and 'Actual' in str(col)) else
            'late_count' if ('Late' in str(col) and 'Times' in str(col)) else
            'late_minutes' if ('Late' in str(col) and 'Min.' in str(col)) else
            'early_departure_count' if ('Early Leave' in str(col) and 'Times' in str(col)) else
            'early_departure_minutes' if ('Early Leave' in str(col) and 'Min.' in str(col)) else
            'overtime_regular' if ('Overtime' in str(col) and 'Regular' in str(col)) else
            'overtime_special' if ('Overtime' in str(col) and 'Special' in str(col)) else
            'attendance_ratio' if 'Attend' in str(col) else
            'absences' if 'Absence' in str(col) else
            col
            for col in df.columns
        ]

        # Convertir columnas numéricas
        numeric_columns = [
            'required_hours', 'actual_hours', 'late_minutes',
            'early_departure_minutes', 'overtime_regular', 'overtime_special',
            'late_count', 'early_departure_count', 'absences'
        ]

        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # Procesar el ratio de asistencia (formato "23/12")
        if 'attendance_ratio' in df.columns:
            def parse_attendance(value):
                try:
                    if isinstance(value, str) and '/' in value:
                        actual, required = map(float, value.split('/'))
                        return actual / required if required != 0 else 0
                    return 0
                except:
                    return 0

            df['attendance_ratio'] = df['attendance_ratio'].apply(parse_attendance)

        print("Processed columns:", df.columns.tolist())
        return df

    def process_exceptional_records(self):
        """Procesa los registros excepcionales para calcular pausas de almuerzo prolongadas"""
        df = pd.read_excel(self.excel_file, sheet_name="Exceptional", header=[2, 3])

        # Mapear las columnas
        df.columns = [
            'employee_id' if 'No.' in str(col) else
            'employee_name' if 'Name' in str(col) else
            'department' if 'Department' in str(col) else
            'date' if 'Date' in str(col) else
            'am_in' if ('AM' in str(col) and 'In' in str(col)) else
            'am_out' if ('AM' in str(col) and 'Out' in str(col)) else
            'pm_in' if ('PM' in str(col) and 'In' in str(col)) else
            'pm_out' if ('PM' in str(col) and 'Out' in str(col)) else
            col
            for col in df.columns
        ]

        # Convertir columnas de tiempo a datetime
        time_columns = ['am_in', 'am_out', 'pm_in', 'pm_out']
        for col in time_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], format='%H:%M:%S', errors='coerce')

        # Calcular llegadas tarde y pausas de almuerzo prolongadas
        df['is_late'] = False
        df['late_minutes'] = 0
        df['lunch_duration'] = None
        df['extended_lunch'] = False
        df['lunch_minutes_exceeded'] = 0

        for idx, row in df.iterrows():
            # Verificar llegadas tarde
            if pd.notna(row['am_in']):
                entry_time = row['am_in'].time()
                if entry_time > self.WORK_START_TIME:
                    df.at[idx, 'is_late'] = True
                    df.at[idx, 'late_minutes'] = (entry_time.hour * 60 + entry_time.minute) - \
                                               (self.WORK_START_TIME.hour * 60 + self.WORK_START_TIME.minute)

            # Verificar almuerzos extendidos
            if pd.notna(row['am_out']) and pd.notna(row['pm_in']):
                lunch_duration = (row['pm_in'] - row['am_out']).total_seconds() / 60
                df.at[idx, 'lunch_duration'] = lunch_duration

                if lunch_duration > self.LUNCH_DURATION_LIMIT:
                    df.at[idx, 'extended_lunch'] = True
                    df.at[idx, 'lunch_minutes_exceeded'] = lunch_duration - self.LUNCH_DURATION_LIMIT

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
            'attendance_ratio': float(employee_summary['attendance_ratio'])
        }

        return stats