import pandas as pd
import numpy as np
from difflib import get_close_matches
from datetime import datetime, timedelta

class ExcelProcessor:
    def __init__(self, file):
        self.excel_file = pd.ExcelFile(file)
        self.validate_sheets()
        self.WORK_START_TIME = datetime.strptime('7:50', '%H:%M').time()
        self.WORK_END_TIME = datetime.strptime('17:10', '%H:%M').time()
        self.LUNCH_DURATION_LIMIT = 20  # minutos
        self.NOON_TIME = datetime.strptime('12:00', '%H:%M').time()
        self.LATE_EVENING = datetime.strptime('17:00', '%H:%M').time()

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
        print("Normalized columns:", ['_'.join(col).strip() for col in df.columns.values])

        # Mapear las columnas del Excel
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
        for col in ['required_hours', 'actual_hours', 'late_minutes', 
                   'early_departure_minutes', 'overtime_regular', 'overtime_special',
                   'late_count', 'early_departure_count', 'absences']:
            if col in df.columns:
                print(f"Converting column: {col}")
                try:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                except Exception as e:
                    print(f"Skipping non-Series column: {col}. {str(e)}")

        return df

    def process_logs(self):
        """Procesa la hoja de Logs para analizar registros de entrada/salida"""
        print("Processing Logs sheet...")
        df = pd.read_excel(self.excel_file, sheet_name="Logs", skiprows=4)

        records = []
        for i in range(0, len(df), 2):
            if i + 1 >= len(df):
                break

            # Obtener el nombre del empleado (está en la columna B)
            employee_name = str(df.iloc[i].iloc[1]).strip()
            print(f"Processing employee: {employee_name}")

            # Obtener la fila de tiempos
            time_row = df.iloc[i + 1]

            # Procesar cada día (cada 4 columnas)
            for j in range(2, len(df.columns), 4):
                try:
                    day_times = []
                    # Recopilar los 4 registros posibles del día
                    for k in range(4):
                        if j + k < len(time_row):
                            time_str = time_row.iloc[j + k]
                            if pd.notna(time_str):
                                try:
                                    time = pd.to_datetime(time_str).time()
                                    day_times.append(time)
                                except:
                                    continue

                    if day_times:  # Solo procesar si hay registros
                        record = {
                            'employee_name': employee_name,
                            'date': df.columns[j].strftime('%Y-%m-%d') if isinstance(df.columns[j], datetime) else None,
                            'first_record': day_times[0],
                            'last_record': day_times[-1],
                            'record_count': len(day_times)
                        }

                        # Verificar si no fichó ingreso
                        record['missing_entry'] = (
                            record['first_record'] >= self.NOON_TIME or
                            record['first_record'] >= self.LATE_EVENING
                        )

                        # Verificar si no fichó egreso
                        record['missing_exit'] = (
                            record['last_record'] <= self.NOON_TIME and 
                            record['record_count'] < 3
                        )

                        # Verificar si no fichó almuerzo
                        record['missing_lunch'] = record['record_count'] <= 2

                        # Verificar si salió antes
                        record['early_departure'] = (
                            record['last_record'] < self.WORK_END_TIME and
                            not record['missing_exit']
                        )

                        records.append(record)

                except Exception as e:
                    print(f"Error processing record for {employee_name} on column {j}: {str(e)}")
                    continue

        # Crear DataFrame con los registros procesados
        result_df = pd.DataFrame(records)
        print(f"Created DataFrame with columns: {result_df.columns.tolist()}")
        print(f"Total records processed: {len(result_df)}")
        print(f"Unique employees: {result_df['employee_name'].unique().tolist()}")
        return result_df

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
        logs = self.process_logs()

        print(f"Looking for employee: {employee_name}")
        if len(logs) > 0:
            print(f"Available employees in logs: {logs['employee_name'].unique()}")

        employee_summary = summary[summary['employee_name'] == employee_name].iloc[0]
        employee_logs = logs[logs['employee_name'] == employee_name]

        print(f"Found {len(employee_logs)} records for employee")

        # Calcular estadísticas
        missing_entry_days = len(employee_logs[employee_logs['missing_entry']])
        missing_exit_days = len(employee_logs[employee_logs['missing_exit']])
        missing_lunch_days = len(employee_logs[employee_logs['missing_lunch']])
        early_departure_days = len(employee_logs[employee_logs['early_departure']])

        stats = {
            'name': employee_name,
            'department': employee_summary['department'],
            'required_hours': float(employee_summary['required_hours']),
            'actual_hours': float(employee_summary['actual_hours']),
            'late_days': int(employee_summary['late_count']),
            'late_minutes': float(employee_summary['late_minutes']),
            'early_departures': early_departure_days,
            'early_minutes': float(employee_summary['early_departure_minutes']),
            'absences': int(employee_summary['absences']),
            'missing_entry_days': missing_entry_days,
            'missing_exit_days': missing_exit_days,
            'missing_lunch_days': missing_lunch_days,
            'attendance_ratio': float(employee_summary['attendance_ratio'])
        }

        return stats