import pandas as pd
import numpy as np
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
        # Leer el Excel con encabezados multinivel
        df = pd.read_excel(self.excel_file, sheet_name="Summary", header=[2, 3])
        print("Original columns:", df.columns.tolist())

        # Normalizar los nombres de las columnas
        normalized_columns = ['_'.join(col).strip() for col in df.columns.values]
        print("Normalized columns:", normalized_columns)
        df.columns = normalized_columns

        # Renombrar las columnas según el mapeo
        rename_dict = {
            'No._Unnamed: 0_level_1': 'employee_id',
            'Name_Unnamed: 1_level_1': 'employee_name',
            'Department_Unnamed: 2_level_1': 'department',
            'Work Hrs._Required': 'required_hours',
            'Work Hrs._Actual': 'actual_hours',
            'Late_Times': 'late_count',
            'Late_ Min.': 'late_minutes',
            'Early Leave_Times': 'early_departure_count',
            'Early Leave_ Min.': 'early_departure_minutes',
            'Overtime_Regular': 'overtime_regular',
            'Overtime_Special': 'overtime_special',
            'Attend (Required/Actual)_Unnamed: 11_level_1': 'attendance_ratio',
            'Absence_Unnamed: 13_level_1': 'absences'
        }

        df = df.rename(columns=rename_dict)

        # Convertir columnas numéricas
        numeric_columns = [
            'required_hours', 'actual_hours', 'late_minutes',
            'early_departure_minutes', 'overtime_regular', 'overtime_special',
            'late_count', 'early_departure_count', 'absences'
        ]

        for col in numeric_columns:
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

            name_row = df.iloc[i]
            time_row = df.iloc[i + 1]

            # El nombre está en la segunda columna (índice 1)
            employee_name = str(name_row.iloc[1]).strip()
            print(f"Processing employee: {employee_name}")

            # Procesar cada día (cada 4 columnas desde la columna C)
            for day_start in range(2, len(df.columns), 4):
                try:
                    # Recopilar los registros del día
                    day_records = []
                    for j in range(4):  # 4 registros por día
                        if day_start + j < len(time_row):
                            time_str = str(time_row.iloc[day_start + j]).strip()
                            if pd.notna(time_str) and time_str != 'nan':
                                try:
                                    time_obj = datetime.strptime(time_str, '%H:%M').time()
                                    day_records.append(time_obj)
                                except ValueError as e:
                                    print(f"Error parsing time {time_str}: {e}")
                                    continue

                    if day_records:  # Solo procesar si hay registros
                        records.append({
                            'employee_name': employee_name,
                            'date': df.columns[day_start],
                            'first_record': day_records[0],
                            'last_record': day_records[-1],
                            'record_count': len(day_records),
                            'missing_entry': (
                                day_records[0] >= self.NOON_TIME or
                                day_records[0] >= self.LATE_EVENING
                            ),
                            'missing_exit': (
                                day_records[-1] <= self.NOON_TIME and 
                                len(day_records) < 3
                            ),
                            'missing_lunch': len(day_records) <= 2,
                            'early_departure': (
                                day_records[-1] < self.WORK_END_TIME and
                                not (day_records[-1] <= self.NOON_TIME and len(day_records) < 3)
                            )
                        })

                except Exception as e:
                    print(f"Error processing day for {employee_name}: {e}")
                    continue

        # Crear DataFrame con los registros
        result_df = pd.DataFrame(records)

        if len(result_df) == 0:
            print("No records processed!")
            return pd.DataFrame(columns=[
                'employee_name', 'date', 'first_record', 'last_record',
                'record_count', 'missing_entry', 'missing_exit',
                'missing_lunch', 'early_departure'
            ])

        print(f"Created DataFrame with {len(result_df)} records")
        print(f"DataFrame columns: {result_df.columns.tolist()}")
        return result_df

    def get_employee_stats(self, employee_name):
        """Obtiene estadísticas completas para un empleado"""
        # Procesar datos
        summary = self.process_attendance_summary()
        print("Summary processed successfully")

        # Verificar si el empleado existe en el resumen
        if employee_name not in summary['employee_name'].values:
            raise ValueError(f"Employee {employee_name} not found in summary data")

        logs = self.process_logs()
        print("Logs processed successfully")

        # Obtener datos del empleado
        employee_summary = summary[summary['employee_name'] == employee_name].iloc[0]
        employee_logs = logs[logs['employee_name'] == employee_name]

        # Calcular estadísticas
        stats = {
            'name': employee_name,
            'department': employee_summary['department'],
            'required_hours': float(employee_summary['required_hours']),
            'actual_hours': float(employee_summary['actual_hours']),
            'late_days': int(employee_summary['late_count']),
            'late_minutes': float(employee_summary['late_minutes']),
            'early_departures': len(employee_logs[employee_logs['early_departure']]),
            'early_minutes': float(employee_summary['early_departure_minutes']),
            'absences': int(employee_summary['absences']),
            'missing_entry_days': len(employee_logs[employee_logs['missing_entry']]),
            'missing_exit_days': len(employee_logs[employee_logs['missing_exit']]),
            'missing_lunch_days': len(employee_logs[employee_logs['missing_lunch']]),
            'attendance_ratio': float(employee_summary.get('attendance_ratio', 0))
        }

        return stats

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