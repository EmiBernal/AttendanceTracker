import pandas as pd
import numpy as np
from datetime import datetime, timedelta

class ExcelProcessor:
    def __init__(self, file):
        self.excel_file = pd.ExcelFile(file)
        self.validate_sheets()
        self.WORK_START_TIME = datetime.strptime('7:50', '%H:%M').time()
        self.WORK_END_TIME = datetime.strptime('17:10', '%H:%M').time()

    def validate_sheets(self):
        """Validar que existan las hojas necesarias"""
        required_sheets = ["Summary", "Shifts", "Logs", "Exceptional"]
        missing_sheets = [sheet for sheet in required_sheets if sheet not in self.excel_file.sheet_names]
        if missing_sheets:
            raise ValueError(f"Missing required sheets: {', '.join(missing_sheets)}")

    def process_attendance_summary(self):
        """Procesa la hoja de resumen de asistencia"""
        df = pd.read_excel(self.excel_file, sheet_name="Summary", header=[2, 3])

        # Normalizar nombres de columnas
        normalized_columns = ['_'.join(col).strip() for col in df.columns.values]
        df.columns = normalized_columns

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
        numeric_columns = ['required_hours', 'actual_hours', 'late_minutes', 
                         'early_departure_minutes', 'overtime_regular', 'overtime_special',
                         'late_count', 'early_departure_count', 'absences']

        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # Procesar ratio de asistencia
        if 'attendance_ratio' in df.columns:
            def convert_ratio(value):
                try:
                    if pd.isna(value) or value == '':
                        return 0.0
                    if isinstance(value, str) and '/' in value:
                        total_days, worked_days = map(float, value.split('/'))
                        return worked_days / total_days if total_days > 0 else 0.0
                    return float(value)
                except:
                    return 0.0

            df['attendance_ratio'] = df['attendance_ratio'].apply(convert_ratio)

        return df

    def process_logs(self):
        """Procesa los registros de asistencia"""
        print("Processing Logs sheet...")
        df = pd.read_excel(self.excel_file, sheet_name="Logs", skiprows=4)

        records = []
        # Procesar de a 2 filas (nombre y tiempos)
        for i in range(0, len(df), 2):
            if i + 1 >= len(df):
                break

            name_row = df.iloc[i]
            time_row = df.iloc[i + 1]

            employee_name = str(name_row.iloc[1]).strip()
            print(f"Processing employee: {employee_name}")

            # Procesar cada día (4 columnas por día)
            for day_start in range(2, len(df.columns), 4):
                try:
                    times = []
                    for j in range(4):  # Leer las 4 marcas de tiempo
                        if day_start + j < len(time_row):
                            time_str = str(time_row.iloc[day_start + j]).strip()
                            if pd.notna(time_str) and time_str != 'nan':
                                try:
                                    time_obj = datetime.strptime(time_str, '%H:%M').time()
                                    times.append(time_obj)
                                except ValueError as e:
                                    print(f"Error parsing time {time_str}: {e}")
                                    continue

                    if times:
                        record = {
                            'employee_name': employee_name,
                            'date': df.columns[day_start],
                            'first_record': times[0] if times else None,
                            'last_record': times[-1] if times else None,
                            'record_count': len(times),
                            'missing_entry': times[0] >= datetime.strptime('12:00', '%H:%M').time() if times else False,
                            'missing_exit': times[-1] <= datetime.strptime('12:00', '%H:%M').time() if times else False,
                            'early_departure': times[-1] < self.WORK_END_TIME if times else False
                        }
                        records.append(record)

                except Exception as e:
                    print(f"Error processing day for {employee_name}: {e}")
                    continue

        if not records:
            print("No records processed!")
            return pd.DataFrame(columns=[
                'employee_name', 'date', 'first_record', 'last_record',
                'record_count', 'missing_entry', 'missing_exit', 'early_departure'
            ])

        result_df = pd.DataFrame(records)
        print(f"Created DataFrame with {len(result_df)} records")
        return result_df

    def get_employee_stats(self, employee_name):
        """Obtiene estadísticas completas para un empleado"""
        summary = self.process_attendance_summary()
        print("Summary processed successfully")

        logs = self.process_logs()
        print("Logs processed successfully")

        employee_summary = summary[summary['employee_name'] == employee_name].iloc[0]
        employee_logs = logs[logs['employee_name'] == employee_name]

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
            'missing_entry_days': len(employee_logs[employee_logs['missing_entry']]),
            'missing_exit_days': len(employee_logs[employee_logs['missing_exit']]),
            'attendance_ratio': employee_summary['attendance_ratio']
        }

        return stats