import pandas as pd
import numpy as np
from datetime import datetime, timedelta

class ExcelProcessor:
    def __init__(self, file):
        self.excel_file = pd.ExcelFile(file)
        self.WORK_START_TIME = datetime.strptime('7:50', '%H:%M').time()
        self.WORK_END_TIME = datetime.strptime('17:10', '%H:%M').time()

    def process_attendance_summary(self):
        """Procesa los datos de asistencia desde la hoja Summary"""
        try:
            print("Leyendo hoja Summary...")
            summary_df = pd.read_excel(self.excel_file, sheet_name="Summary")
            print("\nContenido de las primeras filas de Summary:")
            print(summary_df.head())
            print("\nColumnas en Summary:", summary_df.columns.tolist())

            empleados_df = summary_df.iloc[4:23, [0, 1, 2, 3, 4, 5, 6, 7, 8, 11, 13]]
            print("\nDatos procesados de empleados:")
            print(empleados_df.head())

            empleados_df.columns = [
                'employee_id', 'employee_name', 'department', 'required_hours',
                'actual_hours', 'late_count', 'late_minutes', 'early_departure_count',
                'early_departure_minutes', 'attendance_ratio', 'absences'
            ]

            empleados_df = empleados_df.dropna(subset=['employee_name'])
            empleados_df = empleados_df.fillna(0)

            numeric_cols = ['required_hours', 'actual_hours', 'late_count', 'late_minutes',
                          'early_departure_count', 'early_departure_minutes', 'absences']
            for col in numeric_cols:
                empleados_df[col] = pd.to_numeric(empleados_df[col], errors='coerce').fillna(0)

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
            return pd.DataFrame(columns=[
                'employee_id', 'employee_name', 'department', 'required_hours',
                'actual_hours', 'late_count', 'late_minutes', 'early_departure_count',
                'early_departure_minutes', 'attendance_ratio', 'absences'
            ])

    def count_missing_entry_days(self, employee_name):
        """Cuenta los días sin registro de entrada para un empleado"""
        try:
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]
            missing_entry_days = 0

            for sheet in attendance_sheets:
                df = pd.read_excel(self.excel_file, sheet_name=sheet)

                # Verificar en qué posición está el empleado
                name_positions = [
                    ('J', 'N'),  # Primera persona
                    ('Y', 'AC'), # Segunda persona
                    ('AN', 'AR') # Tercera persona
                ]
                entry_columns = ['B', 'Q', 'AF']  # Columnas de entrada correspondientes
                day_columns = ['A', 'P', 'AE']    # Columnas de días correspondientes

                for idx, (name_start, name_end) in enumerate(name_positions):
                    try:
                        employee_cell = str(df.iloc[2, df.columns.get_loc(name_start)]).strip()
                        if employee_cell == employee_name:
                            # Encontramos al empleado, procesar sus registros
                            entry_col = entry_columns[idx]
                            day_col = day_columns[idx]

                            # Revisar desde la fila 12 hasta encontrar una fila vacía o llegar a 42
                            for row in range(11, 42):  # Empezamos en 11 (índice 0-based para fila 12)
                                try:
                                    day_value = str(df.iloc[row, df.columns.get_loc(day_col)]).strip()
                                    if pd.isna(day_value) or day_value == '' or day_value == 'nan':
                                        break  # Fin del mes

                                    # Verificar si es un día laboral y no es ausencia
                                    if day_value.lower() != 'absence':
                                        try:
                                            day_date = datetime.strptime(str(day_value), '%Y-%m-%d')
                                            if day_date.weekday() < 5:  # 0-4 son días de semana
                                                entry_value = str(df.iloc[row, df.columns.get_loc(entry_col)]).strip()
                                                if pd.isna(entry_value) or entry_value == '' or entry_value == 'nan':
                                                    missing_entry_days += 1
                                        except:
                                            # Si no se puede convertir la fecha, asumimos que es día laboral
                                            entry_value = str(df.iloc[row, df.columns.get_loc(entry_col)]).strip()
                                            if pd.isna(entry_value) or entry_value == '' or entry_value == 'nan':
                                                missing_entry_days += 1
                                except:
                                    continue
                    except:
                        continue

            return missing_entry_days

        except Exception as e:
            print(f"Error contando días sin registro: {str(e)}")
            return 0

    def count_missing_exit_days(self, employee_name):
        """Cuenta los días sin registro de salida para un empleado"""
        try:
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]
            missing_exit_days = 0

            for sheet in attendance_sheets:
                df = pd.read_excel(self.excel_file, sheet_name=sheet)

                # Verificar en qué posición está el empleado
                name_positions = [
                    ('J', 'N'),  # Primera persona
                    ('Y', 'AC'), # Segunda persona
                    ('AN', 'AR') # Tercera persona
                ]
                exit_columns = ['I', 'X', 'AM']  # Columnas de salida correspondientes
                day_columns = ['A', 'P', 'AE']    # Columnas de días correspondientes

                for idx, (name_start, name_end) in enumerate(name_positions):
                    try:
                        employee_cell = str(df.iloc[2, df.columns.get_loc(name_start)]).strip()
                        if employee_cell == employee_name:
                            # Encontramos al empleado, procesar sus registros
                            exit_col = exit_columns[idx]
                            day_col = day_columns[idx]

                            # Revisar desde la fila 12 hasta encontrar una fila vacía o llegar a 42
                            for row in range(11, 42):  # Empezamos en 11 (índice 0-based para fila 12)
                                try:
                                    day_value = str(df.iloc[row, df.columns.get_loc(day_col)]).strip()
                                    if pd.isna(day_value) or day_value == '' or day_value == 'nan':
                                        break  # Fin del mes

                                    # Verificar si es un día laboral y no es ausencia
                                    if day_value.lower() != 'absence':
                                        try:
                                            day_date = datetime.strptime(str(day_value), '%Y-%m-%d')
                                            if day_date.weekday() < 5:  # 0-4 son días de semana
                                                exit_value = str(df.iloc[row, df.columns.get_loc(exit_col)]).strip()
                                                if pd.isna(exit_value) or exit_value == '' or exit_value == 'nan':
                                                    missing_exit_days += 1
                                        except:
                                            # Si no se puede convertir la fecha, asumimos que es día laboral
                                            exit_value = str(df.iloc[row, df.columns.get_loc(exit_col)]).strip()
                                            if pd.isna(exit_value) or exit_value == '' or exit_value == 'nan':
                                                missing_exit_days += 1
                                except:
                                    continue
                    except:
                        continue

            return missing_exit_days

        except Exception as e:
            print(f"Error contando días sin registro de salida: {str(e)}")
            return 0

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

        # Contar días sin registro de entrada y salida
        missing_entry_days = self.count_missing_entry_days(employee_name)
        missing_exit_days = self.count_missing_exit_days(employee_name)

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
            'missing_entry_days': missing_entry_days,
            'missing_exit_days': missing_exit_days,
            'missing_lunch_days': 0,  # Se implementará después
            'attendance_ratio': float(employee_summary['attendance_ratio'])
        }

        return stats