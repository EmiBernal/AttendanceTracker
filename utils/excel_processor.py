import pandas as pd
import numpy as np
from datetime import datetime, timedelta

class ExcelProcessor:
    def __init__(self, file):
        self.excel_file = pd.ExcelFile(file)
        self.WORK_START_TIME = datetime.strptime('7:50', '%H:%M').time()
        self.WORK_END_TIME = datetime.strptime('17:10', '%H:%M').time()
        self.LUNCH_TIME_LIMIT = 20  # minutos máximos permitidos para almuerzo

    def count_lunch_overtime_days(self, employee_name):
        """Cuenta los días que el empleado excedió el tiempo de almuerzo"""
        try:
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]
            lunch_overtime_days = 0

            for sheet in attendance_sheets:
                try:
                    df = pd.read_excel(self.excel_file, sheet_name=sheet)
                    print(f"\nProcesando hoja: {sheet}")

                    # Definir las columnas para cada posición de empleado
                    employee_positions = [
                        {
                            'name_col': 'J',        # Primera persona
                            'lunch_out': 'D',       # Salida almuerzo
                            'lunch_return': 'G',    # Regreso almuerzo
                            'day_col': 'A'          # Columna de días
                        },
                        {
                            'name_col': 'Y',        # Segunda persona
                            'lunch_out': 'S',       # Salida almuerzo
                            'lunch_return': 'V',    # Regreso almuerzo
                            'day_col': 'P'          # Columna de días
                        },
                        {
                            'name_col': 'AN',       # Tercera persona
                            'lunch_out': 'AH',      # Salida almuerzo
                            'lunch_return': 'AK',   # Regreso almuerzo
                            'day_col': 'AE'         # Columna de días
                        }
                    ]

                    # Buscar al empleado en cada posición posible
                    for position in employee_positions:
                        try:
                            # Verificar si el nombre coincide en la fila 3
                            name_cell = df.iloc[2, df.columns.get_loc(position['name_col'])]
                            if pd.isna(name_cell):
                                continue

                            employee_cell = str(name_cell).strip()
                            print(f"Comparando '{employee_cell}' con '{employee_name}'")

                            if employee_cell == employee_name:
                                print(f"Empleado encontrado en columna {position['name_col']}")

                                # Procesar los días desde la fila 12 hasta 42
                                for row in range(11, 42):  # índices 0-based para filas 12-42
                                    try:
                                        # Verificar el valor en la columna de días
                                        day_value = df.iloc[row, df.columns.get_loc(position['day_col'])]
                                        if pd.isna(day_value):
                                            break

                                        day_str = str(day_value).strip().lower()
                                        if day_str == '' or day_str == 'nan':
                                            break

                                        # Si no es una ausencia, verificar tiempo de almuerzo
                                        if day_str != 'absence':
                                            # Obtener horarios de almuerzo
                                            lunch_out = df.iloc[row, df.columns.get_loc(position['lunch_out'])]
                                            lunch_return = df.iloc[row, df.columns.get_loc(position['lunch_return'])]

                                            if not pd.isna(lunch_out) and not pd.isna(lunch_return):
                                                try:
                                                    # Convertir a datetime
                                                    lunch_out_time = pd.to_datetime(lunch_out).time()
                                                    lunch_return_time = pd.to_datetime(lunch_return).time()

                                                    # Calcular diferencia en minutos
                                                    lunch_minutes = (
                                                        datetime.combine(datetime.min, lunch_return_time) - 
                                                        datetime.combine(datetime.min, lunch_out_time)
                                                    ).total_seconds() / 60

                                                    if lunch_minutes > self.LUNCH_TIME_LIMIT:
                                                        lunch_overtime_days += 1
                                                        print(f"Exceso de tiempo de almuerzo encontrado en {sheet}, fila {row+1}: {lunch_minutes:.0f} minutos")
                                                except Exception as e:
                                                    print(f"Error procesando horarios en fila {row+1}: {str(e)}")
                                    except Exception as e:
                                        print(f"Error procesando fila {row+1}: {str(e)}")
                                        continue
                        except Exception as e:
                            print(f"Error verificando posición: {str(e)}")
                            continue

                except Exception as e:
                    print(f"Error procesando hoja {sheet}: {str(e)}")
                    continue

            print(f"Total días con exceso de tiempo de almuerzo: {lunch_overtime_days}")
            return lunch_overtime_days

        except Exception as e:
            print(f"Error contando días con exceso de almuerzo: {str(e)}")
            return 0

    def process_attendance_summary(self):
        """Procesa los datos de asistencia desde la hoja Summary"""
        try:
            print("Leyendo hoja Summary...")
            summary_df = pd.read_excel(self.excel_file, sheet_name="Summary", header=None)
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

            # Limpiar los datos
            empleados_df = empleados_df.dropna(subset=['employee_name'])
            empleados_df = empleados_df.fillna(0)

            # Corregir el nombre del departamento
            empleados_df['department'] = empleados_df['department'].apply(
                lambda x: 'administracion' if str(x).strip().lower() == 'administri' else str(x).strip()
            )

            # Convertir columnas numéricas
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

        # Calcular días con exceso de tiempo de almuerzo
        lunch_overtime_days = self.count_lunch_overtime_days(employee_name)

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
            'lunch_overtime_days': lunch_overtime_days,
            'attendance_ratio': float(employee_summary['attendance_ratio'])
        }

        return stats