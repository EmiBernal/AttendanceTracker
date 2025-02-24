import pandas as pd
import numpy as np
from datetime import datetime, timedelta

class ExcelProcessor:
    def __init__(self, file):
        self.excel_file = pd.ExcelFile(file)
        self.WORK_START_TIME = datetime.strptime('7:50', '%H:%M').time()
        self.WORK_END_TIME = datetime.strptime('17:10', '%H:%M').time()
        self.LUNCH_TIME_LIMIT = 20  # minutos máximos permitidos para almuerzo

    def get_column_index(self, column_letter):
        """Convierte una letra de columna de Excel a índice numérico"""
        result = 0
        for i, letter in enumerate(reversed(column_letter)):
            result += (ord(letter.upper()) - ord('A') + 1) * (26 ** i)
        return result - 1

    def count_lunch_overtime_days(self, employee_name):
        """Cuenta los días que el empleado excedió el tiempo de almuerzo"""
        try:
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]
            lunch_overtime_days = 0
            total_lunch_minutes = 0

            employee_positions = [
                {
                    'name_col': 'J',        # Primera persona
                    'lunch_out': 'D',       # Salida almuerzo
                    'lunch_return': 'G',    # Regreso almuerzo
                    'day_col': 'A',          # Columna de días
                    'entry_col': 'B',       # Entrada
                    'exit_col': 'I'         # Salida
                },
                {
                    'name_col': 'Y',        # Segunda persona
                    'lunch_out': 'S',       # Salida almuerzo
                    'lunch_return': 'V',    # Regreso almuerzo
                    'day_col': 'P',          # Columna de días
                    'entry_col': 'Q',       # Entrada
                    'exit_col': 'X'         # Salida
                },
                {
                    'name_col': 'AN',       # Tercera persona
                    'lunch_out': 'AH',      # Salida almuerzo
                    'lunch_return': 'AK',   # Regreso almuerzo
                    'day_col': 'AE',         # Columna de días
                    'entry_col': 'AF',      # Entrada
                    'exit_col': 'AM'        # Salida
                }
            ]

            for sheet in attendance_sheets:
                try:
                    print(f"\nProcesando hoja: {sheet}")
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)

                    for position in employee_positions:
                        try:
                            cols = {key: self.get_column_index(value) for key, value in position.items()}
                            name_cell = df.iloc[2, cols['name_col']]
                            if pd.isna(name_cell):
                                continue

                            employee_cell = str(name_cell).strip()
                            if employee_cell == employee_name:
                                print(f"Empleado encontrado en hoja {sheet}, columna {position['name_col']}")

                                for row in range(11, 42):
                                    try:
                                        day_value = df.iloc[row, cols['day_col']]
                                        if pd.isna(day_value):
                                            break

                                        day_str = str(day_value).strip().lower()
                                        if day_str == '' or day_str == 'nan':
                                            break

                                        if day_str != 'absence':
                                            lunch_out = df.iloc[row, cols['lunch_out']]
                                            lunch_return = df.iloc[row, cols['lunch_return']]

                                            if not pd.isna(lunch_out) and not pd.isna(lunch_return):
                                                try:
                                                    lunch_out_time = pd.to_datetime(lunch_out).time()
                                                    lunch_return_time = pd.to_datetime(lunch_return).time()

                                                    lunch_minutes = (
                                                        datetime.combine(datetime.min, lunch_return_time) -
                                                        datetime.combine(datetime.min, lunch_out_time)
                                                    ).total_seconds() / 60

                                                    if lunch_minutes > self.LUNCH_TIME_LIMIT:
                                                        lunch_overtime_days += 1
                                                        total_lunch_minutes += lunch_minutes
                                                        print(f"Exceso de tiempo de almuerzo en fila {row+1}: {lunch_minutes:.0f} minutos")

                                                except Exception as e:
                                                    print(f"Error procesando horarios en fila {row+1}: {str(e)}")
                                    except Exception as e:
                                        print(f"Error en fila {row+1}: {str(e)}")
                        except Exception as e:
                            print(f"Error procesando posición en hoja {sheet}: {str(e)}")

                except Exception as e:
                    print(f"Error procesando hoja {sheet}: {str(e)}")

            print(f"Total días con exceso de tiempo de almuerzo: {lunch_overtime_days}")
            print(f"Total minutos excedidos: {total_lunch_minutes:.0f}")
            return lunch_overtime_days, total_lunch_minutes

        except Exception as e:
            print(f"Error general: {str(e)}")
            return 0, 0

    def count_late_days(self, employee_name):
        """Cuenta los días que el empleado llegó tarde (después de 7:50)"""
        try:
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]
            late_days = 0
            total_late_minutes = 0

            for sheet in attendance_sheets:
                try:
                    print(f"\nProcesando hoja: {sheet}")
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)

                    employee_positions = [
                        {
                            'name_col': 'J',        # Primera persona
                            'entry_col': 'B',       # Entrada
                            'day_col': 'A',          # Columna de días
                            'exit_col': 'I'         # Salida
                        },
                        {
                            'name_col': 'Y',        # Segunda persona
                            'entry_col': 'Q',       # Entrada
                            'day_col': 'P',          # Columna de días
                            'exit_col': 'X'         # Salida
                        },
                        {
                            'name_col': 'AN',       # Tercera persona
                            'entry_col': 'AF',      # Entrada
                            'day_col': 'AE',         # Columna de días
                            'exit_col': 'AM'        # Salida
                        }
                    ]

                    for position in employee_positions:
                        try:
                            cols = {key: self.get_column_index(value) for key, value in position.items()}
                            name_cell = df.iloc[2, cols['name_col']]
                            if pd.isna(name_cell):
                                continue

                            employee_cell = str(name_cell).strip()
                            if employee_cell == employee_name:
                                print(f"Empleado encontrado en hoja {sheet}, columna {position['name_col']}")

                                for row in range(11, 42):
                                    try:
                                        day_value = df.iloc[row, cols['day_col']]
                                        if pd.isna(day_value):
                                            break

                                        day_str = str(day_value).strip().lower()
                                        if day_str == '' or day_str == 'nan':
                                            break

                                        # Solo procesar si no es ausencia y hay un valor en la entrada
                                        if day_str != 'absence':
                                            entry_time = df.iloc[row, cols['entry_col']]
                                            if not pd.isna(entry_time) and str(entry_time).strip() != '':
                                                try:
                                                    # Convertir a datetime asegurándose de manejar diferentes formatos
                                                    if isinstance(entry_time, str):
                                                        try:
                                                            entry_time = pd.to_datetime(entry_time).time()
                                                        except:
                                                            print(f"Error convirtiendo hora de entrada en fila {row+1}")
                                                            continue
                                                    elif isinstance(entry_time, datetime):
                                                        entry_time = entry_time.time()
                                                    else:
                                                        print(f"Formato de hora no reconocido en fila {row+1}")
                                                        continue

                                                    # Verificar si llegó tarde
                                                    if entry_time > self.WORK_START_TIME:
                                                        late_days += 1
                                                        late_minutes = (
                                                            datetime.combine(datetime.min, entry_time) -
                                                            datetime.combine(datetime.min, self.WORK_START_TIME)
                                                        ).total_seconds() / 60
                                                        total_late_minutes += late_minutes
                                                        print(f"Llegada tarde en fila {row+1}: {late_minutes:.0f} minutos")

                                                except Exception as e:
                                                    print(f"Error procesando hora de entrada en fila {row+1}: {str(e)}")
                                    except Exception as e:
                                        print(f"Error en fila {row+1}: {str(e)}")
                        except Exception as e:
                            print(f"Error procesando posición en hoja {sheet}: {str(e)}")

                except Exception as e:
                    print(f"Error procesando hoja {sheet}: {str(e)}")

            print(f"Total días de llegada tarde: {late_days}")
            print(f"Total minutos de tardanza: {total_late_minutes:.0f}")
            return late_days, total_late_minutes

        except Exception as e:
            print(f"Error general: {str(e)}")
            return 0, 0

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

            empleados_df = empleados_df.dropna(subset=['employee_name'])
            empleados_df = empleados_df.fillna(0)

            empleados_df['department'] = empleados_df['department'].apply(
                lambda x: 'administracion' if str(x).strip().lower() == 'administri' else str(x).strip()
            )

            numeric_cols = ['required_hours', 'actual_hours', 'late_count', 'late_minutes',
                          'early_departure_count', 'early_departure_minutes']
            for col in numeric_cols:
                empleados_df[col] = pd.to_numeric(empleados_df[col], errors='coerce').fillna(0)

            # Procesar ausencias excluyendo fines de semana
            def process_absences(absence_str):
                try:
                    if pd.isna(absence_str) or str(absence_str).strip() == '':
                        return 0
                    # Contar solo las ausencias en días laborales (no fines de semana)
                    absence_count = int(str(absence_str))
                    # Asumimos que las ausencias ya están contadas solo para días laborales
                    return absence_count
                except:
                    return 0

            empleados_df['absences'] = empleados_df['absences'].apply(process_absences)

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

    def count_missing_records(self, employee_name):
        """Cuenta los días sin registros de entrada, salida y almuerzo"""
        try:
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]
            missing_entry = 0
            missing_exit = 0
            missing_lunch = 0

            for sheet in attendance_sheets:
                try:
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)

                    employee_positions = [
                        {
                            'name_col': 'J',        # Primera persona
                            'entry_col': 'B',       # Entrada
                            'exit_col': 'I',        # Salida
                            'lunch_out': 'D',       # Salida almuerzo
                            'lunch_return': 'G',    # Regreso almuerzo
                            'day_col': 'A'          # Días
                        },
                        {
                            'name_col': 'Y',        # Segunda persona
                            'entry_col': 'Q',       # Entrada
                            'exit_col': 'X',        # Salida
                            'lunch_out': 'S',       # Salida almuerzo
                            'lunch_return': 'V',    # Regreso almuerzo
                            'day_col': 'P'          # Días
                        },
                        {
                            'name_col': 'AN',       # Tercera persona
                            'entry_col': 'AF',      # Entrada
                            'exit_col': 'AM',       # Salida
                            'lunch_out': 'AH',      # Salida almuerzo
                            'lunch_return': 'AK',   # Regreso almuerzo
                            'day_col': 'AE'         # Días
                        }
                    ]

                    for position in employee_positions:
                        try:
                            # Convertir letras de columnas a índices
                            cols = {key: self.get_column_index(value)
                                   for key, value in position.items()}

                            name_cell = df.iloc[2, cols['name_col']]
                            if pd.isna(name_cell):
                                continue

                            employee_cell = str(name_cell).strip()
                            if employee_cell == employee_name:
                                print(f"Verificando registros en hoja {sheet}")

                                for row in range(11, 42):
                                    try:
                                        day_value = df.iloc[row, cols['day_col']]
                                        if pd.isna(day_value):
                                            break

                                        day_str = str(day_value).strip().lower()
                                        if day_str == '' or day_str == 'nan':
                                            break

                                        # Solo verificar días no marcados como ausencia
                                        if day_str != 'absence':
                                            try:
                                                # Verificar si es día laboral
                                                day_date = datetime.strptime(str(day_value), '%Y-%m-%d')
                                                if day_date.weekday() < 5:  # Lunes a Viernes
                                                    # Verificar entrada
                                                    entry = df.iloc[row, cols['entry_col']]
                                                    if pd.isna(entry) or str(entry).strip() == '':
                                                        missing_entry += 1
                                                        print(f"Falta registro de entrada en fila {row+1}")

                                                    # Verificar salida
                                                    exit_val = df.iloc[row, cols['exit_col']]
                                                    if pd.isna(exit_val) or str(exit_val).strip() == '':
                                                        missing_exit += 1
                                                        print(f"Falta registro de salida en fila {row+1}")

                                                    # Verificar almuerzo
                                                    lunch_out = df.iloc[row, cols['lunch_out']]
                                                    lunch_return = df.iloc[row, cols['lunch_return']]
                                                    if (pd.isna(lunch_out) or str(lunch_out).strip() == '' or
                                                            pd.isna(lunch_return) or str(lunch_return).strip() == ''):
                                                        missing_lunch += 1
                                                        print(f"Falta registro de almuerzo en fila {row+1}")
                                            except:
                                                print(f"Error procesando fecha en fila {row+1}")
                                    except Exception as e:
                                        print(f"Error en fila {row+1}: {str(e)}")
                        except Exception as e:
                            print(f"Error procesando posición: {str(e)}")

                except Exception as e:
                    print(f"Error procesando hoja {sheet}: {str(e)}")

            print(f"Total días sin registro - Entrada: {missing_entry}, Salida: {missing_exit}, Almuerzo: {missing_lunch}")
            return missing_entry, missing_exit, missing_lunch

        except Exception as e:
            print(f"Error general: {str(e)}")
            return 0, 0, 0

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
        lunch_overtime_days, total_lunch_minutes = self.count_lunch_overtime_days(employee_name)

        # Calcular días de llegada tarde
        late_days, late_minutes = self.count_late_days(employee_name)

        # Calcular días sin registros
        missing_entry, missing_exit, missing_lunch = self.count_missing_records(employee_name)

        stats = {
            'name': employee_name,
            'department': str(employee_summary['department']),
            'required_hours': float(employee_summary['required_hours']),
            'actual_hours': float(employee_summary['actual_hours']),
            'late_days': late_days,
            'late_minutes': late_minutes,
            'early_departures': int(employee_summary['early_departure_count']),
            'early_minutes': float(employee_summary['early_departure_minutes']),
            'absences': int(employee_summary['absences']),
            'lunch_overtime_days': lunch_overtime_days,
            'total_lunch_minutes': total_lunch_minutes,
            'attendance_ratio': float(employee_summary['attendance_ratio']),
            'missing_entry_days': missing_entry,
            'missing_exit_days': missing_exit,
            'missing_lunch_days': missing_lunch,
            'mid_day_departures': 0  # Placeholder hasta implementar la funcionalidad
        }

        return stats