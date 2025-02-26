import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import datetime, timedelta

class ExcelProcessor:
    def __init__(self, file):
        self.excel_file = pd.ExcelFile(file)
        self.WORK_START_TIME = datetime.strptime('7:50', '%H:%M').time()
        self.WORK_END_TIME = datetime.strptime('17:10', '%H:%M').time()
        self.LUNCH_TIME_LIMIT = 20  # minutos máximos permitidos para almuerzo

        # Excepciones de horarios especiales
        self.SPECIAL_SCHEDULES = {
            'soledad silv': {
                'half_day': True,
                'end_time': datetime.strptime('12:00', '%H:%M').time(),
                'no_lunch': True,  # Nueva flag para indicar que no tiene almuerzo
                'sheet_name': '17.18',  # Hoja específica para este empleado
                'position': {
                    'name_col': 'J',
                    'entry_col': 'B',
                    'exit_col': 'D',
                    'day_col': 'A',
                    'absence_col': 'G'
                }
            },
            'agustin taba': {
                'half_day': True,
                'end_time': datetime.strptime('12:40', '%H:%M').time(),
                'no_lunch': True,  # Nueva flag para indicar que no tiene almuerzo
                'absence_cell': 'AE7',  # Celda específica para ausencias
                'sheet_name': '4.5.6'  # Hoja específica para este empleado
            },
            'valentina al': {
                'half_day': True,
                'end_time': datetime.strptime('12:40', '%H:%M').time(),
                'no_lunch': True,  # No tiene almuerzo
                'sheet_name': '7.8.9',  # Hoja específica
                'position': {
                    'name_col': 'AN',
                    'entry_col': 'AF',
                    'exit_col': 'AH',
                    'day_col': 'AE',
                    'absence_col': 'AK'
                }
            }
        }

    def is_early_departure(self, employee_name, exit_time):
        """Determina si la salida es temprana considerando excepciones"""
        if not exit_time:
            return False

        if employee_name.lower() in self.SPECIAL_SCHEDULES:
            schedule = self.SPECIAL_SCHEDULES[employee_name.lower()]
            expected_end = schedule['end_time']
        else:
            expected_end = self.WORK_END_TIME

        return exit_time < expected_end

    def count_early_departures(self, employee_name):
        """Cuenta las salidas tempranas considerando horarios especiales"""
        try:
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]
            early_departures = 0
            total_early_minutes = 0

            for sheet in attendance_sheets:
                try:
                    print(f"\nProcesando hoja: {sheet}")
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)

                    name_col_index = self.get_column_index('Y')
                    name_cell = df.iloc[2, name_col_index]
                    if pd.isna(name_cell):
                        continue

                    employee_cell = str(name_cell).strip()
                    if employee_cell == employee_name:
                        print(f"Empleado encontrado en hoja {sheet}, columna Y")

                        for row in range(11, 42):
                            try:
                                day_value = df.iloc[row, self.get_column_index('P')]
                                if pd.isna(day_value):
                                    break

                                day_str = str(day_value).strip().lower()
                                if day_str == '' or day_str == 'nan':
                                    break

                                if day_str != 'absence':
                                    exit_time = df.iloc[row, self.get_column_index('X')]
                                    if not pd.isna(exit_time):
                                        try:
                                            exit_time = pd.to_datetime(exit_time).time()
                                            if self.is_early_departure(employee_name, exit_time):
                                                early_departures += 1
                                                early_minutes = (
                                                    datetime.combine(datetime.min, self.WORK_END_TIME) -
                                                    datetime.combine(datetime.min, exit_time)
                                                ).total_seconds() / 60
                                                total_early_minutes += early_minutes
                                                print(f"Salida temprana en fila {row+1}: {early_minutes:.0f} minutos")
                                        except Exception as e:
                                            print(f"Error procesando hora de salida en fila {row+1}: {str(e)}")
                            except Exception as e:
                                print(f"Error en fila {row+1}: {str(e)}")

                except Exception as e:
                    print(f"Error procesando hoja {sheet}: {str(e)}")

            print(f"Total días con salida temprana: {early_departures}")
            print(f"Total minutos de salida temprana: {total_early_minutes:.0f}")
            return early_departures, total_early_minutes
        except Exception as e:
            print(f"Error general: {str(e)}")
            return 0, 0

    def get_column_index(self, column_letter):
        """Convierte una letra de columna de Excel a índice numérico"""
        result = 0
        for i, letter in enumerate(reversed(column_letter)):
            result += (ord(letter.upper()) - ord('A') + 1) * (26 ** i)
        return result - 1

    def count_lunch_overtime_days(self, employee_name):
        """Cuenta los días que el empleado excedió el tiempo de almuerzo"""
        # Si es Valentina, Agustín o Soledad, retornar 0 ya que no tienen almuerzo
        if employee_name.lower() in ['valentina al', 'agustin taba', 'soledad silv']:
            print(f"{employee_name} no tiene horario de almuerzo")
            return 0, 0

        try:
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]
            lunch_overtime_days = 0
            total_lunch_minutes = 0

            for sheet in attendance_sheets:
                try:
                    print(f"\nProcesando hoja: {sheet}")
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)

                    # Buscar en las tres posiciones posibles del empleado
                    employee_positions = [
                        {
                            'name_col': 'J',        # Primera persona
                            'lunch_out': 'D',       # Salida almuerzo
                            'lunch_return': 'G',    # Regreso almuerzo
                            'day_col': 'A',         # Días
                            'exit_col': 'I'         # Salida final
                        },
                        {
                            'name_col': 'Y',        # Segunda persona
                            'lunch_out': 'S',       # Salida almuerzo
                            'lunch_return': 'V',    # Regreso almuerzo
                            'day_col': 'P',         # Días
                            'exit_col': 'X'         # Salida final
                        },
                        {
                            'name_col': 'AN',       # Tercera persona
                            'lunch_out': 'AH',      # Salida almuerzo
                            'lunch_return': 'AK',   # Regreso almuerzo
                            'day_col': 'AE',        # Días
                            'exit_col': 'AM'        # Salida final
                        }
                    ]

                    for position in employee_positions:
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
                                        continue

                                    day_str = str(day_value).strip().lower()
                                    if day_str == '' or day_str == 'nan':
                                        continue

                                    if day_str != 'absence':
                                        lunch_out = df.iloc[row, cols['lunch_out']]
                                        lunch_return = df.iloc[row, cols['lunch_return']]
                                        final_exit = df.iloc[row, cols['exit_col']]

                                        # Solo procesar si existen ambos horarios de almuerzo
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
                                                    # Solo contar los minutos que exceden el límite
                                                    excess_minutes = lunch_minutes - self.LUNCH_TIME_LIMIT
                                                    total_lunch_minutes += excess_minutes
                                                    print(f"Exceso de tiempo de almuerzo en fila {row+1}: {excess_minutes:.0f} minutos (total almuerzo: {lunch_minutes:.0f} minutos)")

                                            except Exception as e:
                                                print(f"Error procesando horarios en fila {row+1}: {str(e)}")
                                        else:
                                            print(f"Fila {row+1}: No hay información completa de almuerzo")

                                except Exception as e:
                                    print(f"Error en fila {row+1}: {str(e)}")

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

                    # Definir las tres posibles posiciones del empleado
                    positions = [
                        {'name_col': 'J', 'entry_col': 'B', 'day_col': 'A'},  # Primera persona
                        {'name_col': 'Y', 'entry_col': 'Q', 'day_col': 'P'},  # Segunda persona
                        {'name_col': 'AN', 'entry_col': 'AF', 'day_col': 'AE'}  # Tercera persona
                    ]

                    for position in positions:
                        try:
                            name_col_index = self.get_column_index(position['name_col'])
                            name_cell = df.iloc[2, name_col_index]

                            if pd.isna(name_cell):
                                continue

                            employee_cell = str(name_cell).strip()
                            if employee_cell == employee_name:
                                print(f"Empleado encontrado en hoja {sheet}, columna {position['name_col']}")

                                entry_col = self.get_column_index(position['entry_col'])
                                day_col = self.get_column_index(position['day_col'])

                                for row in range(11, 42):  # Filas 12 a 42
                                    try:
                                        day_value = df.iloc[row, day_col]
                                        if pd.isna(day_value):
                                            continue

                                        day_str = str(day_value).strip().lower()
                                        if day_str == '' or day_str == 'nan':
                                            continue

                                        # Solo procesar si no es ausencia
                                        if day_str != 'absence':
                                            entry_time = df.iloc[row, entry_col]
                                            print(f"Fila {row+1}: Entrada={entry_time}")

                                            if not pd.isna(entry_time):
                                                try:
                                                    # Convertir a datetime
                                                    if isinstance(entry_time, str):
                                                        entry_time = pd.to_datetime(entry_time).time()
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
                                                        print(f"Llegada tarde en fila {row+1}: {late_minutes:.0f} minutos (hora: {entry_time})")

                                                except Exception as e:
                                                    print(f"Error procesando hora de entrada en fila {row+1}: {str(e)}")
                                            else:
                                                print(f"Sin registro de entrada en fila {row+1}")

                                    except Exception as e:
                                        print(f"Error en fila {row+1}: {str(e)}")

                        except Exception as e:
                            print(f"Error procesando posición: {str(e)}")

                except Exception as e:
                    print(f"Error procesando hoja {sheet}: {str(e)}")

            print(f"Total días de llegada tarde: {late_days}")
            print(f"Total minutos de tardanza: {total_late_minutes:.0f}")
            return late_days, total_late_minutes

        except Exception as e:
            print(f"Error general: {str(e)}")
            return 0, 0

    def count_missing_records(self, employee_name):
        """Cuenta los días sin registros de entrada, salida y almuerzo"""
        try:
            missing_entry = 0
            missing_exit = 0
            missing_lunch = 0

            # Procesamiento especial para Soledad
            if employee_name.lower() == 'soledad silv':
                try:
                    schedule = self.SPECIAL_SCHEDULES['soledad silv']
                    df = pd.read_excel(self.excel_file, sheet_name=schedule['sheet_name'], header=None)

                    # Verificar registros de entrada y salida
                    for row in range(11, 42):  # Filas 12-42
                        try:
                            pos = schedule['position']
                            day_value = df.iloc[row, self.get_column_index(pos['day_col'])]
                            if pd.isna(day_value):
                                continue

                            # Extraer el día y nombre del día
                            day_parts = str(day_value).strip().lower().split()
                            if len(day_parts) >= 2:
                                day_abbr = day_parts[1].lower()
                                # Verificar si es fin de semana
                                if day_abbr in ['sa', 'su']:
                                    print(f"Fila {row+1}: Fin de semana ({day_value}), ignorando")
                                    continue

                                # Verificar entrada
                                entry_value = df.iloc[row, self.get_column_index(pos['entry_col'])]
                                if pd.isna(entry_value) or str(entry_value).strip() == '':
                                    missing_entry += 1
                                    print(f"Falta registro de entrada en fila {row+1} ({schedule['sheet_name']})")

                                # Verificar salida
                                exit_value = df.iloc[row, self.get_column_index(pos['exit_col'])]
                                if pd.isna(exit_value) or str(exit_value).strip() == '':
                                    missing_exit += 1
                                    print(f"Falta registro de salida en fila {row+1} ({schedule['sheet_name']})")

                        except Exception as e:
                            print(f"Error en fila {row+1}: {str(e)}")
                            continue

                    # Verificar ausencias y decrementar contadores
                    for row in range(11, 42):
                        try:
                            absence_value = df.iloc[row, self.get_column_index(pos['absence_col'])]
                            if not pd.isna(absence_value) and str(absence_value).strip().lower() == 'absence':
                                missing_entry -= 1
                                missing_exit -= 1
                                print(f"Encontrado 'Absence' en fila {row+1}, decrementando contadores")
                        except Exception as e:
                            continue

                    print(f"Total días sin registro - Entrada: {missing_entry}, Salida: {missing_exit}, Almuerzo: 0 (No aplica)")
                    return missing_entry, missing_exit, 0

                except Exception as e:
                    print(f"Error procesando hoja {schedule['sheet_name']}: {str(e)}")

            # Procesamiento normal para otros empleados
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]
            for sheet in attendance_sheets:
                try:
                    print(f"\nVerificando registros en hoja {sheet}")
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)

                    positions = [
                        {
                            'name_col': 'J',
                            'entry_col': 'B',
                            'exit_col': 'I',
                            'day_col': 'A',
                            'absence_col': 'G',
                            'lunch_out': 'D',
                            'lunch_return': 'G'
                        },
                        {
                            'name_col': 'Y',
                            'entry_col': 'Q',
                            'exit_col': 'X',
                            'day_col': 'P',
                            'absence_col': 'V',
                            'lunch_out': 'S',
                            'lunch_return': 'V'
                        },
                        {
                            'name_col': 'AN',
                            'entry_col': 'AF',
                            'exit_col': 'AM',
                            'day_col': 'AE',
                            'absence_col': 'AK',
                            'lunch_out': 'AH',
                            'lunch_return': 'AK'
                        }
                    ]

                    for position in positions:
                        try:
                            name_col_index = self.get_column_index(position['name_col'])
                            name_cell = df.iloc[2, name_col_index]

                            if pd.isna(name_cell):
                                continue

                            employee_cell = str(name_cell).strip()
                            if employee_cell == employee_name:
                                print(f"Empleado encontrado en hoja {sheet}, columna {position['name_col']}")

                                entry_col = self.get_column_index(position['entry_col'])
                                exit_col = self.get_column_index(position['exit_col'])
                                day_col = self.get_column_index(position['day_col'])
                                absence_col = self.get_column_index(position['absence_col'])
                                lunch_out_col = self.get_column_index(position['lunch_out'])
                                lunch_return_col = self.get_column_index(position['lunch_return'])

                                for row in range(11, 42):  # Filas 12-42
                                    try:
                                        day_value = df.iloc[row, day_col]
                                        if pd.isna(day_value):
                                            continue

                                        try:
                                            # Extraer el día y nombre del día
                                            day_parts = str(day_value).strip().lower().split()
                                            if len(day_parts) >= 2:
                                                day_abbr = day_parts[1].lower()
                                                if day_abbr in ['sa', 'su']:
                                                    print(f"Fila {row+1}: Fin de semana ({day_value}), ignorando")
                                                    continue

                                                # Verificar ausencia
                                                absence_value = df.iloc[row, absence_col]
                                                is_absence = not pd.isna(absence_value) and str(absence_value).strip().lower() == 'absence'

                                                # Verificar entrada
                                                entry_value = df.iloc[row, entry_col]
                                                if pd.isna(entry_value) or str(entry_value).strip() == '':
                                                    missing_entry += 1
                                                    print(f"Falta registro de entrada en fila {row+1} ({sheet})")

                                                # Verificar salida solo si no es ausencia
                                                if not is_absence:
                                                    exit_value = df.iloc[row, exit_col]
                                                    if pd.isna(exit_value) or str(exit_value).strip() == '':
                                                        missing_exit += 1
                                                        print(f"Falta registro de salida en fila {row+1} ({sheet})")

                                                    # Verificar almuerzo solo si hay salida final
                                                    if not pd.isna(exit_value) and str(exit_value).strip() != '':
                                                        lunch_out = df.iloc[row, lunch_out_col]
                                                        lunch_return = df.iloc[row, lunch_return_col]

                                                        # Caso 1: Hay salida almuerzo pero no regreso
                                                        if (not pd.isna(lunch_out) and pd.isna(lunch_return)):
                                                            missing_lunch += 1
                                                            print(f"Falta registro de regreso almuerzo en fila {row+1} ({sheet})")

                                                        # Caso 2: No hay salida ni regreso almuerzo
                                                        elif (pd.isna(lunch_out) and pd.isna(lunch_return)):
                                                            missing_lunch += 1
                                                            print(f"Falta registro completo de almuerzo en fila {row+1} ({sheet})")

                                        except Exception as e:
                                            print(f"Error procesando fecha en fila {row+1}: {str(e)}")
                                            continue

                                    except Exception as e:
                                        print(f"Error en fila {row+1}: {str(e)}")
                                        continue

                                # Verificar ausencias y decrementar solo entradas
                                for row in range(11, 42):
                                    try:
                                        absence_value = df.iloc[row, absence_col]
                                        if not pd.isna(absence_value) and str(absence_value).strip().lower() == 'absence':
                                            missing_entry -= 1
                                            print(f"Encontrado 'Absence' en fila {row+1}, decrementando contador de entradas")
                                    except Exception as e:
                                        continue

                        except Exception as e:
                            print(f"Error procesando posición: {str(e)}")

                except Exception as e:
                    print(f"Error procesando hoja {sheet}: {str(e)}")

            print(f"Total días sin registro - Entrada: {missing_entry}, Salida: {missing_exit}, Almuerzo: {missing_lunch}")
            return missing_entry, missing_exit, missing_lunch

        except Exception as e:
            print(f"Error general: {str(e)}")
            return 0, 0, 0

    def calculate_agustin_hours(self, df, start_row=11, end_row=42):
        """Calcula las horas trabajadas específicamente para Agustín Tabasso"""
        try:
            total_hours = 0.0
            entry_col = self.get_column_index('AF')
            exit_col = self.get_column_index('AH')  # Cambio de AK a AH para hora de salida

            print("\nCalculando horas trabajadas para Agustín:")

            for row in range(start_row, end_row):
                try:
                    # Verificar si hay datos en la fila
                    entry_time = df.iloc[row, entry_col]
                    exit_time = df.iloc[row, exit_col]

                    # Saltar si específicamente dice "Absence" en la columna AK
                    if pd.isna(exit_time) or str(exit_time).strip().lower() == 'absence':
                        print(f"Fila {row+1}: Ausencia registrada")
                        continue

                    # Verificar si hay una entrada válida
                    if not pd.isna(entry_time):
                        try:
                            # Convertir a datetime
                            entry_time = pd.to_datetime(entry_time).time()
                            exit_time = pd.to_datetime(exit_time).time()

                            # Calcular horas trabajadas
                            hours = (datetime.combine(datetime.min, exit_time) -
                                    datetime.combine(datetime.min, entry_time)).total_seconds() / 3600

                            if hours > 0:
                                total_hours += hours
                                print(f"Fila {row+1}: {hours:.2f} horas")

                        except Exception as e:
                            print(f"Error procesando horarios en fila {row+1}: {str(e)}")
                            continue

                except Exception as e:
                    print(f"Error en fila {row+1}: {str(e)}")
                    continue

            print(f"Total horas trabajadas: {total_hours:.2f}")
            return total_hours

        except Exception as e:
            print(f"Error calculando horas: {str(e)}")
            return 0.0

    def process_attendance_summary(self):
        """Procesa los datos de asistencia desde la hoja Summary"""
        try:
            print("Leyendo hoja Summary...")
            summary_df = pd.read_excel(self.excel_file, sheet_name="Summary", header=None)

            # Special handling for Agustín's absences
            try:
                agustin_sheet = pd.read_excel(self.excel_file, sheet_name="4.5.6", header=None)
                agustin_absences = agustin_sheet.iloc[6, self.get_column_index('AE')]  # AE7 is [6, AE_index]
                if not pd.isna(agustin_absences):
                    agustin_idx = summary_df[summary_df.iloc[:, 1].str.strip() == 'agustin taba'].index
                    if len(agustin_idx) > 0:
                        summary_df.iloc[agustin_idx[0], 13] = agustin_absences
            except Exception as e:
                print(f"Error processing Agustin's special absences: {str(e)}")

            # Special handling for Valentina's absences
            try:
                valentina_sheet = pd.read_excel(self.excel_file, sheet_name="7.8.9", header=None)
                valentina_absences = self.calculate_valentina_absences(valentina_sheet)
                valentina_idx = summary_df[summary_df.iloc[:, 1].str.strip() == 'valentina al'].index
                if len(valentina_idx) > 0:
                    summary_df.iloc[valentina_idx[0], 13] = valentina_absences
            except Exception as e:
                print(f"Error processing Valentina's special absences: {str(e)}")

            # Special handling for Soledad's absences
            try:
                soledad_sheet = pd.read_excel(self.excel_file, sheet_name="17.18", header=None)
                soledad_absences = self.calculate_soledad_absences(soledad_sheet)
                soledad_idx = summary_df[summary_df.iloc[:, 1].str.strip() == 'soledad silv'].index
                if len(soledad_idx) > 0:
                    summary_df.iloc[soledad_idx[0], 13] = soledad_absences
            except Exception as e:
                print(f"Error processing Soledad's special absences: {str(e)}")

            # Extraer datos directamente de las celdas específicas (filas 5-21)
            empleados_df = summary_df.iloc[4:21, [0, 1, 2]]  # ID, Nombre, Departamento
            # Agregar horas requeridas y trabajadas de las columnas D y E
            empleados_df['required_hours'] = summary_df.iloc[4:21, 3]  # Columna D
            empleados_df['actual_hours'] = summary_df.iloc[4:21, 4]    # Columna E
            # Agregar el resto de las columnas
            empleados_df = pd.concat([
                empleados_df,
                summary_df.iloc[4:21, [5, 6, 7, 8, 13]]  # late_count hasta absences
            ], axis=1)

            print("\nDatos procesados de empleados:")
            print(empleados_df.head())

            empleados_df.columns = [
                'employee_id', 'employee_name', 'department', 'required_hours',
                'actual_hours', 'late_count', 'late_minutes', 'early_departure_count',
                'early_departure_minutes', 'absences'
            ]

            empleados_df = empleados_df.dropna(subset=['employee_name'])
            empleados_df['employee_name'] = empleados_df['employee_name'].astype(str).apply(
                lambda x: x.strip()
            )

            numeric_cols = ['required_hours', 'actual_hours', 'late_count', 'late_minutes',
                          'early_departure_count', 'early_departure_minutes']
            for col in numeric_cols:
                empleados_df[col] = pd.to_numeric(empleados_df[col], errors='coerce').fillna(0)

            # Convert absences to numeric, handling potential string values
            def parse_absence(absence_str):
                try:
                    if pd.isna(absence_str) or str(absence_str).strip() == '':
                        return 0
                    return float(absence_str)
                except:
                    return 0

            empleados_df['absences'] = empleados_df['absences'].apply(parse_absence)

            print("\nEmpleados disponibles:", empleados_df['employee_name'].tolist())
            return empleados_df

        except Exception as e:
            print(f"Error processing attendance summary: {str(e)}")
            return pd.DataFrame(columns=[
                'employee_id', 'employee_name', 'department', 'required_hours',
                'actual_hours', 'late_count', 'late_minutes', 'early_departure_count',
                'early_departure_minutes', 'absences'
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
        lunch_overtime_days, total_lunch_minutes = self.count_lunch_overtime_days(employee_name)

        # Calcular días de llegada tarde
        late_days, late_minutes = self.count_late_days(employee_name)

        # Calcular días sin registros
        missing_entry, missing_exit, missing_lunch = self.count_missing_records(employee_name)

        stats = {'name': employee_name,
            'department': str(employee_summary['department']),
            'required_hours': float(employee_summary['required_hours']),
            'actual_hours': float(employee_summary['actualhours']),
            'late_days': late_days,
            'late_minutes': late_minutes,
            'early_departures': int(employee_summary['early_departure_count']),
            'early_minutes': float(employee_summary['early_departure_minutes']),
            'absences': int(employee_summary['absences']),
            'lunch_overtime_days': lunch_overtime_days,
            'total_lunch_minutes': total_lunch_minutes,
            'missing_entry_days': missing_entry,
            'missing_exit_days': missing_exit,
            'missing_lunch_days': missing_lunch,
            'mid_day_departures': 0,  # Placeholder hasta implementar la funcionalidad
            'special_schedule': employee_name.lower() in self.SPECIAL_SCHEDULES
        }

        return stats

    def calculate_valentina_absences(self, df):
        """Calcula las ausencias de Valentina verificando solo la columna AK"""
        absences = 0
        try:
            for row in range(11, 42):  # Filas 12-42
                try:
                    absence_value = df.iloc[row, self.get_column_index('AK')]
                    if not pd.isna(absence_value) and str(absence_value).strip().lower() == 'absence':
                        absences += 1
                except Exception as e:
                    continue
        except Exception as e:
            print(f"Error calculando ausencias de Valentina: {str(e)}")
        return absences

    def calculate_soledad_absences(self, df):
        """Calcula las ausencias de Soledad verificando solo la columna G"""
        absences = 0
        try:
            for row in range(11, 42):  # Filas 12-42
                try:
                    absence_value = df.iloc[row, self.get_column_index('G')]
                    if not pd.isna(absence_value) and str(absence_value).strip().lower() == 'absence':
                        absences += 1
                except Exception as e:
                    continue
        except Exception as e:
            print(f"Error calculando ausencias de Soledad: {str(e)}")
        return absences

    def get_weekly_attendance_data(self, employee_name):
        """Calcula las estadísticas de asistencia semanal"""
        try:
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]
            weekly_stats = {}

            for sheet in attendance_sheets:
                df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)

                employee_positions = [
                    {'name_col': 'J', 'entry_col': 'B', 'exit_col': 'I', 'lunch_out': 'D', 'lunch_return': 'G', 'day_col': 'A'},
                    {'name_col': 'Y', 'entry_col': 'Q', 'exit_col': 'X', 'lunch_out': 'S', 'lunch_return': 'V', 'day_col': 'P'},
                    {'name_col': 'AN', 'entry_col': 'AF', 'exit_col': 'AM', 'lunch_out': 'AH', 'lunch_return': 'AK', 'day_col': 'AE'}
                ]

                for position in employee_positions:
                    cols = {key: self.get_column_index(value) for key, value in position.items()}
                    name_cell = df.iloc[2, cols['name_col']]

                    if pd.isna(name_cell) or str(name_cell).strip() != employee_name:
                        continue

                    for row in range(11, 42):
                        try:
                            day_value = df.iloc[row, cols['day_col']]
                            if pd.isna(day_value):
                                continue

                            try:
                                date = pd.to_datetime(day_value)
                                week_start = date - timedelta(days=date.weekday())
                                week_key = week_start.strftime('%Y-%m-%d')

                                if week_key not in weekly_stats:
                                    weekly_stats[week_key] = {
                                        'total_days': 0,
                                        'present_days': 0,
                                        'late_days': 0,
                                        'lunch_overtime_days': 0,
                                        'early_departure_days': 0,
                                        'events': []
                                    }

                                if date.weekday() < 5:  # Solo días laborables
                                    weekly_stats[week_key]['total_days'] += 1

                                    day_str = str(day_value).strip().lower()
                                    if day_str != 'absence':
                                        weekly_stats[week_key]['present_days'] += 1

                                        # Verificar llegada tarde
                                        entry_time = df.iloc[row, cols['entry_col']]
                                        if not pd.isna(entry_time):
                                            entry_time = pd.to_datetime(entry_time).time()
                                            if entry_time > self.WORK_START_TIME:
                                                weekly_stats[week_key]['late_days'] += 1
                                                weekly_stats[week_key]['events'].append(
                                                    f"{date.strftime('%d/%m')}: Llegada tarde"
                                                )

                                        # Verificar exceso en almuerzo
                                        lunch_out = df.iloc[row, cols['lunch_out']]
                                        lunch_return = df.iloc[row, cols['lunch_return']]
                                        if not pd.isna(lunch_out) and not pd.isna(lunch_return):
                                            lunch_out = pd.to_datetime(lunch_out).time()
                                            lunch_return = pd.to_datetime(lunch_return).time()
                                            lunch_minutes = (
                                                datetime.combine(datetime.min, lunch_return) -
                                                datetime.combine(datetime.min, lunch_out)
                                            ).total_seconds() / 60
                                            if lunch_minutes > self.LUNCH_TIME_LIMIT:
                                                weekly_stats[week_key]['lunch_overtime_days'] += 1
                                                weekly_stats[week_key]['events'].append(
                                                    f"{date.strftime('%d/%m')}: Exceso almuerzo"
                                                )

                                        # Verificar salida temprana
                                        exit_time = df.iloc[row, cols['exit_col']]
                                        if not pd.isna(exit_time):
                                            exit_time = pd.to_datetime(exit_time).time()
                                            if self.is_early_departure(employee_name, exit_time):
                                                weekly_stats[week_key]['early_departure_days'] += 1
                                                weekly_stats[week_key]['events'].append(
                                                    f"{date.strftime('%d/%m')}: Salida temprana"
                                                )

                            except Exception as e:
                                print(f"Error procesando fecha en fila {row+1}: {str(e)}")
                                continue

                        except Exception as e:
                            print(f"Error en fila {row+1}: {str(e)}")
                            continue

            return weekly_stats

        except Exception as e:
            print(f"Error procesando estadísticas semanales: {str(e)}")
            return {}

    def create_weekly_attendance_chart(self, employee_name):
        """Crea un gráfico de asistencia semanal"""
        weekly_stats = self.get_weekly_attendance_data(employee_name)

        weeks = list(weekly_stats.keys())
        attendance_rates = [
            (stats['present_days'] / stats['total_days']) * 100 if stats['total_days'] > 0 else 0
            for stats in weekly_stats.values()
        ]

        # Crear el gráfico base
        fig = go.Figure()

        # Agregar la línea de asistencia
        fig.add_trace(go.Scatter(
            x=weeks,
            y=attendance_rates,
            mode='lines+markers',
            name='Asistencia',            line=dict(color='#2196F3', width=3),
            marker=dict(size=8)
        ))

        # Agregar anotaciones para eventos importantes
        annotations =[]
        for week, stats in weekly_stats.items():
            if stats['events']:
                annotations.extend([
                    dict(
                        x=week,
                        y=stats['present_days'] / stats['total_days'] * 100,
                        text=event,
                        showarrow=True,
                        arrowhead=2,
                        arrowsize=1,
                        arrowwidth=2,
                        arrowcolor="#636363",
                        ax=-40,
                        ay=-40
                    )
                    for event in stats['events']
                ])

        # Configurar el diseño
        fig.update_layout(
            title=dict(
                text="Porcentaje de Asistencia Semanal",
                font=dict(size=24)
            ),
            xaxis=dict(
                title="Semana",
                tickformat="%d/%m",
                tickangle=45
            ),
            yaxis=dict(
                title="Porcentaje de Asistencia",
                range=[0, 100],
                ticksuffix="%"
            ),
            showlegend=False,
            annotations=annotations,
            height=400,
            margin=dict(t=50, r=50, b=100, l=50)
        )

        return fig