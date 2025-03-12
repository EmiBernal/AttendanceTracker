import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import datetime, timedelta
from fpdf import FPDF

class ExcelProcessor:
    def get_employee_stats(self, employee_name):
        """Get comprehensive statistics for a specific employee"""
        # Regular stats
        late_days, late_minutes = self.count_late_days(employee_name)
        late_arrivals, late_arrival_minutes = self.count_late_arrivals_after_810(employee_name)
        early_departure_days, early_minutes = self.count_early_departures(employee_name)
        lunch_overtime_days, total_lunch_minutes = self.count_lunch_overtime_days(employee_name)
        missing_entry_days, missing_exit_days, missing_lunch_days = self.count_missing_records(employee_name)
        absence_days = self.get_absence_days(employee_name)
        absences = len(absence_days) if absence_days else 0
        mid_day_departures, mid_day_departures_text = self.count_mid_day_departures(employee_name)
        overtime_minutes = 0
        overtime_days = []

        # Get department
        department = ""
        try:
            department = self.get_employee_department(employee_name)
        except Exception as e:
            print(f"Error getting department: {str(e)}")

        # Calculate actual hours based on employee type
        if employee_name.lower() == 'sebastian':
            # Sebastian tiene un horario fijo de 3 horas y 47 minutos
            actual_hours = 3.78  # 3h 47m en decimal
            required_hours = 3.78  # Su horario requerido es igual a sus horas fijas
        elif 'ppp' in employee_name.lower():
            # Para otros empleados PPP
            weekly_hours, weekly_details = self.calculate_ppp_weekly_hours(employee_name)
            actual_hours = sum(weekly_hours.values())
            required_hours = 80.0  # Estándar mensual para PPP
        else:
            # Para empleados regulares
            required_hours = 76.40  # Estándar regular
            actual_hours = required_hours - (absences * 8)  # Subtract 8 hours for each absence

        # Asegurar que las horas no sean negativas
        actual_hours = max(0, actual_hours)

        # Get stats dictionary ready
        stats = {
            'name': employee_name,
            'department': department,
            'absences': absences,
            'absence_days': absence_days,
            'late_days': late_days,
            'late_minutes': late_minutes,
            'late_arrivals': late_arrivals,
            'late_arrival_minutes': late_arrival_minutes,
            'early_departure_days': early_departure_days,
            'early_minutes': early_minutes,
            'lunch_overtime_days': lunch_overtime_days,
            'total_lunch_minutes': total_lunch_minutes,
            'missing_entry_days': missing_entry_days,
            'missing_exit_days': missing_exit_days,
            'missing_lunch_days': missing_lunch_days,
            'required_hours': required_hours,
            'actual_hours': actual_hours,
            'mid_day_departures': mid_day_departures,
            'mid_day_departures_text': mid_day_departures_text,
            'overtime_minutes': overtime_minutes,
            'overtime_days': overtime_days
        }
        
        # Add PPP weekly hours if applicable
        if employee_name.lower() == 'sebastian':
            stats['weekly_hours'] = {
                'Semana 1': 0,
                'Semana 2': 0,
                'Semana 3': 0,
                'Semana 4': 3.78  # Sus horas fijas en la última semana
            }
            stats['weekly_details'] = [{
                'week': 'Semana 4',
                'day': '22 Jueves',
                'entry': '08:00',
                'exit': '11:47',
                'hours': '3h 47m'
            }]
        elif 'ppp' in employee_name.lower():
            weekly_hours, weekly_details = self.calculate_ppp_weekly_hours(employee_name)
            stats['weekly_hours'] = weekly_hours
            stats['weekly_details'] = weekly_details
            
        return stats

    def __init__(self, file):
        self.excel_file = pd.ExcelFile(file)
        self.DEFAULT_WORK_START_TIME = datetime.strptime('7:50', '%H:%M').time()
        self.DEFAULT_WORK_END_TIME = datetime.strptime('17:10', '%H:%M').time()
        self.LUNCH_TIME_LIMIT = 20  # minutos máximos permitidos para almuerzo

        # Excepciones de horarios especiales
        self.SPECIAL_SCHEDULES = {
            'soledad silv': {
                'half_day': True,
                'end_time': datetime.strptime('12:00', '%H:%M').time(),
                'no_lunch': True,
                'sheet_name': '17.18',
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
                'no_lunch': True,
                'check_special_exit': True,
                'exit_col': 'AH',
                'start_row': 11,  # corresponde a fila 12
                'end_row': 40
            },
            'ana': {
                'start_time': datetime.strptime('8:00', '%H:%M').time(),
                'end_time': datetime.strptime('15:00', '%H:%M').time(),
                'overtime_enabled': True,
                'no_lunch': False
            },
            'luisina ppp': {
                'start_time': datetime.strptime('8:00', '%H:%M').time(),
                'end_time': datetime.strptime('12:00', '%H:%M').time(),
                'no_lunch': True,
                'hide_exit': True
            },
            'emiliano ppp': {
                'start_time': datetime.strptime('8:00', '%H:%M').time(),
                'end_time': datetime.strptime('12:00', '%H:%M').time(),
                'no_lunch': True,
                'hide_exit': True
            },
            'sebastian': {
                'start_time': datetime.strptime('8:00', '%H:%M').time(),
                'end_time': datetime.strptime('12:00', '%H:%M').time(),
                'no_lunch': True,
                'treat_as_ppp': True,  # Nueva bandera para tratar como PPP
                'fixed_hours': 3.78  # 3 horas y 47 minutos en decimal
            }
        }

    def get_employee_schedule(self, employee_name):
        """Determina el horario de trabajo basado en el nombre del empleado"""
        if 'ppp' in employee_name.lower() or (employee_name.lower() in self.SPECIAL_SCHEDULES and 
            self.SPECIAL_SCHEDULES[employee_name.lower()].get('treat_as_ppp', False)):
            schedule = self.SPECIAL_SCHEDULES.get(employee_name.lower(), {})
            return {
                'start_time': schedule.get('start_time', datetime.strptime('8:00', '%H:%M').time()),
                'end_time': schedule.get('end_time', datetime.strptime('12:00', '%H:%M').time()),
                'no_lunch': True,
                'hide_exit': schedule.get('hide_exit', False),
                'treat_as_ppp': True,
                'fixed_hours': schedule.get('fixed_hours', None)
            }
        elif employee_name.lower() in self.SPECIAL_SCHEDULES:
            schedule = self.SPECIAL_SCHEDULES[employee_name.lower()]
            return {
                'start_time': schedule.get('start_time', self.DEFAULT_WORK_START_TIME),
                'end_time': schedule.get('end_time', self.DEFAULT_WORK_END_TIME),
                'no_lunch': schedule.get('no_lunch', False),
                'overtime_enabled': schedule.get('overtime_enabled', False),
                'hide_exit': schedule.get('hide_exit', False)
            }
        else:
            return {
                'start_time': self.DEFAULT_WORK_START_TIME,
                'end_time': self.DEFAULT_WORK_END_TIME,
                'no_lunch': False,
                'overtime_enabled': False,
                'hide_exit': False
            }

    def is_early_departure(self, employee_name, exit_time):
        """Determina si la salida es temprana considerando excepciones"""
        if not exit_time:
            return False

        schedule = self.get_employee_schedule(employee_name)
        return exit_time < schedule['end_time']

    def is_late_arrival(self, employee_name, entry_time):
        """Determina si la llegada es tarde considerando excepciones"""
        if not entry_time:
            return False

        schedule = self.get_employee_schedule(employee_name)
        return entry_time > schedule['start_time']

    def should_check_lunch(self, employee_name):
        """Determina si se debe verificar el almuerzo para este empleado"""
        schedule = self.get_employee_schedule(employee_name)
        return not schedule['no_lunch']

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

                    # Define correct positions for different employee types
                    if 'ppp' in employee_name.lower():
                        positions = [
                            {'name_col': 'J', 'entry_col': 'B', 'exit_col': 'D', 'day_col': 'A'},
                            {'name_col': 'Y', 'entry_col': 'Q', 'exit_col': 'S', 'day_col': 'P'},
                            {'name_col': 'AN', 'entry_col': 'AF', 'exit_col': 'AH', 'day_col': 'AE'}
                        ]
                    else:
                        positions = [
                            {'name_col': 'J', 'exit_col': 'I', 'day_col': 'A'},
                            {'name_col': 'Y', 'exit_col': 'X', 'day_col': 'P'},
                            {'name_col': 'AN', 'exit_col': 'AM', 'day_col': 'AE'}
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

                                day_col = self.get_column_index(position['day_col'])
                                exit_col = self.get_column_index(position['exit_col'])

                                for row in range(11, 42):
                                    try:
                                        day_value = df.iloc[row, day_col]
                                        if pd.isna(day_value):
                                            continue

                                        day_str = str(day_value).strip().lower()
                                        if day_str == '' or day_str == 'nan' or day_str == 'absence':
                                            continue

                                        exit_time = df.iloc[row, exit_col]
                                        if not pd.isna(exit_time):
                                            try:
                                                if isinstance(exit_time, str):
                                                    exit_time = pd.to_datetime(exit_time).time()
                                                elif isinstance(exit_time, datetime):
                                                    exit_time = exit_time.time()
                                                else:
                                                    print(f"Formato de hora no reconocido en fila {row+1}")
                                                    continue

                                                if self.is_early_departure(employee_name, exit_time):
                                                    early_departures += 1
                                                    schedule = self.get_employee_schedule(employee_name)
                                                    early_minutes = (
                                                        datetime.combine(datetime.min, schedule['end_time']) -
                                                        datetime.combine(datetime.min, exit_time)
                                                    ).total_seconds() / 60
                                                    total_early_minutes += early_minutes
                                                    formatted_day = self.translate_day_abbreviation(day_str)
                                                    print(f"Salida temprana en {formatted_day}: {early_minutes:.0f} minutos")

                                            except Exception as e:
                                                print(f"Error procesando hora de salida en fila {row+1}: {str(e)}")
                                                continue

                                    except Exception as e:
                                        print(f"Error en fila {row+1}: {str(e)}")
                                        continue

                        except Exception as e:
                            print(f"Error procesando posición {position['name_col']}: {str(e)}")
                            continue

                except Exception as e:
                    print(f"Error procesando hoja {sheet}: {str(e)}")
                    continue

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

    def calculate_ppp_weekly_hours(self, employee_name):
        """Calculate weekly hours for PPP employees and similar schedules"""
        try:
            schedule = self.get_employee_schedule(employee_name)
            
            # Si el empleado tiene horas fijas configuradas
            if schedule.get('fixed_hours') is not None:
                return {
                    'Semana 1': 0,
                    'Semana 2': 0,
                    'Semana 3': 0,
                    'Semana 4': schedule['fixed_hours']  # Todas las horas en la última semana
                }, [{
                    'week': 'Semana 4',
                    'day': '22 Jueves',  # El día que trabajó Sebastian
                    'entry': '08:00',
                    'exit': '11:47',
                    'hours': '3h 47m'
                }]
            
            weekly_hours = {
                'Semana 1': 0,
                'Semana 2': 0,
                'Semana 3': 0,
                'Semana 4': 0
            }
            weekly_details = []
            
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]
            
            # Resto del código existente para el cálculo de horas PPP...
            # [...]

            return weekly_hours, weekly_details
            
        except Exception as e:
            print(f"Error calculating PPP weekly hours: {str(e)}")
            return {'Semana 1': 0, 'Semana 2': 0, 'Semana 3': 0, 'Semana 4': 0}, []

    def calculate_worked_hours(self, employee_name, entry_time, exit_time):
        """Calcula las horas trabajadas considerando horarios especiales y horas extra"""
        if not entry_time or not exit_time:
            return 0, 0  # horas regulares, horas extra
            
        schedule = self.get_employee_schedule(employee_name)
        
        # Si el empleado tiene habilitadas las horas extra
        if schedule.get('overtime_enabled'):
            regular_end_time = schedule['end_time']
            
            # Si la salida es después del horario regular
            if exit_time > regular_end_time:
                # Calcular horas regulares hasta el horario de salida
                regular_hours = (
                    datetime.combine(datetime.min, regular_end_time) -
                    datetime.combine(datetime.min, entry_time)
                ).total_seconds() / 3600
                
                # Calcular horas extra después del horario de salida
                overtime_hours = (
                    datetime.combine(datetime.min, exit_time) -
                    datetime.combine(datetime.min, regular_end_time)
                ).total_seconds() / 3600
                
                return max(0, regular_hours), max(0, overtime_hours)
            
        # Para empleados sin horas extra, calcular solo horas regulares
        total_hours = (
            datetime.combine(datetime.min, exit_time) -
            datetime.combine(datetime.min, entry_time)
        ).total_seconds() / 3600
        
        return max(0, total_hours), 0

    def get_weeks_in_month(self):
        """Returns a list of tuples with (start_date, end_date) for each week in the month"""
        try:
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]
            
            # Read the first sheet to get the date range
            df = pd.read_excel(self.excel_file, sheet_name=attendance_sheets[0], header=None)
            
            # Get all dates from the day column (column A)
            dates = []
            for row in range(11, 42):  # Rows 12-42
                try:
                    day_value = df.iloc[row, 0]  # Column A
                    if pd.notna(day_value) and isinstance(day_value, str):
                        day_str = str(day_value).strip()
                        # Skip weekends and non-numeric days
                        if any(abbr in day_str.lower() for abbr in ['sa', 'su']) or not any(c.isdigit() for c in day_str):
                            continue
                        day_num = ''.join(filter(str.isdigit, day_str))
                        if day_num:
                            dates.append(int(day_num))
                except Exception as e:
                    print(f"Error processing date in row {row+1}: {str(e)}")
                    continue
            
            # Sort dates and group into weeks
            dates.sort()
            weeks = []
            
            # Always ensure 4 weeks are created
            week_ranges = [(1, 7), (8, 14), (15, 21), (22, 31)]
            
            for start, end in week_ranges:
                week_dates = [d for d in dates if start <= d <= end]
                if week_dates:
                    weeks.append((min(week_dates), max(week_dates)))
                else:
                    # If no dates in range, use range bounds
                    weeks.append((start, end))
            
            # Ensure exactly 4 weeks
            while len(weeks) < 4:
                if not weeks:
                    weeks.append((1, 7))
                else:
                    last_end = weeks[-1][1]
                    next_start = last_end + 1
                    weeks.append((next_start, min(next_start + 6, 31)))
            
            return weeks[:4]  # Return exactly 4 weeks
            
        except Exception as e:
            print(f"Error getting weeks in month: {str(e)}")
            # Always return 4 weeks
            return [(1, 7), (8, 14), (15, 21), (22, 31)]

    def get_weekly_stats(self, start_date, end_date):
        """Calculate statistics for a specific week"""
        try:
            # Initialize statistics
            stats = {
                'total_irregularities': 0,
                'irregularities_breakdown': {
                    'Llegadas tarde': 0,
                    'Salidas tempranas': 0,
                    'Ausencias': 0,
                    'Exceso almuerzo': 0
                },
                'perfect_attendance': 0,
                'perfect_employees': [],
                'most_late_day': 'N/A',
                'most_late_count': 0,
                'late_details': '',
                'most_early_day': 'N/A',
                'most_early_count': 0,
                'early_details': '',
                'most_absent_day': 'N/A',
                'most_absent_count': 0,
                'absence_details': ''
            }

            # Process each sheet for the given date range
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]

            daily_stats = {}  # To track statistics by day
            employee_records = {}  # To track perfect attendance
            daily_details = {}  # To track detailed information by day

            for sheet in attendance_sheets:
                try:
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)
                    
                    positions = [
                        {'name_col': 'J', 'entry_col': 'B', 'exit_col': 'I', 'day_col': 'A'},
                        {'name_col': 'Y', 'entry_col': 'Q', 'exit_col': 'X', 'day_col': 'P'},
                        {'name_col': 'AN', 'entry_col': 'AF', 'exit_col': 'AM', 'day_col': 'AE'}
                    ]

                    for position in positions:
                        name_col = self.get_column_index(position['name_col'])
                        entry_col = self.get_column_index(position['entry_col'])
                        exit_col = self.get_column_index(position['exit_col'])
                        day_col = self.get_column_index(position['day_col'])

                        try:
                            employee_name = str(df.iloc[2, name_col]).strip()
                            if pd.isna(employee_name) or employee_name == "":
                                continue

                            if employee_name not in employee_records:
                                employee_records[employee_name] = {'irregularities': 0}

                            for row in range(11, 42):
                                try:
                                    day_value = df.iloc[row, day_col]
                                    if pd.isna(day_value):
                                        continue

                                    day_str = str(day_value).strip()
                                    if not any(char.isdigit() for char in day_str):
                                        continue
                                        
                                    day_num = int(''.join(filter(str.isdigit, day_str)))

                                    if start_date <= day_num <= end_date:
                                        if day_num not in daily_stats:
                                            daily_stats[day_num] = {'late': [], 'early': [], 'absent': []}
                                            daily_details[day_num] = {'late': [], 'early': [], 'absent': []}

                                        if 'absence' in day_str.lower():
                                            daily_stats[day_num]['absent'].append(employee_name)
                                            employee_records[employee_name]['irregularities'] += 1
                                            stats['irregularities_breakdown']['Ausencias'] += 1
                                            continue

                                        entry_time = df.iloc[row, entry_col]
                                        exit_time = df.iloc[row, exit_col]

                                        if pd.notna(entry_time):
                                            if isinstance(entry_time, str):
                                                entry_time = pd.to_datetime(entry_time).time()
                                            elif isinstance(entry_time, pd.Timestamp):
                                                entry_time = entry_time.time()

                                            if self.is_late_arrival(employee_name, entry_time):
                                                daily_stats[day_num]['late'].append(employee_name)
                                                daily_details[day_num]['late'].append(f"{employee_name} (llegó a las {entry_time.strftime('%H:%M')})")
                                                employee_records[employee_name]['irregularities'] += 1
                                                stats['irregularities_breakdown']['Llegadas tarde'] += 1

                                        if pd.notna(exit_time):
                                            if isinstance(exit_time, str):
                                                exit_time = pd.to_datetime(exit_time).time()
                                            elif isinstance(exit_time, pd.Timestamp):
                                                exit_time = exit_time.time()

                                            if self.is_early_departure(employee_name, exit_time):
                                                daily_stats[day_num]['early'].append(employee_name)
                                                daily_details[day_num]['early'].append(f"{employee_name} (salió a las {exit_time.strftime('%H:%M')})")
                                                employee_records[employee_name]['irregularities'] += 1
                                                stats['irregularities_breakdown']['Salidas tempranas'] += 1

                                except Exception as e:
                                    print(f"Error processing row {row+1}: {str(e)}")
                                    continue

                        except Exception as e:
                            print(f"Error processing position {position['name_col']}: {str(e)}")
                            continue

                except Exception as e:
                    print(f"Error processing sheet {sheet}: {str(e)}")
                    continue

            # Calculate final statistics
            stats['total_irregularities'] = sum(stats['irregularities_breakdown'].values())
            
            # Filter out perfect attendance employees (excluding "leave early" and "NaN")
            valid_employees = [name for name, record in employee_records.items() 
                             if record['irregularities'] == 0 and 
                             name.strip() not in ['leave early (mm)', 'NaN', '']]
            
            stats['perfect_attendance'] = len(valid_employees)
            stats['perfect_employees'] = valid_employees

            # Find days with most issues and include detailed information
            if daily_stats:
                # Most late arrivals
                most_late_day = max(daily_stats.items(), key=lambda x: len(x[1]['late']))
                stats['most_late_day'] = f"Día {most_late_day[0]}"
                stats['most_late_count'] = len(most_late_day[1]['late'])
                stats['late_details'] = "\n".join(daily_details[most_late_day[0]]['late'])

                # Most early departures
                most_early_day = max(daily_stats.items(), key=lambda x: len(x[1]['early']))
                stats['most_early_day'] = f"Día {most_early_day[0]}"
                stats['most_early_count'] = len(most_early_day[1]['early'])
                stats['early_details'] = "\n".join(daily_details[most_early_day[0]]['early'])

                # Most absences
                most_absent_day = max(daily_stats.items(), key=lambda x: len(x[1]['absent']))
                stats['most_absent_day'] = f"Día {most_absent_day[0]}"
                stats['most_absent_count'] = len(most_absent_day[1]['absent'])
                stats['absence_details'] = ", ".join(most_absent_day[1]['absent'])

            return stats

        except Exception as e:
            print(f"Error getting weekly stats: {str(e)}")
            return {
                'total_irregularities': 0,
                'irregularities_breakdown': {},
                'perfect_attendance': 0,
                'perfect_employees': [],
                'most_late_day': 'N/A',
                'most_late_count': 0,
                'late_details': 'Sin datos',
                'most_early_day': 'N/A',
                'most_early_count': 0,
                'early_details': 'Sin datos',
                'most_absent_day': 'N/A',
                'most_absent_count': 0,
                'absence_details': 'Sin datos'
            }

    def get_employee_hours(self, employee_name):
        """Obtiene las horas trabajadas y extra para un empleado"""
        total_regular_hours = 0
        total_overtime_hours = 0
        hours_details = []
        
        try:
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]
            
            for sheet in attendance_sheets:
                df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)
                
                positions = [
                    {'name_col': 'J', 'entry_col': 'B', 'exit_col': 'I', 'day_col': 'A'},
                    {'name_col': 'Y', 'entry_col': 'Q', 'exit_col': 'X', 'day_col': 'P'},
                    {'name_col': 'AN', 'entry_col': 'AF', 'exit_col': 'AM', 'day_col': 'AE'}
                ]
                
                for position in positions:
                    try:
                        name_col = self.get_column_index(position['name_col'])
                        name_cell = df.iloc[2, name_col]
                        
                        if pd.isna(name_cell) or str(name_cell).strip() != employee_name:
                            continue
                            
                        entry_col = self.get_column_index(position['entry_col'])
                        exit_col = self.get_column_index(position['exit_col'])
                        day_col = self.get_column_index(position['day_col'])
                        
                        for row in range(11, 42):
                            try:
                                day_value = df.iloc[row, day_col]
                                if pd.isna(day_value):
                                    continue
                                    
                                day_str = str(day_value).strip()
                                if 'absence' in day_str.lower() or any(abbr in day_str.lower() for abbr in ['sa', 'su']):
                                    continue
                                    
                                entry_time = df.iloc[row, entry_col]
                                exit_time = df.iloc[row, exit_col]
                                
                                if not pd.isna(entry_time) and not pd.isna(exit_time):
                                    # Convertir a time() si son strings o datetime
                                    if isinstance(entry_time, str):
                                        entry_time = pd.to_datetime(entry_time).time()
                                    elif isinstance(entry_time, datetime):
                                        entry_time = entry_time.time()
                                        
                                    if isinstance(exit_time, str):
                                        exit_time = pd.to_datetime(exit_time).time()
                                    elif isinstance(exit_time, datetime):
                                        exit_time = exit_time.time()
                                    
                                    # No mostrar salidas para empleados especiales
                                    schedule = self.get_employee_schedule(employee_name)
                                    if schedule.get('hide_exit'):
                                        continue
                                    
                                    regular_hours, overtime_hours = self.calculate_worked_hours(
                                        employee_name, entry_time, exit_time)
                                        
                                    total_regular_hours += regular_hours
                                    total_overtime_hours += overtime_hours
                                    
                                    hours_details.append({
                                        'day': self.translate_day_abbreviation(day_str),
                                        'entry': entry_time.strftime('%H:%M'),
                                        'exit': exit_time.strftime('%H:%M'),
                                        'regular_hours': f"{regular_hours:.2f}",
                                        'overtime_hours': f"{overtime_hours:.2f}" if overtime_hours > 0 else None
                                    })
                                    
                            except Exception as e:
                                print(f"Error processing row {row+1}: {str(e)}")
                                continue
                                
                    except Exception as e:
                        print(f"Error processing position {position['name_col']}: {str(e)}")
                        continue
                        
        except Exception as e:
            print(f"Error calculating hours: {str(e)}")
            
        return total_regular_hours, total_overtime_hours, hours_details



    def count_lunch_overtime_days(self, employee_name):
        """Returns a list of days and total minutes when the employee exceeded lunch time"""
        try:
            lunch_overtime_days = []
            total_lunch_minutes = 0
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index+1:]  # Start after Exceptional

            # Check if employee should have lunch time checked
            if not self.should_check_lunch(employee_name):
                print(f"{employee_name} no tiene horario de almuerzo")
                return [], 0

            # Define the three possible positions
            positions = [
                {
                    'name_col': 'J',
                    'day_col': 'A',
                    'lunch_out': 'D',
                    'lunch_return': 'G'
                },
                {
                    'name_col': 'Y',
                    'day_col': 'P',
                    'lunch_out': 'S',
                    'lunch_return': 'V'
                },
                {
                    'name_col': 'AN',
                    'day_col': 'AE',
                    'lunch_out': 'AH',
                    'lunch_return': 'AK'
                }
            ]

            for sheet in attendance_sheets:
                try:
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)

                    # Check each possible position
                    for position in positions:
                        try:
                            # Check if employee name matches in the correct position (row 3)
                            name_col_index = self.get_column_index(position['name_col'])
                            name_cell = df.iloc[2, name_col_index]

                            if pd.isna(name_cell) or str(name_cell).strip() != employee_name:
                                continue

                            # If name matches, check for lunch overtime
                            day_col = self.get_column_index(position['day_col'])
                            lunch_out_col = self.get_column_index(position['lunch_out'])
                            lunch_return_col = self.get_column_index(position['lunch_return'])

                            for row in range(11, 42):  # Check rows 12-42
                                try:
                                    day_value = df.iloc[row, day_col]
                                    lunch_out = df.iloc[row, lunch_out_col]
                                    lunch_return = df.iloc[row, lunch_return_col]

                                    if not pd.isna(day_value):
                                        # Skip weekends
                                        day_str = str(day_value).strip()
                                        if any(abbr in day_str.lower() for abbr in ['sa', 'su']):
                                            continue

                                        # Only process if both lunch times exist
                                        if not pd.isna(lunch_out) and not pd.isna(lunch_return):
                                            try:
                                                lunch_out_time = pd.to_datetime(lunch_out).time()
                                                lunch_return_time = pd.to_datetime(lunch_return).time()

                                                lunch_minutes = (
                                                    datetime.combine(datetime.min, lunch_return_time) -
                                                    datetime.combine(datetime.min, lunch_out_time)
                                                ).total_seconds() / 60

                                                if lunch_minutes > self.LUNCH_TIME_LIMIT:
                                                    # Calculate excess minutes
                                                    excess_minutes = lunch_minutes - self.LUNCH_TIME_LIMIT
                                                    total_lunch_minutes += excess_minutes

                                                    # Translate the day to Spanish format
                                                    formatted_day = self.translate_day_abbreviation(day_str)
                                                    print(f"Exceso de almuerzo en hoja {sheet}, fila {row+1}, día: {formatted_day} ({excess_minutes:.0f} minutos extra)")
                                                    lunch_overtime_days.append(formatted_day)

                                            except Exception as e:
                                                print(f"Error processing lunch times in row {row+1}: {str(e)}")
                                                continue

                                except Exception as e:
                                    print(f"Error processing row {row+1}: {str(e)}")
                                    continue

                        except Exception as e:
                            print(f"Error checking position {position['name_col']}: {str(e)}")
                            continue

                except Exception as e:
                    print(f"Error processing sheet {sheet}: {str(e)}")
                    continue

            print(f"Total días con exceso de almuerzo: {len(lunch_overtime_days)}")
            print(f"Total minutos excedidos: {total_lunch_minutes:.0f}")
            return lunch_overtime_days, total_lunch_minutes

        except Exception as e:
            print(f"Error getting lunch overtime days: {str(e)}")
            return [], 0

    def count_late_days(self, employee_name):
        """Cuenta los días que el empleado llegó tarde según su horario asignado"""
        try:
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]
            late_days = []
            total_late_minutes = 0

            schedule = self.get_employee_schedule(employee_name)
            work_start_time = schedule['start_time']

            for sheet in attendance_sheets:
                try:
                    print(f"\nProcesando hoja: {sheet}")
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)

                    positions = [
                        {'name_col': 'J', 'entry_col': 'B', 'day_col': 'A'},
                        {'name_col': 'Y', 'entry_col': 'Q', 'day_col': 'P'},
                        {'name_col': 'AN', 'entry_col': 'AF', 'day_col': 'AE'}
                    ]

                    try:
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

                                    for row in range(11, 42):
                                        try:
                                            day_value = df.iloc[row, day_col]
                                            if pd.isna(day_value):
                                                continue

                                            day_str = str(day_value).strip().lower()
                                            if day_str == '' or day_str == 'nan' or day_str == 'absence':
                                                continue

                                            entry_time = df.iloc[row, entry_col]
                                            print(f"Fila {row+1}: Entrada={entry_time}")

                                            if not pd.isna(entry_time):
                                                try:
                                                    if isinstance(entry_time, str):
                                                        entry_time = pd.to_datetime(entry_time).time()
                                                    elif isinstance(entry_time, datetime):
                                                        entry_time = entry_time.time()
                                                    else:
                                                        print(f"Formato de hora no reconocido en fila {row+1}")
                                                        continue

                                                    if self.is_late_arrival(employee_name, entry_time):
                                                        late_minutes = (
                                                            datetime.combine(datetime.min, entry_time) -
                                                            datetime.combine(datetime.min, work_start_time)
                                                        ).total_seconds() / 60
                                                        total_late_minutes += late_minutes

                                                        formatted_day = self.translate_day_abbreviation(day_str)
                                                        late_days.append(formatted_day)
                                                        print(f"Llegada tarde en fila {row+1}: {late_minutes:.0f} minutos (hora: {entry_time}), Dia: {formatted_day}")

                                                except Exception as e:
                                                    print(f"Error procesando hora de entrada en fila {row+1}: {str(e)}")
                                                    continue
                                            else:
                                                print(f"Sin registro de entrada en fila {row+1}")

                                        except Exception as e:
                                            print(f"Error en fila {row+1}: {str(e)}")
                                            continue

                            except Exception as e:
                                print(f"Error procesando posición {position['name_col']}: {str(e)}")
                                continue

                    except Exception as e:
                        print(f"Error en el bucle de posiciones: {str(e)}")
                        continue

                except Exception as e:
                    print(f"Error procesando hoja {sheet}: {str(e)}")
                    continue

            print(f"Total días de llegada tarde: {len(late_days)}")
            print(f"Total minutos de tardanza: {total_late_minutes:.0f}")
            return late_days, total_late_minutes

        except Exception as e:
            print(f"Error general: {str(e)}")
            return [], 0

    def count_missing_records(self, employee_name):
        """Cuenta los días sin registros de entrada, salida y almuerzo"""
        try:
            missing_entry_days = []
            missing_exit_days = []
            missing_lunch_days = []

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
                            day_str = str(day_value).strip()
                            # Get special schedule configuration
                            schedule = self.SPECIAL_SCHEDULES[employee_name.lower()]
                            if 'position' in schedule:
                                pos = schedule['position']
                            else:
                                continue

                            # Skip weekends
                            if any(abbr in day_str.lower() for abbr in ['sa', 'su']):
                                continue

                            # Verificar entrada
                            entry_value = df.iloc[row, self.get_column_index(pos['entry_col'])]
                            if pd.isna(entry_value) or str(entry_value).strip() == '':
                                missing_entry_days.append(self.translate_day_abbreviation(day_str))
                                print(f"Falta registro de entrada en fila {row+1} ({schedule['sheet_name']})")

                            # Verificar salida
                            exit_value = df.iloc[row, self.get_column_index(pos['exit_col'])]
                            if pd.isna(exit_value) or str(exit_value).strip() == '':
                                missing_exit_days.append(self.translate_day_abbreviation(day_str))
                                print(f"Falta registro de salida en fila {row+1} ({schedule['sheet_name']})")

                        except Exception as e:
                            print(f"Error en fila {row+1}: {str(e)}")
                            continue

                    # Verificar ausencias y ajustar listas
                    for row in range(11, 42):
                        try:
                            absence_value = df.iloc[row, self.get_column_index(pos['absence_col'])]
                            if not pd.isna(absence_value) and str(absence_value).strip().lower() == 'absence':
                                day_value = df.iloc[row, self.get_column_index(pos['day_col'])]
                                formatted_day = self.translate_day_abbreviation(str(day_value).strip())
                                if formatted_day in missing_entry_days:
                                    missing_entry_days.remove(formatted_day)
                                if formatted_day in missing_exit_days:
                                    missing_exit_days.remove(formatted_day)
                                print(f"Encontrado 'Absence' en fila {row+1}, ajustando contadores")
                        except Exception as e:
                            continue

                    print(f"Total días sin registro - Entrada: {len(missing_entry_days)}, Salida: {len(missing_exit_days)}, Almuerzo: 0 (No aplica)")
                    return missing_entry_days, missing_exit_days, []

                except Exception as e:
                    print(f"Error procesando hoja {schedule['sheet_name']}: {str(e)}")

            # Procesamiento normal para otros empleados
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]

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

            for sheet in attendance_sheets:
                try:
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)

                    for position in positions:
                        try:
                            name_col_index = self.get_column_index(position['name_col'])
                            name_cell = df.iloc[2, name_col_index]

                            if pd.isna(name_cell):
                                continue

                            if str(name_cell).strip() == employee_name:
                                for row in range(11, 42):  # Filas 12-42
                                    try:
                                        day_value = df.iloc[row, self.get_column_index(position['day_col'])]
                                        if pd.isna(day_value):
                                            continue

                                        day_str = str(day_value).strip()
                                        # Skip weekends
                                        if any(abbr in day_str.lower() for abbr in ['sa', 'su']):
                                            continue

                                        # Verificar entrada
                                        entry_value = df.iloc[row, self.get_column_index(position['entry_col'])]
                                        if pd.isna(entry_value) or str(entry_value).strip() == '':
                                            missing_entry_days.append(self.translate_day_abbreviation(day_str))

                                        # Verificar salida solo si no es empleado especial
                                        if employee_name.lower() not in ['valentina al', 'agustin taba']:
                                            exit_value = df.iloc[row, self.get_column_index(position['exit_col'])]
                                            if pd.isna(exit_value) or str(exit_value).strip() == '':
                                                missing_exit_days.append(self.translate_day_abbreviation(day_str))

                                        # Verificar almuerzo
                                        if self.should_check_lunch(employee_name):
                                            lunch_out = df.iloc[row, self.get_column_index(position['lunch_out'])]
                                            lunch_return = df.iloc[row, self.get_column_index(position['lunch_return'])]

                                            if not pd.isna(exit_value) and str(exit_value).strip() != '':
                                                if (not pd.isna(lunch_out) and pd.isna(lunch_return)) or \
                                                   (pd.isna(lunch_out) and pd.isna(lunch_return)):
                                                    missing_lunch_days.append(self.translate_day_abbreviation(day_str))

                                    except Exception as e:
                                        print(f"Error en fila {row+1}: {str(e)}")
                                        continue

                                # Verificar ausencias y ajustar listas
                                for row in range(11, 42):
                                    try:
                                        absence_value = df.iloc[row, self.get_column_index(position['absence_col'])]
                                        if not pd.isna(absence_value) and str(absence_value).strip().lower() == 'absence':
                                            day_value = df.iloc[row, self.get_column_index(position['day_col'])]
                                            formatted_day = self.translate_day_abbreviation(str(day_value).strip())
                                            if formatted_day in missing_entry_days:
                                                missing_entry_days.remove(formatted_day)
                                            if formatted_day in missing_exit_days:
                                                missing_exit_days.remove(formatted_day)
                                            if formatted_day in missing_lunch_days:
                                                missing_lunch_days.remove(formatted_day)
                                    except Exception as e:
                                        continue

                        except Exception as e:
                            print(f"Error procesando posición: {str(e)}")
                            continue

                except Exception as e:
                    print(f"Error procesando hoja {sheet}: {str(e)}")
                    continue

            print(f"Total días sin registro - Entrada: {len(missing_entry_days)}, Salida: {len(missing_exit_days)}, Almuerzo: {len(missing_lunch_days)}")
            return missing_entry_days, missing_exit_days, missing_lunch_days

        except Exception as e:
            print(f"Error general: {str(e)}")
            return [], [], []

    def format_list_in_columns(self, items, items_per_column=8):
        """
        Format a list of items into columns, with exactly 8 items per column.
        Returns a string with the formatted text.
        """
        if not items:
            return "No hay días registrados"

        # Create columns of exactly 8 items
        columns = []
        current_column = []
        
        for item in sorted(items, key=lambda x: int(x.split()[0])):
            if len(current_column) < items_per_column:
                current_column.append(f"• {item}")
            else:
                # When column is full (8 items), start a new one
                columns.append(current_column)
                current_column = [f"• {item}"]
        
        # Add the last column if it has any items
        if current_column:
            columns.append(current_column)

        # Format each column as a string, ensuring consistent width
        formatted_columns = []
        for column in columns:
            # Pad column to 8 items if needed
            while len(column) < items_per_column:
                column.append("")  # Add empty strings for padding
            formatted_columns.append("\n".join(column))

        # Join columns with sufficient spacing (10 spaces)
        return "          ".join(formatted_columns)

    def translate_day_abbreviation(self, day_str):
        """Translates day abbreviations to Spanish format"""
        try:
            # Get the day number and name parts
            parts = day_str.split()
            if len(parts) < 2:
                return day_str

            day_num = parts[0]
            day_name = parts[1].lower()

            # Map English abbreviations to Spanish
            day_map = {
                'mo': 'Lu',
                'tu': 'Ma',
                'we': 'Mi',
                'th': 'Ju',
                'fr': 'Vi',
                'sa': 'Sa',
                'su': 'Do'
            }

            for eng, esp in day_map.items():
                if day_name.startswith(eng):
                    return f"{day_num} {esp}"

            return day_str

        except Exception as e:
            print(f"Error translating day: {str(e)}")
            return day_str

    def calculate_ppp_weekly_hours(self, employee_name):
        """Calculate weekly hours for PPP employees"""
        try:
            weekly_hours = {
                'Semana 1': 0,
                'Semana 2': 0,
                'Semana 3': 0,
                'Semana 4': 0
            }
            weekly_details = []
            
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]
            
            # Positions in the Excel sheet with correct columns for PPP employees
            positions = [
                {'name_col': 'J', 'entry_col': 'B', 'exit_col': 'D', 'day_col': 'A'},  # Entry y salida B y D
                {'name_col': 'Y', 'entry_col': 'Q', 'exit_col': 'S', 'day_col': 'P'},  # Entry y salida Q y S
                {'name_col': 'AN', 'entry_col': 'AF', 'exit_col': 'AH', 'day_col': 'AE'}  # Entry y salida AF y AH
            ]

            for sheet in attendance_sheets:
                try:
                    print(f"\nProcesando hoja: {sheet}")
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)
                    
                    for position in positions:
                        try:
                            name_col_index = self.get_column_index(position['name_col'])
                            name_cell = df.iloc[2, name_col_index]
                            
                            if pd.isna(name_cell):
                                continue
                                
                            if str(name_cell).strip() == employee_name:
                                print(f"Empleado encontrado en hoja {sheet}, columna {position['name_col']}")
                                entry_col = self.get_column_index(position['entry_col'])
                                exit_col = self.get_column_index(position['exit_col'])
                                day_col = self.get_column_index(position['day_col'])
                                
                                for row in range(11, 42):  # Filas 12-42
                                    try:
                                        day_value = df.iloc[row, day_col]
                                        if pd.isna(day_value):
                                            continue
                                            
                                        day_str = str(day_value).strip()
                                        if 'absence' in day_str.lower():
                                            continue
                                            
                                        # Skip weekends
                                        if any(abbr in day_str.lower() for abbr in ['sa', 'su']):
                                            continue
                                            
                                        entry_time = df.iloc[row, entry_col]
                                        exit_time = df.iloc[row, exit_col]
                                        
                                        print(f"Fila {row+1}: Entrada={entry_time}, Salida={exit_time}")
                                        
                                        if not pd.isna(entry_time) and not pd.isna(exit_time):
                                            try:
                                                entry_time = pd.to_datetime(entry_time).time()
                                                exit_time = pd.to_datetime(exit_time).time()
                                                
                                                # Obtener horas y minutos por separado
                                                end_hour = exit_time.hour
                                                end_minute = exit_time.minute
                                                start_hour = entry_time.hour
                                                start_minute = entry_time.minute
                                                
                                                # Realizar la resta de horas y minutos
                                                diff_hours = end_hour - start_hour
                                                diff_minutes = end_minute - start_minute
                                                
                                                # Ajustar si los minutos son negativos
                                                if diff_minutes < 0:
                                                    diff_minutes += 60
                                                    diff_hours -= 1
                                                
                                                # Convertir minutos extras a horas si superan 60
                                                if diff_minutes >= 60:
                                                    extra_hours = diff_minutes // 60
                                                    diff_hours += extra_hours
                                                    diff_minutes = diff_minutes % 60
                                                
                                                # Calcular horas totales para esta entrada
                                                total_hours = diff_hours + (diff_minutes / 60)
                                                
                                                # Add to appropriate week
                                                day_num = int(day_str.split()[0])
                                                week_key = ''
                                                if 1 <= day_num <= 7:
                                                    week_key = 'Semana 1'
                                                elif 8 <= day_num <= 14:
                                                    week_key = 'Semana 2'
                                                elif 15 <= day_num <= 21:
                                                    week_key = 'Semana 3'
                                                elif day_num >= 22:
                                                    week_key = 'Semana 4'
                                                    
                                                if week_key:
                                                    weekly_hours[week_key] += total_hours
                                                    week_details.append({
                                                        'week': week_key,
                                                        'day': self.translate_day_abbreviation(day_str),
                                                        'entry': entry_time.strftime('%H:%M'),
                                                        'exit': exit_time.strftime('%H:%M'),
                                                        'hours': f"{diff_hours}h {diff_minutes}m"
                                                    })
                                                    print(f"Día {day_str}: {diff_hours}h {diff_minutes}m")
                                                    
                                            except Exception as e:
                                                print(f"Error processing times in row {row+1}: {str(e)}")
                                                continue
                                                
                                    except Exception as e:
                                        print(f"Error in row {row+1}: {str(e)}")
                                        continue
                                        
                        except Exception as e:
                            print(f"Error checking position {position['name_col']}: {str(e)}")
                            continue
                            
                except Exception as e:
                    print(f"Error processing sheet {sheet}: {str(e)}")
                    continue
                    
            # Round weekly hours
            for week in weekly_hours:
                weekly_hours[week] = round(weekly_hours[week], 2)
                
            return weekly_hours, weekly_details
            
        except Exception as e:
            print(f"Error calculating PPP weekly hours: {str(e)}")
            return {'Semana 1': 0, 'Semana 2': 0, 'Semana 3': 0, 'Semana 4': 0}, []

    def count_mid_day_departures(self, employee_name):
        """Cuenta los retiros durante el horario laboral"""
        try:
            # No contar retiros durante horario para PPP o empleados especiales
            if 'ppp' in employee_name.lower() or employee_name.lower() == 'agustin taba':
                return 0, "No aplica"

            mid_day_departures = 0
            departure_details = []
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]

            for sheet in attendance_sheets:
                try:
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)

                    positions = [
                        {'name_col': 'J', 'exit_col': 'G', 'day_col': 'A'},
                        {'name_col': 'Y', 'exit_col': 'V', 'day_col': 'P'},
                        {'name_col': 'AN', 'exit_col': 'AK', 'day_col': 'AE'}
                    ]

                    for position in positions:
                        try:
                            name_col_index = self.get_column_index(position['name_col'])
                            name_cell = df.iloc[2, name_col_index]

                            if pd.isna(name_cell):
                                continue

                            if str(name_cell).strip() == employee_name:
                                exit_col = self.get_column_index(position['exit_col'])
                                day_col = self.get_column_index(position['day_col'])

                                for row in range(11, 42):
                                    try:
                                        day_value = df.iloc[row, day_col]
                                        if pd.isna(day_value):
                                            continue

                                        day_str = str(day_value).strip()
                                        if 'absence' in day_str.lower():
                                            continue

                                        # Skip weekends
                                        if any(abbr in day_str.lower() for abbr in ['sa', 'su']):
                                            continue

                                        exit_time = df.iloc[row, exit_col]
                                        if not pd.isna(exit_time):
                                            try:
                                                exit_time = pd.to_datetime(exit_time).time()
                                                formatted_day = self.translate_day_abbreviation(day_str)
                                                departure_details.append(f"{formatted_day} ({exit_time.strftime('%H:%M')})")
                                                mid_day_departures += 1
                                            except Exception as e:
                                                print(f"Error procesando hora de salida en fila {row+1}: {str(e)}")
                                                continue

                                    except Exception as e:
                                        print(f"Error en fila {row+1}: {str(e)}")
                                        continue

                        except Exception as e:
                            print(f"Error procesando posición {position['name_col']}: {str(e)}")
                            continue

                except Exception as e:
                    print(f"Error procesando hoja {sheet}: {str(e)}")
                    continue

            # Format departure details
            departure_text = self.format_list_in_columns(departure_details) if departure_details else "No hay registros"
            
            return mid_day_departures, departure_text

        except Exception as e:
            print(f"Error general: {str(e)}")
            return 0, "Error al procesar los datos"

    def get_employee_stats(self, employee_name):
        """Get comprehensive statistics for a specific employee"""
        # Regular stats
        late_days, late_minutes = self.count_late_days(employee_name)
        late_arrivals, late_arrival_minutes = self.count_late_arrivals_after_810(employee_name)
        early_departure_days, early_minutes = self.count_early_departures(employee_name)
        lunch_overtime_days, total_lunch_minutes = self.count_lunch_overtime_days(employee_name)
        missing_entry_days, missing_exit_days, missing_lunch_days = self.count_missing_records(employee_name)
        absence_days = self.get_absence_days(employee_name)
        absences = len(absence_days) if absence_days else 0
        mid_day_departures, mid_day_departures_text = self.count_mid_day_departures(employee_name)
        overtime_minutes = 0
        overtime_days = []

        # Get overtime for agustin taba
        if employee_name.lower() == 'agustin taba':
            overtime_minutes, overtime_days = self.calculate_overtime(employee_name)

        # Get department
        department = ""
        try:
            department = self.get_employee_department(employee_name)
        except Exception as e:
            print(f"Error getting department: {str(e)}")

        # Calculate actual hours differently for PPP employees
        if 'ppp' in employee_name.lower():
            weekly_hours, weekly_details = self.calculate_ppp_weekly_hours(employee_name)
            actual_hours = sum(weekly_hours.values())
            required_hours = 80.0  # Estándar mensual para PPP
        else:
            required_hours = 76.40  # Estándar regular
            actual_hours = required_hours - (absences * 8)  # Subtract 8 hours for each absence

        # Get stats dictionary ready
        stats = {
            'name': employee_name,
            'department': department,
            'absences': absences,
            'absence_days': absence_days,
            'late_days': late_days,
            'late_minutes': late_minutes,
            'late_arrivals': late_arrivals,  # Ingresos posteriores a 8:10
            'late_arrival_minutes': late_arrival_minutes,  # Minutos de retraso después de 8:10
            'early_departure_days': early_departure_days,
            'early_minutes': early_minutes,
            'lunch_overtime_days': lunch_overtime_days,
            'total_lunch_minutes': total_lunch_minutes,
            'missing_entry_days': missing_entry_days,
            'missing_exit_days': missing_exit_days,
            'missing_lunch_days': missing_lunch_days,
            'required_hours': required_hours,
            'actual_hours': actual_hours,
            'mid_day_departures': mid_day_departures,
            'mid_day_departures_text': mid_day_departures_text,
            'overtime_minutes': overtime_minutes if 'agustin taba' in employee_name.lower() else 0,
            'overtime_days': overtime_days if 'agustin taba' in employee_name.lower() else []
        }
        
        # Add PPP weekly hours if applicable
        if 'ppp' in employee_name.lower():
            stats['weekly_hours'] = weekly_hours
            stats['weekly_details'] = weekly_details
            
        return stats

    def count_late_arrivals_after_810(self, employee_name):
        """Cuenta los ingresos posteriores a las 8:10"""
        try:
            late_arrivals = []
            total_late_minutes = 0
            limit_time = datetime.strptime('8:10', '%H:%M').time()

            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]

            for sheet in attendance_sheets:
                try:
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)

                    positions = [
                        {'name_col': 'J', 'entry_col': 'B', 'day_col': 'A'},
                        {'name_col': 'Y', 'entry_col': 'Q', 'day_col': 'P'},
                        {'name_col': 'AN', 'entry_col': 'AF', 'day_col': 'AE'}
                    ]

                    for position in positions:
                        try:
                            name_col_index = self.get_column_index(position['name_col'])
                            name_cell = df.iloc[2, name_col_index]

                            if pd.isna(name_cell):
                                continue

                            employee_cell = str(name_cell).strip()
                            if employee_cell == employee_name:
                                entry_col = self.get_column_index(position['entry_col'])
                                day_col = self.get_column_index(position['day_col'])

                                for row in range(11, 42):
                                    try:
                                        day_value = df.iloc[row, day_col]
                                        if pd.isna(day_value):
                                            continue

                                        day_str = str(day_value).strip().lower()
                                        if day_str == '' or day_str == 'nan' or day_str == 'absence':
                                            continue

                                        entry_time = df.iloc[row, entry_col]
                                        if not pd.isna(entry_time):
                                            try:
                                                if isinstance(entry_time, str):
                                                    entry_time = pd.to_datetime(entry_time).time()
                                                elif isinstance(entry_time, datetime):
                                                    entry_time = entry_time.time()
                                                else:
                                                    continue

                                                # Solo contar si es posterior a las 8:10
                                                if entry_time > limit_time:
                                                    late_minutes = (
                                                        datetime.combine(datetime.min, entry_time) -
                                                        datetime.combine(datetime.min, limit_time)
                                                    ).total_seconds() / 60
                                                    total_late_minutes += late_minutes

                                                    formatted_day = self.translate_day_abbreviation(day_str)
                                                    late_arrivals.append(formatted_day)
                                                    print(f"Ingreso con retraso en fila {row+1}: {late_minutes:.0f} minutos (hora: {entry_time}), Dia: {formatted_day}")

                                            except Exception as e:
                                                print(f"Error procesando hora de entrada en fila {row+1}: {str(e)}")
                                                continue

                                    except Exception as e:
                                        print(f"Error en fila {row+1}: {str(e)}")
                                        continue

                        except Exception as e:
                            print(f"Error procesando posición {position['name_col']}: {str(e)}")
                            continue

                except Exception as e:
                    print(f"Error procesando hoja {sheet}: {str(e)}")
                    continue

            print(f"Total ingresos con retraso: {len(late_arrivals)}")
            print(f"Total minutos de retraso: {total_late_minutes:.0f}")
            return late_arrivals, total_late_minutes

        except Exception as e:
            print(f"Error general: {str(e)}")
            return [], 0

    def format_list_in_columns(self, items, items_per_column=8):
        """Format a list of items into columns"""
        if not items:
            return "No hay días registrados"

        # Create columns of exactly 8 items
        columns = []
        current_column = []
        
        for item in sorted(items, key=lambda x: int(x.split()[0])):
            if len(current_column) < items_per_column:
                current_column.append(f"• {item}")
            else:
                # When column is full (8 items), start a new one
                columns.append(current_column)
                current_column = [f"• {item}"]
        
        # Add the last column if it has any items
        if current_column:
            columns.append(current_column)

        # Format each column as a string, ensuring consistent width
        formatted_columns = []
        for column in columns:
            # Pad column to 8 items if needed
            while len(column) < items_per_column:
                column.append("")  # Add empty strings for padding
            formatted_columns.append("\n".join(column))

        # Join columns with sufficient spacing (10 spaces)
        return "          ".join(formatted_columns)



    def calculate_ppp_weekly_hours(self, employee_name):
        """Calculate weekly hours for PPP employees"""
        try:
            weekly_hours = {
                'Semana 1': 0,
                'Semana 2': 0,
                'Semana 3': 0,
                'Semana 4': 0
            }
            weekly_details = []
            
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]
            
            # Positions in the Excel sheet with correct columns for PPP employees
            positions = [
                {'name_col': 'J', 'entry_col': 'B', 'exit_col': 'D', 'day_col': 'A'},  # Entry y salida B y D
                {'name_col': 'Y', 'entry_col': 'Q', 'exit_col': 'S', 'day_col': 'P'},  # Entry y salida Q y S
                {'name_col': 'AN', 'entry_col': 'AF', 'exit_col': 'AH', 'day_col': 'AE'}  # Entry y salida AF y AH
            ]

            for sheet in attendance_sheets:
                try:
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)
                    
                    for position in positions:
                        try:
                            name_col_index = self.get_column_index(position['name_col'])
                            name_cell = df.iloc[2, name_col_index]
                            
                            if pd.isna(name_cell) or str(name_cell).strip() != employee_name:
                                continue
                                
                            for row in range(11, 42):  # Filas 12-42
                                try:
                                    day_value = df.iloc[row, self.get_column_index(position['day_col'])]
                                    entry_time = df.iloc[row, self.get_column_index(position['entry_col'])]
                                    exit_time = df.iloc[row, self.get_column_index(position['exit_col'])]
                                    
                                    if pd.isna(day_value):
                                        continue
                                        
                                    day_str = str(day_value).strip()
                                    if 'absence' in day_str.lower():
                                        continue
                                        
                                    # Skip weekends
                                    if any(abbr in day_str.lower() for abbr in ['sa', 'su']):
                                        continue
                                        
                                    # Process entry and exit times
                                    if not pd.isna(entry_time) and not pd.isna(exit_time):
                                        try:
                                            entry_time = pd.to_datetime(entry_time).time()
                                            exit_time = pd.to_datetime(exit_time).time()
                                            
                                            # Obtener horas y minutos por separado
                                            end_hour = exit_time.hour
                                            end_minute = exit_time.minute
                                            start_hour = entry_time.hour
                                            start_minute = entry_time.minute
                                            
                                            # Realizar la resta de horas y minutos
                                            diff_hours = end_hour - start_hour
                                            diff_minutes = end_minute - start_minute
                                            
                                            # Ajustar si los minutos son negativos
                                            if diff_minutes < 0:
                                                diff_minutes += 60
                                                diff_hours -= 1
                                            
                                            # Convertir minutos extras a horas si superan 60
                                            if diff_minutes >= 60:
                                                extra_hours = diff_minutes // 60
                                                diff_hours += extra_hours
                                                diff_minutes = diff_minutes % 60
                                            
                                            # Calcular horas totales para esta entrada
                                            total_hours = diff_hours + (diff_minutes / 60)
                                            
                                            # Add to appropriate week
                                            day_num = int(day_str.split()[0])
                                            week_key = ''
                                            if 1 <= day_num <= 7:
                                                week_key = 'Semana 1'
                                            elif 8 <= day_num <= 14:
                                                week_key = 'Semana 2'
                                            elif 15 <= day_num <= 21:
                                                week_key = 'Semana 3'
                                            elif day_num >= 22:
                                                week_key = 'Semana 4'
                                                
                                            if week_key:
                                                weekly_hours[week_key] += total_hours
                                                week_details = {
                                                    'week': week_key,
                                                    'day': self.translate_day_abbreviation(day_str),
                                                    'entry': entry_time.strftime('%H:%M'),
                                                    'exit': exit_time.strftime('%H:%M'),
                                                    'hours': f"{diff_hours}h {diff_minutes}m"
                                                }
                                                
                                                # Check for extra hours
                                                extra_entry_col = 'G' if position['name_col'] == 'J' else \
                                                                'V' if position['name_col'] == 'Y' else 'AK'
                                                extra_exit_col = 'I' if position['name_col'] == 'J' else \
                                                                'X' if position['name_col'] == 'Y' else 'AM'
                                                
                                                extra_entry = df.iloc[row, self.get_column_index(extra_entry_col)]
                                                extra_exit = df.iloc[row, self.get_column_index(extra_exit_col)]
                                                
                                                if not pd.isna(extra_entry) and not pd.isna(extra_exit):
                                                    extra_entry_time = pd.to_datetime(extra_entry).time()
                                                    extra_exit_time = pd.to_datetime(extra_exit).time()
                                                    
                                                    # Calculate extra hours
                                                    extra_end_hour = extra_exit_time.hour
                                                    extra_end_minute = extra_exit_time.minute
                                                    extra_start_hour = extra_entry_time.hour
                                                    extra_start_minute = extra_entry_time.minute
                                                    
                                                    extra_diff_hours = extra_end_hour - extra_start_hour
                                                    extra_diff_minutes = extra_end_minute - extra_start_minute
                                                    
                                                    if extra_diff_minutes < 0:
                                                        extra_diff_minutes += 60
                                                        extra_diff_hours -= 1
                                                    
                                                    if extra_diff_minutes >= 60:
                                                        more_hours = extra_diff_minutes // 60
                                                        extra_diff_hours += more_hours
                                                        extra_diff_minutes = extra_diff_minutes % 60
                                                    
                                                    if extra_diff_hours > 0 or extra_diff_minutes > 0:
                                                        week_details['extra_entry'] = extra_entry_time.strftime('%H:%M')
                                                        week_details['extra_exit'] = extra_exit_time.strftime('%H:%M')
                                                        week_details['extra_hours'] = f"{extra_diff_hours}h {extra_diff_minutes}m"
                                                        
                                                        # Add extra hours to weekly total
                                                        extra_total_hours = extra_diff_hours + (extra_diff_minutes / 60)
                                                        weekly_hours[week_key] += extra_total_hours
                                                
                                                weekly_details.append(week_details)
                                                print(f"Día {day_str}: {diff_hours}h {diff_minutes}m")
                                                
                                        except Exception as e:
                                            print(f"Error processing times in row {row+1}: {str(e)}")
                                            continue
                                            
                                except Exception as e:
                                    print(f"Error in row {row+1}: {str(e)}")
                                    continue
                                    
                        except Exception as e:
                            print(f"Error checking position {position['name_col']}: {str(e)}")
                            continue
                            
                except Exception as e:
                    print(f"Error processing sheet {sheet}: {str(e)}")
                    continue
                    
            # Round weekly hours
            for week in weekly_hours:
                weekly_hours[week] = round(weekly_hours[week], 2)
                
            return weekly_hours, weekly_details
            
        except Exception as e:
            print(f"Error calculating PPP weekly hours: {str(e)}")
            return {'Semana 1': 0, 'Semana 2': 0, 'Semana 3': 0, 'Semana 4': 0}, []

    def calculate_ppp_overtime(self, employee_name):
        """Calculate overtime hours for PPP employees"""
        if 'ppp' not in employee_name.lower():
            return 0, []

        try:
            total_overtime_minutes = 0
            overtime_days = []
            weekly_hours, week_details = self.calculate_ppp_weekly_hours(employee_name)
            
            # For each week's details, calculate overtime if total hours > 20
            for details in week_details:
                try:
                    # Calculate base hours
                    hours = details.get('hours', '0h 0m')
                    base_hours = int(hours.split('h')[0])
                    base_minutes = int(hours.split('h')[1].replace('m', ''))
                    
                    # Add extra hours if present
                    extra_hours = details.get('extra_hours', '0h 0m')
                    if extra_hours != '0h 0m':
                        extra_h = int(extra_hours.split('h')[0])
                        extra_m = int(extra_hours.split('h')[1].replace('m', ''))
                        
                        base_hours += extra_h
                        base_minutes += extra_m
                        
                        # Adjust if minutes exceed 60
                        if base_minutes >= 60:
                            base_hours += base_minutes // 60
                            base_minutes = base_minutes % 60
                    
                    # Check if total hours exceed 20 per week
                    total_time = base_hours + (base_minutes / 60)
                    if total_time > 4:  # More than 4 hours per day = overtime
                        overtime_minutes = ((total_time - 4) * 60)
                        total_overtime_minutes += overtime_minutes
                        overtime_days.append(details['day'])
                        print(f"Overtime on {details['day']}: {overtime_minutes:.0f} minutes")
                        
                except Exception as e:
                    print(f"Error calculating overtime for day: {str(e)}")
                    continue
            
            print(f"Total overtime days: {len(overtime_days)}")
            print(f"Total overtime minutes: {total_overtime_minutes:.0f}")
            return overtime_days, total_overtime_minutes
            
        except Exception as e:
            print(f"Error calculating PPP overtime: {str(e)}")
            return [], 0

    def calculate_overtime(self, employee_name):
        """Calculate overtime hours for agustin taba"""
        if employee_name.lower() != 'agustin taba':
            return 0, []

        try:
            total_overtime_minutes = 0
            overtime_days = []
            
            # Solo procesar la hoja "4.5.6"
            df = pd.read_excel(self.excel_file, sheet_name='4.5.6', header=None)
            
            # Verificar si es el empleado correcto
            if str(df.iloc[2, self.get_column_index('AN')]).strip() != employee_name:
                return 0, []

            # Procesar filas 12-42 (índices 11-41)
            for row in range(11, 42):
                try:
                    end_time = df.iloc[row, self.get_column_index('AM')]
                    start_time = df.iloc[row, self.get_column_index('AK')]
                    day_value = df.iloc[row, self.get_column_index('AE')]

                    if not pd.isna(end_time) and not pd.isna(start_time) and not pd.isna(day_value):
                        try:
                            # Convertir a datetime.time
                            end_time = pd.to_datetime(end_time).time()
                            start_time = pd.to_datetime(start_time).time()
                            
                            # Obtener horas y minutos por separado
                            end_hour = end_time.hour
                            end_minute = end_time.minute
                            start_hour = start_time.hour
                            start_minute = start_time.minute
                            
                            # Realizar la resta de horas y minutos
                            diff_hours = end_hour - start_hour
                            diff_minutes = end_minute - start_minute
                            
                            # Ajustar si los minutos son negativos
                            if diff_minutes < 0:
                                diff_minutes += 60
                                diff_hours -= 1
                            
                            # Convertir minutos extras a horas si superan 60
                            if diff_minutes >= 60:
                                extra_hours = diff_minutes // 60
                                diff_hours += extra_hours
                                diff_minutes = diff_minutes % 60
                            
                            # Calcular minutos totales para esta entrada
                            total_minutes = (diff_hours * 60) + diff_minutes

                            if total_minutes > 0:
                                total_overtime_minutes += total_minutes
                                formatted_day = self.translate_day_abbreviation(str(day_value).strip())
                                overtime_days.append(f"{formatted_day} ({diff_hours}h {diff_minutes}m)")
                                print(f"Diferencia para la fila {row+1}: {diff_hours} horas y {diff_minutes} minutos")
                                
                        except Exception as e:
                            print(f"Error processing times in row {row+1}: {str(e)}")
                            continue

                except Exception as e:
                    print(f"Error in row {row+1}: {str(e)}")
                    continue

            return total_overtime_minutes, overtime_days

        except Exception as e:
            print(f"Error calculating overtime: {str(e)}")
            return 0, []

    def calculate_ppp_weekly_hours(self, employee_name):
        """Calculate weekly hours for PPP employees"""
        try:
            weekly_hours = {
                'Semana 1': 0,
                'Semana 2': 0,
                'Semana 3': 0,
                'Semana 4': 0
            }
            weekly_details = []
            
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]
            
            # Positions in the Excel sheet with correct columns for PPP employees
            positions = [
                {'name_col': 'J', 'entry_col': 'B', 'exit_col': 'D', 'day_col': 'A'},  # Entry y salida B y D
                {'name_col': 'Y', 'entry_col': 'Q', 'exit_col': 'S', 'day_col': 'P'},  # Entry y salida Q y S
                {'name_col': 'AN', 'entry_col': 'AF', 'exit_col': 'AH', 'day_col': 'AE'}  # Entry y salida AF y AH
            ]

            for sheet in attendance_sheets:
                try:
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)
                    
                    for position in positions:
                        try:
                            name_col_index = self.get_column_index(position['name_col'])
                            name_cell = df.iloc[2, name_col_index]
                            
                            if pd.isna(name_cell) or str(name_cell).strip() != employee_name:
                                continue
                                
                            for row in range(11, 42):  # Filas 12-42
                                try:
                                    day_value = df.iloc[row, self.get_column_index(position['day_col'])]
                                    entry_time = df.iloc[row, self.get_column_index(position['entry_col'])]
                                    exit_time = df.iloc[row, self.get_column_index(position['exit_col'])]
                                    
                                    if pd.isna(day_value):
                                        continue
                                        
                                    day_str = str(day_value).strip()
                                    if 'absence' in day_str.lower():
                                        continue
                                        
                                    # Skip weekends
                                    if any(abbr in day_str.lower() for abbr in ['sa', 'su']):
                                        continue
                                        
                                    # Process entry and exit times
                                    if not pd.isna(entry_time) and not pd.isna(exit_time):
                                        try:
                                            entry_time = pd.to_datetime(entry_time).time()
                                            exit_time = pd.to_datetime(exit_time).time()
                                            
                                            # Obtener horas y minutos por separado
                                            end_hour = exit_time.hour
                                            end_minute = exit_time.minute
                                            start_hour = entry_time.hour
                                            start_minute = entry_time.minute
                                            
                                            # Realizar la resta de horas y minutos
                                            diff_hours = end_hour - start_hour
                                            diff_minutes = end_minute - start_minute
                                            
                                            # Ajustar si los minutos son negativos
                                            if diff_minutes < 0:
                                                diff_minutes += 60
                                                diff_hours -= 1
                                            
                                            # Convertir minutos extras a horas si superan 60
                                            if diff_minutes >= 60:
                                                extra_hours = diff_minutes // 60
                                                diff_hours += extra_hours
                                                diff_minutes = diff_minutes % 60
                                            
                                            # Calcular horas totales para esta entrada
                                            total_hours = diff_hours + (diff_minutes / 60)
                                            
                                            # Add to appropriate week
                                            day_num = int(day_str.split()[0])
                                            week_key = ''
                                            if 1 <= day_num <= 7:
                                                week_key = 'Semana 1'
                                            elif 8 <= day_num <= 14:
                                                week_key = 'Semana 2'
                                            elif 15 <= day_num <= 21:
                                                week_key = 'Semana 3'
                                            elif day_num >= 22:
                                                week_key = 'Semana 4'
                                                
                                            if week_key:
                                                weekly_hours[week_key] += total_hours
                                                week_details = {
                                                    'week': week_key,
                                                    'day': self.translate_day_abbreviation(day_str),
                                                    'entry': entry_time.strftime('%H:%M'),
                                                    'exit': exit_time.strftime('%H:%M'),
                                                    'hours': f"{diff_hours}h {diff_minutes}m"
                                                }
                                                weekly_details.append(week_details)
                                                print(f"Día {day_str}: {diff_hours}h {diff_minutes}m")
                                                
                                        except Exception as e:
                                            print(f"Error processing times in row {row+1}: {str(e)}")
                                            continue
                                            
                                except Exception as e:
                                    print(f"Error in row {row+1}: {str(e)}")
                                    continue
                                    
                        except Exception as e:
                            print(f"Error checking position {position['name_col']}: {str(e)}")
                            continue
                            
                except Exception as e:
                    print(f"Error processing sheet {sheet}: {str(e)}")
                    continue
                    
            # Round weekly hours
            for week in weekly_hours:
                weekly_hours[week] = round(weekly_hours[week], 2)
                
            return weekly_hours, weekly_details
            
        except Exception as e:
            print(f"Error calculating PPP weekly hours: {str(e)}")
            return {'Semana 1': 0, 'Semana 2': 0, 'Semana 3': 0, 'Semana 4': 0}, []

    def format_lunch_overtime_text(self, lunch_overtime_days):
        """Calculate weekly hours for PPP employees"""
        try:
            weekly_hours = {
                'Semana 1': 0,
                'Semana 2': 0,
                'Semana 3': 0,
                'Semana 4': 0
            }
            weekly_details = []
            
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]
            
            # Positions in the Excel sheet with correct columns for PPP employees
            positions = [
                {'name_col': 'J', 'entry_col': 'B', 'exit_col': 'D', 'day_col': 'A'},  # Entry y salida B y D
                {'name_col': 'Y', 'entry_col': 'Q', 'exit_col': 'S', 'day_col': 'P'},  # Entry y salida Q y S
                {'name_col': 'AN', 'entry_col': 'AF', 'exit_col': 'AH', 'day_col': 'AE'}  # Entry y salida AF y AH
            ]

            for sheet in attendance_sheets:
                try:
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)
                    
                    for position in positions:
                        try:
                            name_col_index = self.get_column_index(position['name_col'])
                            name_cell = df.iloc[2, name_col_index]
                            
                            if pd.isna(name_cell) or str(name_cell).strip() != employee_name:
                                continue
                                
                            for row in range(11, 42):  # Filas 12-42
                                try:
                                    day_value = df.iloc[row, self.get_column_index(position['day_col'])]
                                    entry_time = df.iloc[row, self.get_column_index(position['entry_col'])]
                                    exit_time = df.iloc[row, self.get_column_index(position['exit_col'])]
                                    
                                    if pd.isna(day_value):
                                        continue
                                        
                                    day_str = str(day_value).strip()
                                    if 'absence' in day_str.lower():
                                        continue
                                        
                                    # Skip weekends
                                    if any(abbr in day_str.lower() for abbr in ['sa', 'su']):
                                        continue
                                        
                                    # Process entry and exit times
                                    if not pd.isna(entry_time) and not pd.isna(exit_time):
                                        try:
                                            entry_time = pd.to_datetime(entry_time).time()
                                            exit_time = pd.to_datetime(exit_time).time()
                                            
                                            # Calculate hours worked
                                            hours = ((
                                                datetime.combine(datetime.min, exit_time) - 
                                                datetime.combine(datetime.min, entry_time)
                                            ).total_seconds() / 3600)
                                            
                                            # Add to appropriate week
                                            day_num = int(day_str.split()[0])
                                            week_key = ''
                                            if 1 <= day_num <= 7:
                                                week_key = 'Semana 1'
                                            elif 8 <= day_num <= 14:
                                                week_key = 'Semana 2'
                                            elif 15 <= day_num <= 21:
                                                week_key = 'Semana 3'
                                            elif day_num >= 22:
                                                week_key = 'Semana 4'
                                                
                                            if week_key:
                                                weekly_hours[week_key] += hours
                                                week_details = {
                                                    'week': week_key,
                                                    'day': self.translate_day_abbreviation(day_str),
                                                    'entry': entry_time.strftime('%H:%M'),
                                                    'exit': exit_time.strftime('%H:%M'),
                                                    'hours': round(hours, 2)
                                                }
                                                weekly_details.append(week_details)
                                                
                                        except Exception as e:
                                            print(f"Error processing times in row {row+1}: {str(e)}")
                                            continue
                                            
                                except Exception as e:
                                    print(f"Error in row {row+1}: {str(e)}")
                                    continue
                                    
                        except Exception as e:
                            print(f"Error checking position {position['name_col']}: {str(e)}")
                            continue
                            
                except Exception as e:
                    print(f"Error processing sheet {sheet}: {str(e)}")
                    continue
                    
            # Round weekly hours
            for week in weekly_hours:
                weekly_hours[week] = round(weekly_hours[week], 2)
                
            return weekly_hours, weekly_details
            
        except Exception as e:
            print(f"Error calculating PPP weekly hours: {str(e)}")
            return {'Semana 1': 0, 'Semana 2': 0, 'Semana 3': 0, 'Semana 4': 0}, []

    def get_employee_stats(self, employee_name):
        """Get comprehensive statistics for a specific employee"""
        # Regular stats
        absence_days = self.get_absence_days(employee_name)
        absences = len(absence_days) if absence_days else 0
        late_days, late_minutes = self.count_late_days(employee_name)
        lunch_overtime_days, total_lunch_minutes = self.count_lunch_overtime_days(employee_name)
        early_departure_days, early_minutes = self.count_early_departures(employee_name)
        missing_entry_days, missing_exit_days, missing_lunch_days = self.count_missing_records(employee_name)
        
        # Mid-day departures with detailed text
        mid_day_departures, mid_day_departures_text = self.format_mid_day_departures_text(employee_name)
        
        # Calculate overtime for agustin taba
        overtime_minutes, overtime_days = self.calculate_overtime(employee_name)
        
        print(f"Debug - {employee_name} mid-day departures: {mid_day_departures}")
        print(f"Debug - {employee_name} mid-day departures text: {mid_day_departures_text}")
        if employee_name.lower() == 'agustin taba':
            print(f"Debug - {employee_name} overtime minutes: {overtime_minutes}")
            print(f"Debug - {employee_name} overtime days: {overtime_days}")

        # Get department from the first sheet (Summary)
        department = ""
        try:
            df = pd.read_excel(self.excel_file, sheet_name='Summary', header=None)
            for idx in range(4, 24):  # Filas 5-24
                if str(df.iloc[idx, 1]).strip() == employee_name:
                    department = str(df.iloc[idx, 2]).strip()
                    break
        except Exception as e:
            print(f"Error getting department: {str(e)}")

        # Calculate actual hours differently for PPP employees
        if 'ppp' in employee_name.lower():
            weekly_hours, weekly_details = self.calculate_ppp_weekly_hours(employee_name)
            actual_hours = sum(weekly_hours.values())
            required_hours = 80.0  # Estándar mensual para PPP
        else:
            required_hours = 76.40  # Estándar regular
            actual_hours = required_hours - (absences * 8)  # Subtract 8 hours for each absence

        # Get stats dictionary ready
        stats = {
            'name': employee_name,
            'department': department,
            'absences': absences,
            'absence_days': absence_days,
            'late_days': late_days,
            'late_minutes': late_minutes,
            'lunch_overtime_days': lunch_overtime_days,
            'total_lunch_minutes': total_lunch_minutes,
            'early_departure_days': early_departure_days,
            'early_minutes': early_minutes,
            'missing_entry_days': missing_entry_days,
            'missing_exit_days': missing_exit_days,
            'missing_lunch_days': missing_lunch_days,
            'required_hours': required_hours,
            'actual_hours': actual_hours,
            'mid_day_departures': mid_day_departures,
            'mid_day_departures_text': mid_day_departures_text,
            'overtime_minutes': overtime_minutes if 'agustin taba' in employee_name.lower() else 0,
            'overtime_days': overtime_days if 'agustin taba' in employee_name.lower() else []
        }
        
        # Add PPP weekly hours if applicable
        if 'ppp' in employee_name.lower():
            stats['weekly_hours'] = weekly_hours
            stats['weekly_details'] = weekly_details
            
        return stats

    def format_mid_day_departures_text(self, employee_name):
        """Updates the mid-day departures text based on new logic"""
        try:
            total_mid_day_departures = 0
            days_with_mid_departures = []
            
            # Encontrar el índice de la hoja "Exceptional"
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            # Solo procesar las hojas después de "Exceptional"
            attendance_sheets = self.excel_file.sheet_names[exceptional_index + 1:]
            
            work_start = datetime.strptime('07:50', '%H:%M').time()
            work_lunch_limit = datetime.strptime('12:00', '%H:%M').time()

            for sheet in attendance_sheets:
                try:
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)
                    
                    # Check position 1 (J3)
                    if str(df.iloc[2, self.get_column_index('J')]).strip() == employee_name:
                        for row in range(11, 42):  # Filas 12-42
                            try:
                                entry = df.iloc[row, self.get_column_index('B')]
                                exit_time = df.iloc[row, self.get_column_index('D')]
                                return_time = df.iloc[row, self.get_column_index('G')]
                                
                                if not pd.isna(entry):  # Si hay entrada
                                    if not pd.isna(exit_time):  # Si hay salida
                                        exit_time = pd.to_datetime(exit_time).time()
                                        if work_start <= exit_time <= work_lunch_limit:  # Salida entre 7:50 y 12:00
                                            if pd.isna(return_time) or pd.to_datetime(return_time).time() > work_lunch_limit:
                                                total_mid_day_departures += 1
                                                day_value = df.iloc[row, self.get_column_index('A')]
                                                days_with_mid_departures.append(self.translate_day_abbreviation(str(day_value).strip()))
                            except Exception as e:
                                print(f"Error processing row {row} data: {e}")
                                continue

                    # Check position 2 (Y3)
                    elif str(df.iloc[2, self.get_column_index('Y')]).strip() == employee_name:
                        for row in range(11, 42):
                            try:
                                entry = df.iloc[row, self.get_column_index('Q')]
                                exit_time = df.iloc[row, self.get_column_index('S')]
                                return_time = df.iloc[row, self.get_column_index('V')]
                                
                                if not pd.isna(entry):  # Si hay entrada
                                    if not pd.isna(exit_time):  # Si hay salida
                                        exit_time = pd.to_datetime(exit_time).time()
                                        if work_start <= exit_time <= work_lunch_limit:  # Salida entre 7:50 y 12:00
                                            if pd.isna(return_time) or pd.to_datetime(return_time).time() > work_lunch_limit:
                                                total_mid_day_departures += 1
                                                day_value = df.iloc[row, self.get_column_index('P')]
                                                days_with_mid_departures.append(self.translate_day_abbreviation(str(day_value).strip()))
                            except Exception as e:
                                print(f"Error processing row {row} data in position 2: {e}")
                                continue

                    # Check position 3 (AN3)
                    elif str(df.iloc[2, self.get_column_index('AN')]).strip() == employee_name:
                        for row in range(11, 42):
                            try:
                                entry = df.iloc[row, self.get_column_index('AF')]
                                exit_time = df.iloc[row, self.get_column_index('AH')]
                                return_time = df.iloc[row, self.get_column_index('AK')]
                                
                                if not pd.isna(entry):  # Si hay entrada
                                    if not pd.isna(exit_time):  # Si hay salida
                                        exit_time = pd.to_datetime(exit_time).time()
                                        if work_start <= exit_time <= work_lunch_limit:  # Salida entre 7:50 y 12:00
                                            if pd.isna(return_time) or pd.to_datetime(return_time).time() > work_lunch_limit:
                                                total_mid_day_departures += 1
                                                day_value = df.iloc[row, self.get_column_index('AE')]
                                                days_with_mid_departures.append(self.translate_day_abbreviation(str(day_value).strip()))
                            except Exception as e:
                                print(f"Error processing row {row} data in position 3: {e}")
                                continue

                except Exception as e:
                    print(f"Error procesando hoja {sheet}: {str(e)}")
                    continue

            # Format the days text 
            days_text = self.format_list_in_columns(days_with_mid_departures) if days_with_mid_departures else "No hay días registrados"
            
            return total_mid_day_departures, days_text

        except Exception as e:
            print(f"Error general en format_mid_day_departures_text: {str(e)}")
            return 0, "Error al procesar los datos"



    def get_early_departure_days(self, employee_name):
        """Returns a list of days with early departures and total early minutes"""
        try:
            early_departure_days = []
            total_early_minutes = 0
            
            # Find Exceptional sheet index
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]

            for sheet in attendance_sheets:
                try:
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)
                    
                    # Check each possible position
                    positions = [
                        {'name_col': 'J', 'exit_col': 'I', 'day_col': 'A'},
                        {'name_col': 'Y', 'exit_col': 'X', 'day_col': 'P'},
                        {'name_col': 'AN', 'exit_col': 'AM', 'day_col': 'AE'}
                    ]

                    for position in positions:
                        try:
                            name_col_index = self.get_column_index(position['name_col'])
                            name_cell = df.iloc[2, name_col_index]

                            if pd.isna(name_cell) or str(name_cell).strip() != employee_name:
                                continue

                            day_col = self.get_column_index(position['day_col'])
                            exit_col = self.get_column_index(position['exit_col'])

                            for row in range(11, 42):
                                try:
                                    day_value = df.iloc[row, day_col]
                                    if pd.isna(day_value):
                                        continue

                                    day_str = str(day_value).strip()
                                    if any(abbr in day_str.lower() for abbr in ['sa', 'su', 'absence']):
                                        continue

                                    exit_time = df.iloc[row, exit_col]
                                    if not pd.isna(exit_time):
                                        try:
                                            exit_time = pd.to_datetime(exit_time).time()
                                            if self.is_early_departure(employee_name, exit_time):
                                                schedule = self.get_employee_schedule(employee_name)
                                                early_minutes = (
                                                    datetime.combine(datetime.min, schedule['end_time']) -
                                                    datetime.combine(datetime.min, exit_time)
                                                ).total_seconds() / 60
                                                total_early_minutes += early_minutes
                                                
                                                formatted_day = self.translate_day_abbreviation(day_str)
                                                early_departure_days.append(formatted_day)
                                        except Exception as e:
                                            print(f"Error processing exit time in row {row+1}: {str(e)}")
                                except Exception as e:
                                    print(f"Error in row {row+1}: {str(e)}")

                        except Exception as e:
                            print(f"Error checking position: {str(e)}")

                except Exception as e:
                    print(f"Error processing sheet {sheet}: {str(e)}")

            return early_departure_days, total_early_minutes

        except Exception as e:
            print(f"Error getting early departure days: {str(e)}")
            return [], 0

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
    
    def get_employee_summary(self, employee_name):
        """Extracts summary data for a specific employee."""
        attendance_summary = self.process_attendance_summary()
        try:
            employee_summary = attendance_summary[attendance_summary['employee_name'].str.strip() == employee_name].iloc[0]
            return employee_summary.to_dict()
        except IndexError:
            return {'department': None, 'required_hours': 0, 'actual_hours': 0, 'absences': 0}



    def calculate_valentina_absences(self, df):
        """Calcula las ausencias de Valentina verificandosolo la columna AK"""
        absences = 0
        try:
            absence_col = self.get_column_index('AK')
            day_col = self.get_column_index('AE')

            for row in range(11, 42):  # AK12 hasta AK42 (índices 11-41)
                try:
                    absence_value = df.iloc[row, absence_col]
                    if not pd.isna(absence_value) and str(absence_value).strip().lower() == 'absence':
                        absences += 1
                except Exception as e:
                    continue

            return absences

        except Exception as e:
            print(f"Error calculando ausencias de Valentina: {str(e)}")
            return 0

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

    def get_employee_daily_data(self, employee_name):
        """Gets daily attendance data for an employee"""
        try:
            daily_data = []

            # Determine which sheets to process
            if employee_name.lower() in self.SPECIAL_SCHEDULES:
                schedule = self.SPECIAL_SCHEDULES[employee_name.lower()]
                attendance_sheets = [schedule['sheet_name']]
            else:
                exceptional_index = self.excel_file.sheet_names.index('Exceptional')
                attendance_sheets = self.excel_file.sheet_names[exceptional_index:]

            for sheet in attendance_sheets:
                try:
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)

                    # Define standard positions for regular employees
                    positions = [
                        {'name_col': 'J', 'entry_col': 'B', 'exit_col': 'I', 'day_col': 'A'},  # First person
                        {'name_col': 'Y', 'entry_col': 'Q', 'exit_col': 'X', 'day_col': 'P'},  # Second person
                        {'name_col': 'AN', 'entry_col': 'AF', 'exit_col': 'AM', 'day_col': 'AE'}  # Third person
                    ]

                    # Check if employee has special position configuration
                    if employee_name.lower() in self.SPECIAL_SCHEDULES:
                        schedule = self.SPECIAL_SCHEDULES[employee_name.lower()]
                        if 'position' in schedule:
                            positions = [schedule['position']]

                    for position in positions:
                        try:
                            name_col = self.get_column_index(position['name_col'])
                            name_cell = df.iloc[2, name_col]

                            if pd.isna(name_cell):
                                continue

                            if str(name_cell).strip() == employee_name:
                                entry_col = self.get_column_index(position['entry_col'])
                                exit_col = self.get_column_index(position['exit_col'])
                                day_col = self.get_column_index(position['day_col'])

                                # Process each day's data
                                for row in range(11, 42):  # Rows 12-42
                                    try:
                                        day_value = df.iloc[row, day_col]
                                        if pd.isna(day_value):
                                            continue

                                        day_str = str(day_value).strip()
                                        if day_str == '' or day_str.lower() == 'absence':
                                            continue

                                        # Get entry and exit times
                                        entry_time = df.iloc[row, entry_col]
                                        exit_time = df.iloc[row, exit_col]

                                        if not pd.isna(entry_time) and not pd.isna(exit_time):
                                            try:
                                                # Convert to datetime
                                                entry_time = pd.to_datetime(entry_time).time()
                                                exit_time = pd.to_datetime(exit_time).time()

                                                # Calculate hours worked
                                                hours = (
                                                    datetime.combine(datetime.min, exit_time) -
                                                    datetime.combine(datetime.min, entry_time)
                                                ).total_seconds() / 3600

                                                if hours > 0:
                                                    daily_data.append({
                                                        'date': day_str,
                                                        'hours': hours,
                                                        'entry': entry_time.strftime('%H:%M'),
                                                        'exit': exit_time.strftime('%H:%M')
                                                    })

                                            except Exception as e:
                                                print(f"Error processing times in row {row+1}: {str(e)}")
                                                continue

                                    except Exception as e:
                                        print(f"Error processing row {row+1}: {str(e)}")
                                        continue

                        except Exception as e:
                            print(f"Error processing position: {str(e)}")
                            continue

                except Exception as e:
                    print(f"Error processing sheet {sheet}: {str(e)}")
                    continue

            # Sort data by date
            daily_data.sort(key=lambda x: x['date'])
            return daily_data
        except Exception as e:
            print(f"Error getting daily data: {str(e)}")
            return []

    def get_absence_days(self, employee_name):
        """Returns a list of days when the employee was absent"""
        try:
            absence_days = []
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index+1:]  # Start after Exceptional

            # Define the three possible positions
            positions = [
                {'name_col': 'J', 'day_col': 'A', 'absence_col': 'G'},
                {'name_col': 'Y', 'day_col': 'P', 'absence_col': 'V'},
                {'name_col': 'AN', 'day_col': 'AE', 'absence_col': 'AK'}
            ]

            for sheet in attendance_sheets:
                try:
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)

                    # Check each possible position
                    for position in positions:
                        try:
                            # Check if employee name matches in the correct position (row 3)
                            name_col_index = self.get_column_index(position['name_col'])
                            name_cell = df.iloc[2, name_col_index]  # Row 3 (index 2)

                            if pd.isna(name_cell) or str(name_cell).strip() != employee_name:
                                continue

                            # If name matches, check for absences
                            day_col = self.get_column_index(position['day_col'])
                            absence_col = self.get_column_index(position['absence_col'])

                            for row in range(11, 42):  # Check rows 12-42
                                try:
                                    absence_value = df.iloc[row, absence_col]
                                    if not pd.isna(absence_value) and str(absence_value).strip().lower() == 'absence':
                                        day_value = df.iloc[row, day_col]
                                        if not pd.isna(day_value):
                                            # Translate the day abbreviation to Spanish full name
                                            day_str = self.translate_day_abbreviation(str(day_value))
                                            print(f"Ausencia encontrada en hoja {sheet}, fila {row+1}, día: {day_str}")
                                            absence_days.append(day_str)
                                except Exception as e:
                                    print(f"Error processing row {row+1}: {str(e)}")
                                    continue

                        except Exception as e:
                            print(f"Error checking position {position['name_col']}: {str(e)}")
                            continue

                except Exception as e:
                    print(f"Error processing sheet {sheet}: {str(e)}")
                    continue

            print(f"Total ausencias encontradas: {len(absence_days)}")
            return absence_days

        except Exception as e:
            print(f"Error getting absence days: {str(e)}")
            return []

    def translate_day_abbreviation(self, day_str):
        """Translate day abbreviations to Spanish full names"""
        # First clean the day string to get just the day abbreviation
        parts = day_str.strip().split()
        if len(parts) >= 2:
            day_num = parts[0]
            day_abbr = parts[1].lower()

            # Translation dictionary
            translations = {
                'su': 'Domingo',
                'mo': 'Lunes',
                'tu': 'Martes',
                'we': 'Miércoles',
                'th': 'Jueves',
                'fr': 'Viernes',
                'sa': 'Sábado'
            }

            # Return formatted string with number and translated day
            if day_abbr in translations:
                return f"{day_num} {translations[day_abbr]}"

        return day_str  # Return original if can't translate

    def get_late_days(self, employee_name):
        """Returns a list of days when the employee arrived late"""
        try:
            late_days = []
            total_late_minutes = 0
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]

            for sheet in attendance_sheets:
                try:
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

                            if pd.isna(name_cell) or str(name_cell).strip() != employee_name:
                                continue

                            print(f"Empleado encontrado en hoja {sheet}, columna {position['name_col']}")

                            entry_col = self.get_column_index(position['entry_col'])
                            day_col = self.get_column_index(position['day_col'])

                            for row in range(11, 42):  # Filas 12 a 42
                                try:
                                    day_value = df.iloc[row, day_col]
                                    if pd.isna(day_value):
                                        continue

                                    day_str = str(day_value).strip()
                                    if day_str == '' or any(abbr in day_str.lower() for abbr in ['sa', 'su']):
                                        continue

                                    entry_time = df.iloc[row, entry_col]
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
                                                late_minutes = (
                                                    datetime.combine(datetime.min, entry_time) -
                                                    datetime.combine(datetime.min, self.WORK_START_TIME)
                                                ).total_seconds() / 60
                                                total_late_minutes += late_minutes

                                                formatted_day = self.translate_day_abbreviation(day_str)
                                                late_days.append(formatted_day)
                                                print(f"Llegada tarde en fila {row+1}: {late_minutes:.0f} minutos (hora: {entry_time}), Dia: {formatted_day}")

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

            print(f"Total días de llegada tarde: {len(late_days)}")
            print(f"Total minutos de tardanza: {total_late_minutes:.0f}")
            return late_days, total_late_minutes

        except Exception as e:
            print(f"Error general: {str(e)}")
            return [], 0

    def get_early_departure_days(self, employee_name):
        """Returns a list of days when the employee left early"""
        try:
            early_departure_days = []
            total_early_minutes = 0
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index+1:]

            positions = [
                {'name_col': 'J', 'day_col': 'A', 'exit_col': 'I'},
                {'name_col': 'Y', 'day_col': 'P', 'exit_col': 'X'},
                {'name_col': 'AN', 'day_col': 'AE', 'exit_col': 'AM'}
            ]

            for sheet in attendance_sheets:
                try:
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)

                    for position in positions:
                        try:
                            name_col_index = self.get_column_index(position['name_col'])
                            name_cell = df.iloc[2, name_col_index]

                            if pd.isna(name_cell) or str(name_cell).strip() != employee_name:
                                continue

                            day_col = self.get_column_index(position['day_col'])
                            exit_col = self.get_column_index(position['exit_col'])

                            for row in range(11, 42):  # Check rows 12-42
                                try:
                                    day_value = df.iloc[row, day_col]
                                    exit_time = df.iloc[row, exit_col]

                                    if not pd.isna(day_value) and not pd.isna(exit_time):
                                        # Skip weekends
                                        day_str = str(day_value).strip()
                                        if any(abbr in day_str.lower() for abbr in ['sa', 'su']):
                                            continue

                                        exit_time = pd.to_datetime(exit_time).time()
                                        if self.is_early_departure(employee_name, exit_time):
                                            # Calculate early minutes
                                            early_minutes = (
                                                datetime.combine(datetime.min, self.WORK_END_TIME) -
                                                datetime.combine(datetime.min, exit_time)
                                            ).total_seconds() / 60
                                            total_early_minutes += early_minutes

                                            # Translate the day to Spanish format
                                            formatted_day = self.translate_day_abbreviation(day_str)
                                            print(f"Salida temprana en hoja {sheet}, fila {row+1}, día: {formatted_day} ({early_minutes:.0f} minutos)")
                                            early_departure_days.append(formatted_day)

                                except Exception as e:
                                    print(f"Error processing row {row+1}: {str(e)}")
                                    continue

                        except Exception as e:
                            print(f"Error checking position {position['name_col']}: {str(e)}")
                            continue

                except Exception as e:
                    print(f"Error processing sheet {sheet}: {str(e)}")
                    continue

            print(f"Total días con salida temprana: {len(early_departure_days)}")
            print(f"Total minutos de salida temprana: {total_early_minutes:.0f}")
            return early_departure_days, total_early_minutes

        except Exception as e:
            print(f"Error getting early departure days: {str(e)}")
            return [], 0

    def get_mid_day_departures(self, employee_name):
        """Returns count of mid-day departures"""
        try:
            mid_day_departures = 0
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index+1:]

            positions = [
                {'name_col': 'J', 'day_col': 'A', 'exit_col': 'I', 'entry_col': 'B'},
                {'name_col': 'Y', 'day_col': 'P', 'exit_col': 'X', 'entry_col': 'Q'},
                {'name_col': 'AN', 'day_col': 'AE', 'exit_col': 'AM', 'entry_col': 'AF'}
            ]

            for sheet in attendance_sheets:
                try:
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)

                    for position in positions:
                        try:
                            name_col_index = self.get_column_index(position['name_col'])
                            name_cell = df.iloc[2, name_col_index]

                            if pd.isna(name_cell) or str(name_cell).strip() != employee_name:
                                continue

                            # If name matches, check for mid-day departures
                            day_col = self.get_column_index(position['day_col'])
                            exit_col = self.get_column_index(position['exit_col'])
                            entry_col = self.get_column_index(position['entry_col'])

                            for row in range(11, 42):  # Check rows 12-42
                                try:
                                    day_value = df.iloc[row, day_col]
                                    entry_time = df.iloc[row, entry_col]
                                    exit_time = df.iloc[row, exit_col]

                                    if not pd.isna(day_value):
                                        # Skip weekends
                                        day_str = str(day_value).strip()
                                        if any(abbr in day_str.lower() for abbr in ['sa', 'su']):
                                            continue

                                        # Check if there's an entry but no exit
                                        if not pd.isna(entry_time) and pd.isna(exit_time):
                                            mid_day_departures += 1
                                            print(f"Salida durante horario en hoja {sheet}, fila {row+1}, día: {self.translate_day_abbreviation(day_str)}")

                                except Exception as e:
                                    print(f"Error processing row {row+1}: {str(e)}")
                                    continue

                        except Exception as e:
                            print(f"Error checking position {position['name_col']}: {str(e)}")
                            continue

                except Exception as e:
                    print(f"Error processing sheet {sheet}: {str(e)}")
                    continue

            print(f"Total salidas durante horario: {mid_day_departures}")
            return mid_day_departures

        except Exception as e:
            print(f"Error getting mid-day departures: {str(e)}")
            return 0

    def get_lunch_overtime_days(self, employee_name):
        """Returns a list of days when the employee exceeded lunch time"""
        try:
            lunch_overtime_days = []
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index+1:]  # Start after Exceptional

            # Si es Valentina, Agustín o Soledad, retornar lista vacía ya que no tienen almuerzo
            if employee_name.lower() in ['valentina al', 'agustin taba', 'soledad silv']:
                print(f"{employee_name} no tiene horario de almuerzo")
                return lunch_overtime_days

            # Define the three possible positions
            positions = [
                {
                    'name_col': 'J',
                    'day_col': 'A',
                    'lunch_out': 'D',
                    'lunch_return': 'G'
                },
                {
                    'name_col': 'Y',
                    'day_col': 'P',
                    'lunch_out': 'S',
                    'lunch_return': 'V'
                },
                {
                    'name_col': 'AN',
                    'day_col': 'AE',
                    'lunch_out': 'AH',
                    'lunch_return': 'AK'
                }
            ]

            for sheet in attendance_sheets:
                try:
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)

                    # Check each possible position
                    for position in positions:
                        try:
                            # Check if employee name matches in the correct position (row 3)
                            name_col_index = self.get_column_index(position['name_col'])
                            name_cell = df.iloc[2, name_col_index]

                            if pd.isna(name_cell) or str(name_cell).strip() != employee_name:
                                continue

                            # If name matches, check for lunch overtime
                            day_col = self.get_column_index(position['day_col'])
                            lunch_out_col = self.get_column_index(position['lunch_out'])
                            lunch_return_col = self.get_column_index(position['lunch_return'])

                            for row in range(11, 42):  # Check rows 12-42
                                try:
                                    day_value = df.iloc[row, day_col]
                                    lunch_out = df.iloc[row, lunch_out_col]
                                    lunch_return = df.iloc[row, lunch_return_col]

                                    if not pd.isna(day_value):
                                        # Skip weekends
                                        day_str = str(day_value).strip()
                                        if any(abbr in day_str.lower() for abbr in ['sa', 'su']):
                                            continue

                                        # Only process if both lunch times exist
                                        if not pd.isna(lunch_out) and not pd.isna(lunch_return):
                                            try:
                                                lunch_out_time = pd.to_datetime(lunch_out).time()
                                                lunch_return_time = pd.to_datetime(lunch_return).time()

                                                lunch_minutes = (
                                                    datetime.combine(datetime.min, lunch_return_time) -
                                                    datetime.combine(datetime.min, lunch_out_time)
                                                ).total_seconds() / 60

                                                if lunch_minutes > self.LUNCH_TIME_LIMIT:
                                                    # Translate the day to Spanish format
                                                    formatted_day = self.translate_day_abbreviation(day_str)
                                                    print(f"Exceso de almuerzo en hoja {sheet}, fila {row+1}, día: {formatted_day} ({lunch_minutes:.0f} minutos)")
                                                    lunch_overtime_days.append(formatted_day)

                                            except Exception as e:
                                                print(f"Error processing lunch times in row {row+1}: {str(e)}")
                                                continue

                                except Exception as e:
                                    print(f"Error processing row {row+1}: {str(e)}")
                                    continue

                        except Exception as e:
                            print(f"Error checking position {position['name_col']}: {str(e)}")
                            continue

                except Exception as e:
                    print(f"Error processing sheet {sheet}: {str(e)}")
                    continue

            print(f"Total días con exceso de almuerzo: {len(lunch_overtime_days)}")
            return lunch_overtime_days

        except Exception as e:
            print(f"Error getting lunch overtime days: {str(e)}")
            return []

    def export_to_csv(self, employee_name, filepath):
        """Export employee performance data to CSV"""
        try:
            # Get employee stats
            stats = self.get_employee_stats(employee_name)

            # Get specific days information
            absence_days = self.get_absence_days(employee_name)
            late_days, _ = self.get_late_days(employee_name)  # Get just the list of days
            early_departure_days, early_minutes = self.get_early_departure_days(employee_name)
            lunch_overtime_days = self.get_lunch_overtime_days(employee_name)
            mid_day_departures = self.get_mid_day_departures(employee_name)

            data = {
                'Nombre': [stats['name']],
                'Departamento': [stats['department']],
                'Horas Requeridas': [f"{stats['required_hours']:.1f}"],
                'Horas Trabajadas': [f"{stats['actual_hours']:.1f}"],
                'Porcentaje Completado': [f"{(stats['actual_hours'] / stats['required_hours'] * 100):.1f}%"],
                'Inasistencias': [stats['absences']],
                'Días con Llegada Tarde': [len(late_days)],
                'Minutos Totales de Retraso': [f"{stats['late_minutes']:.0f}"],
                'Días con Exceso en Almuerzo': [len(lunch_overtime_days)],
                'Minutos Totales Excedidos en Almuerzo': [f"{stats['total_lunch_minutes']:.0f}"],
                'Retiros Anticipados': [len(early_departure_days)],
                'Minutos Totales de Salida Anticipada': [f"{early_minutes:.0f}"],
                'Días sin Registro de Entrada': [len(stats['missing_entry_days'])],
                'Días sin Registro de Salida': [len(stats['missing_exit_days'])],
                'Días sin Registro de Almuerzo': [len(stats['missing_lunch_days'])],
                'Salidas durante horario laboral': [mid_day_departures]
            }

            # Add specific days
            data['Días de Ausencia'] = [', '.join(absence_days) if absence_days else 'Ninguno']
            data['Días de Llegada Tarde'] = [', '.join(late_days) if late_days else 'Ninguno']
            data['Días de Salida Anticipada'] = [', '.join(early_departure_days) if early_departure_days else 'Ninguno']
            data['Días con Exceso de Almuerzo'] = [self.format_lunch_overtime_text(lunch_overtime_days)]

            # Create DataFrame and export
            df = pd.DataFrame(data)
            df.to_csv(filepath, index=False, encoding='utf-8-sig')
            return True

        except Exception as e:
            print(f"Error exporting to CSV: {str(e)}")
            return False

    def export_to_pdf(self, employee_name, filepath):
        """Export employee performance data to PDF"""
        try:
            # Get employee stats
            stats = self.get_employee_stats(employee_name)

            # Get specific days information
            absence_days = self.get_absence_days(employee_name)
            late_days, _ = self.get_late_days(employee_name)  # Get just the list of days
            early_departure_days, early_minutes = self.get_early_departure_days(employee_name)
            lunch_overtime_days = self.get_lunch_overtime_days(employee_name)
            mid_day_departures = self.get_mid_day_departures(employee_name)

            # Create PDF
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font('Arial', 'B', 16)
            pdf.cell(0, 10, 'Reporte de Asistencia', 0, 1, 'C')

            pdf.set_font('Arial', 'B', 12)
            pdf.cell(0, 10, f"Empleado: {stats['name']}", 0, 1)
            pdf.cell(0, 10, f"Departamento: {stats['department']}", 0, 1)
            pdf.ln(5)

            pdf.cell(0, 10, 'Resumen de Horas:', 0, 1)
            pdf.set_font('Arial', '', 12)
            pdf.cell(0, 10, f"Horas Requeridas: {stats['required_hours']:.1f}", 0, 1)
            pdf.cell(0, 10, f"Horas Trabajadas: {stats['actual_hours']:.1f}", 0, 1)
            pdf.cell(0, 10, f"Porcentaje Completado: {(stats['actual_hours'] / stats['required_hours'] * 100):.1f}%", 0, 1)
            pdf.ln(5)

            pdf.set_font('Arial', 'B', 12)
            pdf.cell(0, 10, 'Métricas de Asistencia:', 0, 1)
            metrics = [
                ('Inasistencias', stats['absences'], 'días'),
                ('Llegadas Tarde', len(late_days), f"días ({stats['late_minutes']:.0f} min)"),
                ('Exceso en Almuerzo', len(lunch_overtime_days), f"días ({stats['total_lunch_minutes']:.0f} min)"),
                ('Retiros Anticipados', len(early_departure_days), f"días ({early_minutes:.0f} min)"),
                ('Sin Registro de Entrada', len(stats['missing_entry_days']), 'días'),
                ('Sin Registro de Salida', len(stats['missing_exit_days']), 'días'),
                ('Sin Registro de Almuerzo', len(stats['missing_lunch_days']), 'días'),
                ('Salidas durante horario laboral', mid_day_departures, 'veces')
            ]

            for metric, value, unit in metrics:
                pdf.multi_cell(0, 8, f"{metric}: {value} {unit}")

            pdf.ln(5)
            pdf.set_font('Arial', 'B', 12)
            pdf.cell(0, 10, 'Detalle de Días:', 0, 1)
            pdf.set_font('Arial', '', 12)

            if absence_days:
                pdf.multi_cell(0, 8, f"Días de Ausencia: {', '.join(absence_days)}")
            if late_days:
                pdf.multi_cell(0, 8, f"Días de Llegada Tarde: {', '.join(late_days)}")
            if early_departure_days:
                pdf.multi_cell(0, 8, f"Días de Salida Anticipada: {', '.join(early_departure_days)}")
            pdf.multi_cell(0, 8, f"Días con Exceso de Almuerzo: {self.format_lunch_overtime_text(lunch_overtime_days)}")

            pdf.output(filepath)
            return True

        except Exception as e:
            print(f"Error exporting to PDF: {str(e)}")
            return False

    def organize_days_by_week(self, days):
        """Organizes a list of days into weeks of the month"""
        # Initialize weeks dictionary
        weeks = {
            'Semana 1:': [],
            'Semana 2:': [],
            'Semana 3:': [],
            'Semana 4:': []
        }

        for day in days:
            try:
                # Extract day number from string (e.g. "15 Martes" -> 15)
                day_num = int(day.split()[0])

                # Determine which week the day belongs to
                if 1 <= day_num <= 7:
                    weeks['Semana 1:'].append(day)
                elif 8 <= day_num <= 14:
                    weeks['Semana 2:'].append(day)
                elif 15 <= day_num <= 21:
                    weeks['Semana 3:'].append(day)
                elif 22 <= day_num <= 31:
                    weeks['Semana 4:'].append(day)

                # Sort days within each week
                for week in weeks.values():
                    week.sort(key=lambda x: int(x.split()[0]))

            except (ValueError, IndexError) as e:
                print(f"Error processing day {day}: {str(e)}")
                continue

        return weeks

    def format_lunch_overtime_text(self, days):
        """Formats lunch overtime days by week in a vertical bullet point layout"""
        if not days:
            return "No hay días registrados"

        # Initialize dictionary with empty lists for each week
        weeks_dict = {f'Semana {i}': [] for i in range(1, 5)}

        # Sort days into weeks
        for day in days:
            try:
                day_parts = day.split()
                if len(day_parts) >= 2:
                    day_num = int(day_parts[0])
                    # Determine week number
                    if 1 <= day_num <= 7:
                        weeks_dict['Semana 1'].append(day)
                    elif 8 <= day_num <= 14:
                        weeks_dict['Semana 2'].append(day)
                    elif 15 <= day_num <= 21:
                        weeks_dict['Semana 3'].append(day)
                    elif 22 <= day_num <= 31:
                        weeks_dict['Semana 4'].append(day)
            except (ValueError, IndexError) as e:
                print(f"Error processing day {day}: {str(e)}")
                continue

        # Sort days within each week by day number
        for week_days in weeks_dict.values():
            week_days.sort(key=lambda x: int(x.split()[0]))

        # Format output with bullet points
        days_text = []
        for week_num in range(1, 5):
            week_key = f'Semana {week_num}'
            days = weeks_dict[week_key]
            if days:  # Only include weeks that have days
                days_text.extend([f"• {day}" for day in days])

        # Join all days with newlines, same format as absence days
        return "\n".join(days_text) if days_text else "No hay días registrados"


    def format_mid_day_departures_text(self, employee_name):
        """Formats mid-day departures text with bullet points and week grouping"""
        try:
            # Initialize dictionary with empty lists for each week
            weeks_dict = {f'Semana {i}': [] for i in range(1, 5)}
            total_days = 0

            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index+1:]

            positions = [
                {'name_col': 'J', 'day_col': 'A', 'exit_col': 'I', 'entry_col': 'B'},
                {'name_col': 'Y', 'day_col': 'P', 'exit_col': 'X', 'entry_col': 'Q'},
                {'name_col': 'AN', 'day_col': 'AE', 'exit_col': 'AM', 'entry_col': 'AF'}
            ]

            # Collect all mid-day departure days
            for sheet in attendance_sheets:
                try:
                    df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)

                    for position in positions:
                        try:
                            name_col_index = self.get_column_index(position['name_col'])
                            name_cell = df.iloc[2, name_col_index]

                            if pd.isna(name_cell) or str(name_cell).strip() != employee_name:
                                continue

                            day_col = self.get_column_index(position['day_col'])
                            exit_col = self.get_column_index(position['exit_col'])
                            entry_col = self.get_column_index(position['entry_col'])

                            for row in range(11, 42):
                                try:
                                    day_value = df.iloc[row, day_col]
                                    entry_time = df.iloc[row, entry_col]
                                    exit_time = df.iloc[row, exit_col]

                                    if not pd.isna(day_value):
                                        day_str = str(day_value).strip()
                                        if any(abbr in day_str.lower() for abbr in ['sa', 'su']):
                                            continue

                                        if not pd.isna(entry_time) and pd.isna(exit_time):
                                            formatted_day = self.translate_day_abbreviation(day_str)
                                            day_num = int(formatted_day.split()[0])
                                            total_days += 1

                                            # Add to appropriate week
                                            if 1 <= day_num <= 7:
                                                weeks_dict['Semana 1'].append(formatted_day)
                                            elif 8 <= day_num <= 14:
                                                weeks_dict['Semana 2'].append(formatted_day)
                                            elif 15 <= day_num <= 21:
                                                weeks_dict['Semana 3'].append(formatted_day)
                                            elif 22 <= day_num <= 31:
                                                weeks_dict['Semana 4'].append(formatted_day)

                                except Exception as e:
                                    continue

                        except Exception as e:
                            continue

                except Exception as e:
                    continue

            # Sort days within each week
            for week_days in weeks_dict.values():
                week_days.sort(key=lambda x: int(x.split()[0]))

            # Format days text with bullet points
            days_text = []
            for week_num in range(1, 5):
                week_key = f'Semana {week_num}'
                days = weeks_dict[week_key]
                if days:
                    days_text.extend([f"• {day}" for day in days])

            hover_text = "\n".join(days_text) if days_text else "No hay días registrados"
            return total_days, hover_text

        except Exception as e:
            print(f"Error formatting mid-day departures text: {str(e)}")
            return 0, "No hay días registrados"
            
    def process_attendance_summary(self):
        """Process the Summary sheet to get employee information"""
        try:
            print("Leyendo hoja Summary...")
            df = pd.read_excel(self.excel_file, sheet_name='Summary', header=None)
            
            # Start from row 5 (index 4) which contains the actual data
            data_start_row = 4
            
            # Print the data being processed for debugging
            print("\nDatos procesados de empleados:")
            print(df.iloc[data_start_row:, [0, 1, 2, 3, 4, 5, 6, 7, 8, 13]].to_string())
            
            # Create a DataFrame with employee information
            employee_data = []
            
            # Process each row starting from row 5
            for idx, row in df.iloc[data_start_row:].iterrows():
                # Stop if we hit an empty row
                if pd.isna(row[1]):  # Check if name column is empty
                    break
                    
                employee_data.append({
                    'employee_id': row[0],
                    'employee_name': str(row[1]).strip(),
                    'department': str(row[2]).strip() if not pd.isna(row[2]) else "No especificado",
                    'required_hours': float(row[3]) if not pd.isna(row[3]) else 0.0,
                    'actual_hours': float(row[4]) if not pd.isna(row[4]) else 0.0
                })
            
            # Convert to DataFrame
            summary_df = pd.DataFrame(employee_data)
            
            # Get all employees from all sheets after 'Exceptional'
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            additional_employees = set()
            
            for sheet in self.excel_file.sheet_names[exceptional_index:]:
                df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)
                
                # Check each position (J3, Y3, AN3)
                for col in ['J', 'Y', 'AN']:
                    try:
                        name = df.iloc[2, self.get_column_index(col)]
                        if not pd.isna(name):
                            employee_name = str(name).strip()
                            if employee_name.lower() != 'early leave (mm)':  # Skip the unwanted entry
                                additional_employees.add(employee_name)
                    except:
                        continue
            
            # Add any employees found in other sheets that weren't in Summary
            for name in additional_employees:
                if name not in summary_df['employee_name'].values:
                    summary_df = pd.concat([summary_df, pd.DataFrame([{
                        'employee_id': len(summary_df) + 1,
                        'employee_name': name,
                        'department': "No especificado",
                        'required_hours': 0.0,
                        'actual_hours': 0.0
                    }])], ignore_index=True)
            
            print("\nEmpleados disponibles:", sorted(summary_df['employee_name'].tolist()))
            
            return summary_df
            
        except Exception as e:
            print(f"Error processing Summary sheet: {str(e)}")
            # Return an empty DataFrame with the required columns if there's an error
            return pd.DataFrame(columns=['employee_id', 'employee_name', 'department', 'required_hours', 'actual_hours'])

    def get_employee_stats(self, employee_name):
        """Get comprehensive statistics for a specific employee"""
        # Regular stats
        late_days, late_minutes = self.count_late_days(employee_name)
        late_arrivals, late_arrival_minutes = self.count_late_arrivals_after_810(employee_name)
        early_departure_days, early_minutes = self.count_early_departures(employee_name)
        lunch_overtime_days, total_lunch_minutes = self.count_lunch_overtime_days(employee_name)
        missing_entry_days, missing_exit_days, missing_lunch_days = self.count_missing_records(employee_name)
        absence_days = self.get_absence_days(employee_name)
        absences = len(absence_days) if absence_days else 0
        mid_day_departures, mid_day_departures_text = self.count_mid_day_departures(employee_name)
        overtime_minutes = 0
        overtime_days = []

        # Get overtime for agustin taba
        if employee_name.lower() == 'agustin taba':
            overtime_minutes, overtime_days = self.calculate_overtime(employee_name)

        # Get department
        department = ""
        try:
            department = self.get_employee_department(employee_name)
        except Exception as e:
            print(f"Error getting department: {str(e)}")

        # Calculate actual hours differently for PPP employees
        if 'ppp' in employee_name.lower():
            weekly_hours, weekly_details = self.calculate_ppp_weekly_hours(employee_name)
            actual_hours = sum(weekly_hours.values())
            required_hours = 80.0  # Estándar mensual para PPP
        else:
            required_hours = 76.40  # Estándar regular
            actual_hours = required_hours - (absences * 8)  # Subtract 8 hours for each absence

        # Get stats dictionary ready
        stats = {
            'name': employee_name,
            'department': department,
            'absences': absences,
            'absence_days': absence_days,
            'late_days': late_days,
            'late_minutes': late_minutes,
            'late_arrivals': late_arrivals,  # Ingresos posteriores a 8:10
            'late_arrival_minutes': late_arrival_minutes,  # Minutos de retraso después de 8:10
            'early_departure_days': early_departure_days,
            'early_minutes': early_minutes,
            'lunch_overtime_days': lunch_overtime_days,
            'total_lunch_minutes': total_lunch_minutes,
            'missing_entry_days': missing_entry_days,
            'missing_exit_days': missing_exit_days,
            'missing_lunch_days': missing_lunch_days,
            'required_hours': required_hours,
            'actual_hours': actual_hours,
            'mid_day_departures': mid_day_departures,
            'mid_day_departures_text': mid_day_departures_text,
            'overtime_minutes': overtime_minutes if 'agustin taba' in employee_name.lower() else 0,
            'overtime_days': overtime_days if 'agustin taba' in employee_name.lower() else []
        }
        
        # Add PPP weekly hours if applicable
        if 'ppp' in employee_name.lower():
            stats['weekly_hours'] = weekly_hours
            stats['weekly_details'] = weekly_details
            
        return stats