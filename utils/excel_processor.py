import pandas as pd
import numpy as np
from datetime import datetime, timedelta

class ExcelProcessor:
    def __init__(self, file):
        self.excel_file = pd.ExcelFile(file)
        self.WORK_START_TIME = datetime.strptime('7:50', '%H:%M').time()
        self.WORK_END_TIME = datetime.strptime('17:10', '%H:%M').time()

    def process_attendance_summary(self):
        """Procesa los datos de asistencia desde las hojas después de 'Exceptional'"""
        all_records = []

        # Procesar todas las hojas después de Exceptional
        try:
            exceptional_index = self.excel_file.sheet_names.index('Exceptional')
            attendance_sheets = self.excel_file.sheet_names[exceptional_index:]
        except ValueError:
            print("Warning: Sheet 'Exceptional' not found, processing all sheets")
            attendance_sheets = self.excel_file.sheet_names

        for sheet in attendance_sheets:
            try:
                df = pd.read_excel(self.excel_file, sheet_name=sheet, header=None)
                print(f"Processing sheet: {sheet}")

                # Definir los grupos de empleados
                employee_groups = [
                    # Primer empleado
                    {
                        'nombre': (2, 'J'),  # Fila 3, Columna J
                        'departamento': (2, 'B'),  # Fila 3, Columna B
                        'ausencias': (6, 'A'),  # Fila 7, Columna A
                        'llegadas_tarde': (6, 'I'),  # Fila 7, Columna I
                        'minutos_tarde': [(6, 'J'), (6, 'K')],  # Fila 7, Columnas J,K
                        'salidas_temprano': [(6, 'L'), (6, 'M')],  # Fila 7, Columnas L,M
                        'minutos_temprano': (6, 'N')  # Fila 7, Columna N
                    },
                    # Segundo empleado
                    {
                        'nombre': (2, 'Y'),  # Fila 3, Columna Y
                        'departamento': (2, 'Q'),  # Fila 3, Columna Q
                        'ausencias': (6, 'P'),
                        'llegadas_tarde': (6, 'X'),
                        'minutos_tarde': [(6, 'Y'), (6, 'Z')],
                        'salidas_temprano': [(6, 'AA'), (6, 'AB')],
                        'minutos_temprano': (6, 'AC')
                    },
                    # Tercer empleado
                    {
                        'nombre': (2, 'AN'),  # Fila 3, Columna AN
                        'departamento': (2, 'AF'),  # Fila 3, Columna AF
                        'ausencias': (6, 'AE'),
                        'llegadas_tarde': (6, 'AM'),
                        'minutos_tarde': [(6, 'AN'), (6, 'AO')],
                        'salidas_temprano': [(6, 'AP'), (6, 'AQ')],
                        'minutos_temprano': (6, 'AR')
                    }
                ]

                for group in employee_groups:
                    try:
                        # Obtener nombre y departamento
                        name = str(df.iloc[group['nombre'][0], df.columns.get_loc(group['nombre'][1])]).strip()
                        dept = str(df.iloc[group['departamento'][0], df.columns.get_loc(group['departamento'][1])]).strip()

                        if pd.isna(name) or name == '' or name == 'nan':
                            continue

                        print(f"Found employee: {name} in department: {dept}")

                        # Obtener estadísticas
                        absences = float(df.iloc[group['ausencias'][0], df.columns.get_loc(group['ausencias'][1])]) if pd.notna(df.iloc[group['ausencias'][0], df.columns.get_loc(group['ausencias'][1])]) else 0
                        late_count = float(df.iloc[group['llegadas_tarde'][0], df.columns.get_loc(group['llegadas_tarde'][1])]) if pd.notna(df.iloc[group['llegadas_tarde'][0], df.columns.get_loc(group['llegadas_tarde'][1])]) else 0

                        # Sumar minutos de tardanza
                        late_minutes = sum(
                            float(df.iloc[pos[0], df.columns.get_loc(pos[1])]) 
                            if pd.notna(df.iloc[pos[0], df.columns.get_loc(pos[1])]) else 0 
                            for pos in group['minutos_tarde']
                        )

                        # Sumar veces que se fue temprano
                        early_count = sum(
                            float(df.iloc[pos[0], df.columns.get_loc(pos[1])]) 
                            if pd.notna(df.iloc[pos[0], df.columns.get_loc(pos[1])]) else 0 
                            for pos in group['salidas_temprano']
                        )

                        # Obtener minutos de salida temprana
                        early_minutes = float(df.iloc[group['minutos_temprano'][0], df.columns.get_loc(group['minutos_temprano'][1])]) if pd.notna(df.iloc[group['minutos_temprano'][0], df.columns.get_loc(group['minutos_temprano'][1])]) else 0

                        # Calcular horas totales
                        required_hours = 160  # 8 horas * 20 días
                        actual_hours = required_hours - (absences * 8) - (late_minutes / 60) - (early_minutes / 60)

                        record = {
                            'employee_name': name,
                            'department': dept,
                            'required_hours': required_hours,
                            'actual_hours': actual_hours,
                            'late_count': late_count,
                            'late_minutes': late_minutes,
                            'early_departure_count': early_count,
                            'early_departure_minutes': early_minutes,
                            'absences': absences,
                            'attendance_ratio': (required_hours - (absences * 8)) / required_hours if required_hours > 0 else 0
                        }

                        all_records.append(record)
                        print(f"Added record for {name}")

                    except Exception as e:
                        print(f"Error processing employee in group: {str(e)}")
                        continue

            except Exception as e:
                print(f"Error processing sheet {sheet}: {str(e)}")
                continue

        if not all_records:
            print("No records processed!")
            return pd.DataFrame(columns=[
                'employee_name', 'department', 'required_hours', 'actual_hours',
                'late_count', 'late_minutes', 'early_departure_count',
                'early_departure_minutes', 'absences', 'attendance_ratio'
            ])

        return pd.DataFrame(all_records)

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