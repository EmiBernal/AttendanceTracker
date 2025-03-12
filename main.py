import streamlit as st
from utils.excel_processor import ExcelProcessor
import os
import webbrowser
from pathlib import Path

# Page configuration
st.set_page_config(
    page_title="Control de Acceso Gampack",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Updated CSS for hover cards and metrics with larger sizes for summary cards
st.markdown("""
<style>
    /* Base transitions */
    .stApp {
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }

    /* File history link styling */
    .file-link {
        display: block;
        padding: 0.75rem 1rem;
        margin: 0.5rem 0;
        color: #E2E8F0;
        text-decoration: none;
        background: rgba(31, 41, 55, 0.6);
        border-radius: 8px;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        border: 1px solid rgba(75, 85, 99, 0.3);
        cursor: pointer;
        animation: slideIn 0.3s ease-out forwards;
    }

    .file-link:hover {
        background: rgba(31, 41, 55, 0.8);
        transform: translateX(5px);
        border-color: #3B82F6;
        box-shadow: 0 4px 6px rgba(59, 130, 246, 0.1);
    }

    .file-link:active {
        transform: translateX(2px);
    }

    .file-link::before {
        content: "📄 ";
        margin-right: 0.5rem;
    }

    /* Info group styling */
    .info-group {
        background: linear-gradient(135deg, rgba(33, 150, 243, 0.05) 0%, rgba(33, 150, 243, 0.1) 100%);
        background-size: 200% 200%;
        border-radius: 16px;
        padding: 24px;
        margin: 20px 0;
        transition: transform 0.5s cubic-bezier(0.4, 0, 0.2, 1),
                    box-shadow 0.5s cubic-bezier(0.4, 0, 0.2, 1);
        border: 1px solid rgba(33, 150, 243, 0.1);
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
        animation: fadeInUp 0.6s ease-out forwards;
        text-align: center;
        position: relative;
        overflow: hidden;
    }

    .info-group:hover {
        transform: translateY(-4px) scale(1.01);
        box-shadow: 0 8px 30px rgba(33, 150, 243, 0.15);
        background-position: right center;
    }

    /* Stat card styling */
    .stat-card {
        background-color: rgba(17, 25, 40, 0.75);
        border-radius: 16px;
        padding: 32px;
        margin: 16px;
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        border: 1px solid rgba(255, 255, 255, 0.1);
        backdrop-filter: blur(10px);
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        position: relative;
        min-height: 200px;
    }

    /* Larger stat cards for summary view */
    .stat-group h3:contains("Métricas Generales del Mes") {
        text-align: center;  /* Centrar el título */
        font-size: 24px;
        margin: 30px 0;
        color: #E2E8F0;
    }

    .stat-group h3:contains("Métricas Generales del Mes") + div .stat-card {
        padding: 40px;
        min-height: 250px;
        margin: 20px;
        text-align: center;  /* Centrar texto para las tarjetas del resumen general */
    }

    .stat-group h3:contains("Métricas Generales del Mes") + div .stat-card .metric-value {
        font-size: 48px;
        margin: 20px 0;
        text-align: center;  /* Asegurar que los valores estén centrados */
    }

    .stat-group h3:contains("Métricas Generales del Mes") + div .stat-card .metric-label {
        font-size: 18px;
        margin: 10px 0;
        text-align: center;  /* Asegurar que las etiquetas estén centradas */
    }

    .stat-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 15px rgba(0, 0, 0, 0.2);
        border-color: rgba(255, 255, 255, 0.2);
        z-index: 1000;
    }

    /* Hover card effect */
    .stat-card .content {
        transition: all 0.3s ease;
    }

    .stat-card:hover .content {
        filter: blur(3px);
        opacity: 0.15;
    }

    .stat-card .hover-text {
        position: fixed;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        opacity: 0;
        transition: all 0.3s ease;
        font-size: 15px;
        font-weight: 500;
        color: #E2E8F0;
        background: rgba(17, 25, 40, 0.95);
        padding: 20px;
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.2);
        pointer-events: none;
        white-space: pre-line;
        text-align: left;
        max-width: 90%;
        line-height: 1.5;
        min-width: 300px;
        max-height: 80vh;
        overflow-y: auto;
        z-index: 1001;
    }

    /* Custom scrollbar styling */
    .stat-card .hover-text::-webkit-scrollbar {
        width: 8px;
    }

    .stat-card .hover-text::-webkit-scrollbar-track {
        background: rgba(0, 0, 0, 0.2);
        border-radius: 4px;
    }

    .stat-card .hover-text::-webkit-scrollbar-thumb {
        background: rgba(255, 255, 255, 0.3);
        border-radius: 4px;
    }

    .stat-card:hover .hover-text {
        opacity: 1;
        z-index: 1001;
    }

    .metric-value {
        font-size: 28px;
        font-weight: 600;
        margin: 8px 0;
        transition: all 0.3s ease;
        color: #E2E8F0;
    }

    .metric-label {
        font-size: 14px;
        color: #94A3B8;
        transition: color 0.3s ease;
    }

    .warning { color: #FBBF24; }
    .danger { color: #EF4444; }
    .success { color: #10B981; }

    /* Sidebar styling */
    .css-1d391kg {  /* Sidebar container */
        background: linear-gradient(180deg, rgba(17, 24, 39, 0.95) 0%, rgba(31, 41, 55, 0.95) 100%);
        padding: 1.5rem;
        border-right: 1px solid rgba(75, 85, 99, 0.3);
        backdrop-filter: blur(10px);
    }

    /* General summary button styling */
    .stButton > button {
        width: 100%;
        border: none;
        background: linear-gradient(135deg, #3B82F6 0%, #1E40AF 100%);
        color: white;
        padding: 0.75rem 1.5rem;
        margin: 0.75rem 0;
        border-radius: 12px;
        font-weight: 500;
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        box-shadow: 0 4px 6px rgba(59, 130, 246, 0.2);
        position: relative;
        overflow: hidden;
    }

    .stButton > button:before {
        content: '';
        position: absolute;
        top: 0;
        left: -100%;
        width: 100%;
        height: 100%;
        background: linear-gradient(
            90deg,
            transparent,
            rgba(255, 255, 255, 0.2),
            transparent
        );
        transition: 0.5s;
    }

    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 12px rgba(59, 130, 246, 0.3);
        background: linear-gradient(135deg, #2563EB 0%, #1E3A8A 100%);
    }

    .stButton > button:hover:before {
        left: 100%;
    }

    .stButton > button:active {
        transform: translateY(0);
        box-shadow: 0 2px 4px rgba(59, 130, 246, 0.2);
    }

    /* Selectbox styling */
    .stSelectbox > div > div {
        background: rgba(31, 41, 55, 0.8);
        border: 1px solid rgba(75, 85, 99, 0.4);
        border-radius: 12px;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        backdrop-filter: blur(8px);
        color: #E5E7EB;
    }

    .stSelectbox > div > div:hover {
        border-color: #3B82F6;
        box-shadow: 0 0 0 2px rgba(59, 130, 246, 0.2);
        background: rgba(31, 41, 55, 0.9);
    }

    /* Subheader styling in sidebar */
    .sidebar .stSubheader {
        color: #60A5FA;
        font-size: 1.1rem;
        font-weight: 600;
        margin-top: 2rem;
        margin-bottom: 1rem;
        padding-bottom: 0.75rem;
        border-bottom: 2px solid rgba(59, 130, 246, 0.2);
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }

    /* Help text styling */
    .stSelectbox .help {
        color: rgba(209, 213, 219, 0.8);
        font-size: 0.9rem;
        margin-top: 0.5rem;
        font-style: italic;
    }

    /* File uploader styling */
    .stFileUploader {
        background: rgba(31, 41, 55, 0.6);
        border: 2px dashed rgba(59, 130, 246, 0.4);
        border-radius: 12px;
        padding: 1.5rem;
        transition: all 0.3s ease;
    }

    .stFileUploader:hover {
        border-color: #3B82F6;
        background: rgba(31, 41, 55, 0.8);
    }

    @keyframes slideIn {
        from {
            opacity: 0;
            transform: translateX(-20px);
        }
        to {
            opacity: 1;
            transform: translateX(0);
        }
    }

    @keyframes fadeInUp {
        from {
            opacity: 0;
            transform: translateY(20px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
</style>
""", unsafe_allow_html=True)

def save_uploaded_file(uploaded_file):
    """Save the uploaded file and return its path"""
    save_dir = Path("uploads")
    save_dir.mkdir(exist_ok=True)

    file_path = save_dir / uploaded_file.name
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return str(file_path)

def create_employee_dashboard(processor, employee_name):
    """Create a detailed dashboard for a single employee"""
    stats = processor.get_employee_stats(employee_name)

    # Get specific days information
    absence_days = processor.get_absence_days(employee_name)
    late_days = stats['late_days']  # Use the list directly from stats
    early_departure_days = stats['early_departure_days']
    lunch_overtime_days = stats['lunch_overtime_days']

    # Header with employee info
    st.markdown(f"""
        <div class="info-group">
            <h2>{stats['name']}</h2>
            <div class="department-label">Departamento: {stats['department']}</div>
        </div>
    """, unsafe_allow_html=True)

    # Hours Summary Card
    hours_ratio = (stats['actual_hours'] / stats['required_hours'] * 100) if stats['required_hours'] > 0 else 0
    hours_status = 'success' if hours_ratio >= 95 else 'warning' if hours_ratio >= 85 else 'danger'

    st.markdown(f"""
        <div class="info-group">
            <h3>📊 Resumen de Horas</h3>
            <div class="metric-value {hours_status}">
                {stats['actual_hours']:.1f}/{stats['required_hours']:.1f}
            </div>
            <div class="metric-label">({hours_ratio:.1f}%)</div>
        </div>
    """, unsafe_allow_html=True)

    # Format days lists for hover text
    absence_days_text = processor.format_list_in_columns(absence_days) if absence_days else "No hay días registrados"
    late_days_text = processor.format_list_in_columns(late_days) if late_days else "No hay días registrados"
    lunch_days_text = processor.format_lunch_overtime_text(lunch_overtime_days)
    mid_day_departures_count, mid_day_departures_text = processor.format_mid_day_departures_text(employee_name)
    early_departure_days_text = processor.format_list_in_columns(early_departure_days) if early_departure_days else "No hay días registrados"

    # Regular Attendance Metrics
    st.markdown("""
        <div class="stat-group">
            <h3>📈 Métricas de Asistencia Regular</h3>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px;">
    """, unsafe_allow_html=True)

    regular_metrics = [
        ('Inasistencias', len(absence_days) if absence_days else 0, "Total días", f"Días sin asistir al trabajo:\n{absence_days_text}"),
        ('Días con Llegada Tarde', len(late_days) if late_days else 0, f"{stats['late_minutes']:.0f} minutos en total", f"Días con llegada tarde:\n{late_days_text}"),
        ('Días con Exceso en Almuerzo', len(lunch_overtime_days) if lunch_overtime_days else 0, f"{stats['total_lunch_minutes']:.0f} minutos en total", f"Días con exceso:\n{lunch_days_text}")
    ]

    for label, value, subtitle, hover_text in regular_metrics:
        status = get_status(value)
        st.markdown(f"""
            <div class="stat-card">
                <div class="content">
                    <div class="metric-label">{label}</div>
                    <div class="metric-value {status}">{value}</div>
                    <div class="metric-label">{subtitle}</div>
                </div>
                <div class="hover-text">{hover_text}</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("</div></div>", unsafe_allow_html=True)

    # Metrics Requiring Authorization
    st.markdown("""
        <div class="stat-group">
            <h3>🔒 Situaciones que Requieren Autorización</h3>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px;">
    """, unsafe_allow_html=True)

    auth_metrics = [
        ('Retiros Anticipados', len(early_departure_days) if early_departure_days else 0, f"{stats['early_minutes']:.0f} minutos en total", f"Días con salida anticipada:\n{early_departure_days_text}"),
        ('Ingresos con Retraso', len(stats['late_arrivals']) if stats['late_arrivals'] else 0, f"{stats['late_arrival_minutes']:.0f} minutos en total", f"Días con ingreso posterior a 8:10:\n{processor.format_list_in_columns(stats['late_arrivals']) if stats['late_arrivals'] else 'No hay días registrados'}")
    ]

    # Solo agregar "Retiros Durante Horario" si no es PPP ni Ana
    if not 'ppp' in employee_name.lower() and employee_name.lower() != 'ana':
        auth_metrics.append(('Retiros Durante Horario', mid_day_departures_count, "Total salidas", f"Salidas durante horario laboral:\n{mid_day_departures_text}"))

    for label, value, subtitle, hover_text in auth_metrics:
        status = get_status(value)
        auth_note = "Requiere Autorización"
        if label == 'Retiros Durante Horario' and employee_name.lower() == 'agustin taba':
            auth_note = "Horario normal de salida (12:40)"

        st.markdown(f"""
            <div class="stat-card">
                <div class="content">
                    <div class="metric-label">{label}</div>
                    <div class="metric-value {status}">{value}</div>
                    <div class="metric-label">{subtitle}</div>
                    <div class="metric-label warning">{auth_note}</div>
                </div>
                <div class="hover-text">{hover_text}</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("</div></div>", unsafe_allow_html=True)

    # Missing Records Section
    create_missing_records_section(stats, processor, employee_name)

# Updated create_missing_records_section to include employee_name parameter
def create_missing_records_section(stats, processor, employee_name):
    """Creates a section for displaying missing records"""
    # Format the lists using the column format
    missing_entry_text = processor.format_list_in_columns(stats['missing_entry_days']) if stats['missing_entry_days'] else "No hay días registrados"
    missing_exit_text = processor.format_list_in_columns(stats['missing_exit_days']) if stats['missing_exit_days'] else "No hay días registrados"
    missing_lunch_text = processor.format_list_in_columns(stats['missing_lunch_days']) if stats['missing_lunch_days'] else "No hay días registrados"

    # Create list of missing records metrics, excluding "Sin Registro de Salida" for Ana
    missing_records = [
        ('Sin Registro de Entrada', len(stats['missing_entry_days']) if stats['missing_entry_days'] else 0, "Total días sin marcar", missing_entry_text),
        ('Sin Registro de Almuerzo', len(stats['missing_lunch_days']) if stats['missing_lunch_days'] else 0, "Total días sin marcar", missing_lunch_text)
    ]

    # Add "Sin Registro de Salida" only if the employee is not Ana
    if employee_name.lower() != 'ana':
        missing_records.append(
            ('Sin Registro de Salida', len(stats['missing_exit_days']) if stats['missing_exit_days'] else 0, "Total días sin marcar", missing_exit_text)
        )

    st.markdown("""
        <div class="stat-group">
            <h3>📋 Registros Faltantes</h3>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px;">
    """, unsafe_allow_html=True)

    for label, value, subtitle, hover_text in missing_records:
        status = get_status(value)
        st.markdown(f"""
            <div class="stat-card">
                <div class="content">
                    <div class="metric-label">{label}</div>
                    <div class="metric-value {status}">{value}</div>
                    <div class="metric-label">{subtitle}</div>
                </div>
                <div class="hover-text">{hover_text}</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("</div></div>", unsafe_allow_html=True)

def create_monthly_summary(processor, attendance_summary):
    """Create a general monthly summary"""
    # Initialize counters for totals and detail dictionaries
    total_absences = 0
    total_late_minutes = 0
    total_lunch_overtime_minutes = 0
    total_early_departure_minutes = 0
    total_mid_day_departures = 0
    total_missing_entry = 0
    total_missing_exit = 0
    total_missing_lunch = 0
    total_late_arrivals = 0
    total_late_arrival_minutes = 0

    # Diccionarios para almacenar detalles por empleado
    absence_details = {}
    late_details = {}
    lunch_details = {}
    early_details = {}
    mid_day_details = {}
    missing_entry_details = {}
    missing_exit_details = {}
    missing_lunch_details = {}
    late_arrival_details = {}

    # Get stats for all employees
    for employee_name in attendance_summary['employee_name'].unique():
        stats = processor.get_employee_stats(employee_name)

        # Absences
        emp_absences = len(stats['absence_days']) if stats['absence_days'] else 0
        if emp_absences > 0:
            absence_details[employee_name] = emp_absences
        total_absences += emp_absences

        # Late minutes
        if stats['late_minutes'] > 0:
            late_details[employee_name] = f"{stats['late_minutes']:.0f} minutos"
        total_late_minutes += stats['late_minutes']

        # Lunch overtime
        if stats['total_lunch_minutes'] > 0:
            lunch_details[employee_name] = f"{stats['total_lunch_minutes']:.0f} minutos"
        total_lunch_overtime_minutes += stats['total_lunch_minutes']

        # Early departures
        if stats['early_minutes'] > 0:
            early_details[employee_name] = f"{stats['early_minutes']:.0f} minutos"
        total_early_departure_minutes += stats['early_minutes']

        # Mid-day departures
        if not 'ppp' in employee_name.lower():
            emp_mid_day = stats['mid_day_departures']
            if emp_mid_day > 0:
                mid_day_details[employee_name] = emp_mid_day
            total_mid_day_departures += emp_mid_day

        # Missing entries
        emp_missing_entry = len(stats['missing_entry_days']) if stats['missing_entry_days'] else 0
        if emp_missing_entry > 0:
            missing_entry_details[employee_name] = emp_missing_entry
        total_missing_entry += emp_missing_entry

        # Missing exits
        emp_missing_exit = len(stats['missing_exit_days']) if stats['missing_exit_days'] else 0
        if emp_missing_exit > 0:
            missing_exit_details[employee_name] = emp_missing_exit
        total_missing_exit += emp_missing_exit

        # Missing lunch
        emp_missing_lunch = len(stats['missing_lunch_days']) if stats['missing_lunch_days'] else 0
        if emp_missing_lunch > 0:
            missing_lunch_details[employee_name] = emp_missing_lunch
        total_missing_lunch += emp_missing_lunch

        # Late arrivals (después de 8:10)
        emp_late_arrivals = len(stats['late_arrivals']) if stats['late_arrivals'] else 0
        if emp_late_arrivals > 0:
            late_arrival_details[employee_name] = f"{emp_late_arrivals} días ({stats['late_arrival_minutes']:.0f} min)"
        total_late_arrivals += emp_late_arrivals
        total_late_arrival_minutes += stats['late_arrival_minutes']

    # Función para formatear detalles
    def format_details(details_dict):
        if not details_dict:
            return "No hay registros"
        return "\n".join(f"• {name}: {value}" for name, value in details_dict.items())

    # Define the metrics to display with updated descriptions and hover details
    summary_metrics = [
        ('Total Inasistencias', total_absences, "Total ausencias", 
         f"Detalles de inasistencias por persona:\n\n{format_details(absence_details)}"),

        ('Total Minutos de Llegada Tarde', f"{total_late_minutes:.0f}", "Total minutos", 
         f"Detalles de llegadas tarde por persona:\n\n{format_details(late_details)}"),

        ('Total Minutos Exceso Almuerzo', f"{total_lunch_overtime_minutes:.0f}", "Total minutos", 
         f"Detalles de exceso en almuerzo por persona:\n\n{format_details(lunch_details)}"),

        ('Total Minutos Retiro Anticipado', f"{total_early_departure_minutes:.0f}", "Total minutos", 
         f"Detalles de retiros anticipados por persona:\n\n{format_details(early_details)}"),

        ('Total Retiros Durante Horario', total_mid_day_departures, "Total retiros", 
         f"Detalles de retiros durante horario por persona:\n\n{format_details(mid_day_details)}"),

        ('Total Ingresos con Retraso', total_late_arrivals, "Total ingresos >8:10", 
         f"Detalles de ingresos posteriores a 8:10 por persona:\n\n{format_details(late_arrival_details)}"),

        ('Total Sin Registro de Entrada', total_missing_entry, "Total registros", 
         f"Detalles de registros de entrada faltantes por persona:\n\n{format_details(missing_entry_details)}"),

        ('Total Sin Registro de Salida', total_missing_exit, "Total registros", 
         f"Detalles de registros de salida faltantes por persona:\n\n{format_details(missing_exit_details)}"),

        ('Total Sin Registro de Almuerzo', total_missing_lunch, "Total registros", 
         f"Detalles de registros de almuerzo faltantes por persona:\n\n{format_details(missing_lunch_details)}")
    ]

    # Display the totals using the same card format as individual employees
    st.markdown("""
        <div class="stat-group">
            <h3>📈 Métricas Generales del Mes</h3>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px;">
    """, unsafe_allow_html=True)

    # Display each metric in a card
    for label, value, subtitle, hover_text in summary_metrics:
        status = get_status(value)
        st.markdown(f"""
            <div class="stat-card">
                <div class="content">
                    <div class="metric-label">{label}</div>
                    <div class="metric-value {status}">{value}</div>
                    <div class="metric-label">{subtitle}</div>
                </div>
                <div class="hover-text">{hover_text}</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("</div></div>", unsafe_allow_html=True)

def get_status(value, warning_threshold=3, danger_threshold=5):
    """Determina el estado (success, warning, danger) basado en el valor"""
    # Si el valor es una lista, usar su longitud
    if isinstance(value, list):
        value = len(value)
    # Si el valor es un número o puede ser convertido a número
    try:
        value = float(value) if not isinstance(value, (int, float)) else value
        return 'success' if value == 0 else 'warning' if value <= warning_threshold else 'danger'
    except (ValueError, TypeError):
        # Si no se puede convertir a número, retornar 'warning' por defecto
        return 'warning'

def main():
    st.title("📊 Control de Acceso Gampack")

    # Initialize session state for file history if it doesn't exist
    if 'file_history' not in st.session_state:
        st.session_state.file_history = []

    # File uploader in sidebar
    with st.sidebar:
        st.subheader("📂 Fuente de Datos")
        uploaded_file = st.file_uploader(
            "Sube el archivo Excel",
            type=['xlsx', 'xls'],
            help="Sube el archivo Excel de asistencia"
        )

        # Display file history
        if uploaded_file is not None and uploaded_file.name not in [f['name'] for f in st.session_state.file_history]:
            # Save the file and add to history
            file_path = save_uploaded_file(uploaded_file)
            st.session_state.file_history.append({
                'name': uploaded_file.name,
                'path': file_path
            })

        # Show file history with modern styling
        if st.session_state.file_history:
            st.subheader("📋 Historial de Archivos")
            for file in st.session_state.file_history:
                st.markdown(
                    f"""<a href="javascript:void(0)" 
                         class="file-link" 
                         >{file['name']}</a>""",
                    unsafe_allow_html=True
                )


    if uploaded_file:
        try:
            processor = ExcelProcessor(uploaded_file)
            attendance_summary = processor.process_attendance_summary()

            # Employee selector and view selector in sidebar
            with st.sidebar:
                st.subheader("📋 Vistas Disponibles")

                # Contenedor para los botones de resumen
                resumen_container = st.container()
                with resumen_container:
                    show_summary = st.button("Ver Resumen General del Mes")
                    show_weekly = st.button("Ver Resumen Semanal")

                st.subheader("👤 Selección de Empleado")
                selected_employee = st.selectbox(
                    "Selecciona unempleado",
                    attendance_summary['employee_name'].unique()
                )

            # Show either monthly summary, weekly summary or employee dashboard
            if show_summary:
                create_monthly_summary(processor, attendance_summary)
            elif show_weekly:
                # Placeholder for weekly summary - you can implement this function later
                st.info("🚧 Funcionalidad de Resumen Semanal en desarrollo")
            else:
                create_employee_dashboard(processor, selected_employee)

        except Exception as e:
            st.error(f"Error procesando el archivo: {str(e)}")
            st.exception(e)

if __name__ == "__main__":
    main()