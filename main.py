import streamlit as st
from utils.excel_processor import ExcelProcessor

# Page configuration
st.set_page_config(
    page_title="Visualizador de Asistencia",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Updated CSS with blur effects
st.markdown("""
<style>
    /* Base transitions */
    .stApp {
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
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

    /* Stat card styling with blur effects */
    .stat-card {
        background-color: rgba(17, 25, 40, 0.75);
        border-radius: 12px;
        padding: 24px;
        margin: 12px;
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        border: 1px solid rgba(255, 255, 255, 0.1);
        backdrop-filter: blur(10px);
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }

    .stat-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 15px rgba(0, 0, 0, 0.2);
        border-color: rgba(255, 255, 255, 0.2);
    }

    /* Hover blur effect */
    .stat-card .content {
        transition: all 0.3s ease;
    }

    .stat-card:hover .content {
        filter: blur(3px);
        opacity: 0.15;
    }

    .stat-card .hover-text {
        position: absolute;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        opacity: 0;
        transition: all 0.3s ease;
        font-size: 15px;
        font-weight: 500;
        color: #E2E8F0;
        background: rgba(17, 25, 40, 0.95);
        padding: 12px 18px;
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.2);
        pointer-events: none;
        white-space: pre-line;
        text-align: center;
        max-width: 90%;
        line-height: 1.5;
    }

    .stat-card:hover .hover-text {
        opacity: 1;
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

    h1, h2, h3 {
        text-align: center;
        animation: fadeInUp 0.6s ease-out forwards;
        color: #E2E8F0;
        text-shadow: 0 0 20px rgba(33, 150, 243, 0.1);
        transition: text-shadow 0.3s ease;
    }

    h1:hover, h2:hover, h3:hover {
        text-shadow: 0 0 30px rgba(33, 150, 243, 0.2);
    }

    .department-label {
        color: #94A3B8;
        font-size: 14px;
        margin-bottom: 8px;
        transition: color 0.3s ease;
    }

    .stSelectbox, .stFileUploader {
        transition: transform 0.3s ease,
                    box-shadow 0.3s ease;
    }

    .stSelectbox:hover, .stFileUploader:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
    }
</style>
""", unsafe_allow_html=True)

def create_employee_dashboard(processor, employee_name):
    """Create a detailed dashboard for a single employee"""
    stats = processor.get_employee_stats(employee_name)

    # Get specific days information
    absence_days = processor.get_absence_days(employee_name)
    late_days = processor.get_late_days(employee_name)
    early_departure_days = processor.get_early_departure_days(employee_name)
    lunch_overtime_days = processor.get_lunch_overtime_days(employee_name)

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

    # Regular Attendance Metrics
    st.markdown("""
        <div class="stat-group">
            <h3>📈 Métricas de Asistencia Regular</h3>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px;">
    """, unsafe_allow_html=True)

    # Format days lists for hover text
    absence_days_text = "\n".join([f"• {day}" for day in absence_days]) if absence_days else "No hay días registrados"
    late_days_text = "\n".join([f"• {day}" for day in late_days]) if late_days else "No hay días registrados"
    lunch_days_text = "\n".join([f"• {day}" for day in lunch_overtime_days]) if lunch_overtime_days else "No hay días registrados"

    regular_metrics = [
        ('Inasistencias', len(absence_days), "Total días", f"Días sin asistir al trabajo:\n{absence_days_text}"),
        ('Días con Llegada Tarde', len(late_days), f"{stats['late_minutes']:.0f} minutos en total", f"Días con retraso:\n{late_days_text}"),
        ('Días con Exceso en Almuerzo', len(lunch_overtime_days), f"{stats['total_lunch_minutes']:.0f} minutos en total", f"Días con exceso:\n{lunch_days_text}")
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

    # Format early departure days for hover text
    early_departure_days_text = "\n".join([f"• {day}" for day in early_departure_days]) if early_departure_days else "No hay días registrados"

    auth_metrics = [
        ('Retiros Anticipados', len(early_departure_days), f"{stats['early_minutes']:.0f} minutos en total", f"Días con salida anticipada:\n{early_departure_days_text}"),
        ('Ingresos con Retraso', len(late_days), f"{stats['late_minutes']:.0f} minutos en total", f"Días con llegada tarde:\n{late_days_text}"),
        ('Retiros Durante Horario', stats.get('mid_day_departures', 0), "Total salidas", "Salidas durante horario laboral")
    ]

    for label, value, subtitle, hover_text in auth_metrics:
        status = get_status(value)
        st.markdown(f"""
            <div class="stat-card">
                <div class="content">
                    <div class="metric-label">{label}</div>
                    <div class="metric-value {status}">{value}</div>
                    <div class="metric-label">{subtitle}</div>
                    <div class="metric-label warning">Requiere Autorización</div>
                </div>
                <div class="hover-text">{hover_text}</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("</div></div>", unsafe_allow_html=True)

    # Missing Records Section
    create_missing_records_section(stats)

def create_missing_records_section(stats):
    """Creates a section for displaying missing records"""
    missing_records = [
        ('Sin Registro de Entrada', len(stats['missing_entry_days']), "Total días sin marcar", "Días sin registro de entrada"),
        ('Sin Registro de Salida', len(stats['missing_exit_days']), "Total días sin marcar", "Días sin registro de salida"),
        ('Sin Registro de Almuerzo', len(stats['missing_lunch_days']), "Total días sin marcar", "Días sin marcar almuerzo")
    ]

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
    st.title("📊 Visualizador de Asistencia")

    # File uploader in sidebar
    with st.sidebar:
        st.subheader("📂 Fuente de Datos")
        uploaded_file = st.file_uploader(
            "Sube el archivo Excel",
            type=['xlsx', 'xls'],
            help="Sube el archivo Excel de asistencia"
        )

    if uploaded_file:
        try:
            processor = ExcelProcessor(uploaded_file)
            attendance_summary = processor.process_attendance_summary()

            # Employee selector in sidebar
            with st.sidebar:
                st.subheader("👤 Selección de Empleado")
                selected_employee = st.selectbox(
                    "Selecciona un empleado",
                    attendance_summary['employee_name'].unique()
                )

            # Create dashboard for selected employee
            create_employee_dashboard(processor, selected_employee)

        except Exception as e:
            st.error(f"Error procesando el archivo: {str(e)}")
            st.exception(e)

if __name__ == "__main__":
    main()