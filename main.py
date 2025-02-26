import streamlit as st
from utils.excel_processor import ExcelProcessor

# Page configuration
st.set_page_config(
    page_title="Visualizador de Asistencia",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Updated CSS with enhanced animations and transitions
st.markdown("""
<style>
    /* Base transitions and animations */
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

    @keyframes shimmer {
        0% { background-position: -1000px 0; }
        100% { background-position: 1000px 0; }
    }

    /* Premium card styling with enhanced animations */
    .info-group {
        background: linear-gradient(135deg, rgba(33, 150, 243, 0.05) 0%, rgba(33, 150, 243, 0.1) 100%);
        border-radius: 16px;
        padding: 24px;
        margin: 20px 0;
        transition: all 0.5s cubic-bezier(0.4, 0, 0.2, 1);
        border: 1px solid rgba(33, 150, 243, 0.1);
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
        animation: fadeInUp 0.6s ease-out forwards;
        position: relative;
        overflow: hidden;
    }

    .info-group::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: linear-gradient(90deg, transparent, rgba(255, 255, 255, 0.2), transparent);
        transform: translateX(-100%);
        transition: transform 0.6s;
    }

    .info-group:hover::before {
        transform: translateX(100%);
    }

    .info-group:hover {
        transform: translateY(-4px) scale(1.01);
        box-shadow: 0 8px 30px rgba(33, 150, 243, 0.15);
    }

    /* Card styling with loading animation */
    .stat-card {
        background-color: rgba(255, 255, 255, 0.05);
        border-radius: 10px;
        padding: 20px;
        margin: 12px;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        border: 1px solid rgba(233, 236, 239, 0.2);
        animation: fadeInUp 0.6s ease-out forwards;
        position: relative;
        overflow: hidden;
    }

    .stat-card.loading {
        background: linear-gradient(90deg, #f0f0f0 25%, #e0e0e0 50%, #f0f0f0 75%);
        background-size: 1000px 100%;
        animation: shimmer 2s infinite linear;
    }

    .stat-card:hover {
        transform: translateY(-2px) scale(1.02);
        box-shadow: 0 8px 24px rgba(0,0,0,0.12);
    }

    .stat-card::after {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 2px;
        background: linear-gradient(90deg, transparent, #2196F3, transparent);
        transform: scaleX(0);
        transition: transform 0.3s ease;
    }

    .stat-card:hover::after {
        transform: scaleX(1);
    }

    /* Enhanced typewriter effect */
    @keyframes typewriter {
        from { width: 0; }
        to { width: 100%; }
    }

    @keyframes blink {
        50% { border-color: transparent; }
    }

    .typewriter {
        overflow: hidden;
        white-space: nowrap;
        border-right: 2px solid;
        animation: 
            typewriter 1s steps(40, end),
            blink 0.75s step-end infinite;
    }

    /* Status colors with transitions */
    .warning { 
        color: #FFC107;
        transition: color 0.3s ease;
    }
    .danger { 
        color: #DC3545;
        transition: color 0.3s ease;
    }
    .success { 
        color: #28A745;
        transition: color 0.3s ease;
    }

    /* Enhanced metrics styling */
    .metric-value {
        font-size: 24px;
        font-weight: bold;
        margin: 8px 0;
        transition: all 0.3s ease;
    }

    .stat-card:hover .metric-value {
        transform: scale(1.05);
        text-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }

    .metric-label {
        font-size: 14px;
        color: #6C757D;
        transition: color 0.3s ease;
    }

    .stat-card:hover .metric-label {
        color: #2196F3;
    }

    /* Navigation transitions */
    .stSelectbox {
        transition: all 0.3s ease;
    }

    .stSelectbox:hover {
        transform: translateY(-2px);
    }

    /* File uploader animations */
    .stFileUploader {
        transition: all 0.3s ease;
    }

    .stFileUploader:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }

    /* Department and special schedule labels */
    .department-label {
        color: #6C757D;
        font-size: 14px;
        margin-bottom: 8px;
        transition: all 0.3s ease;
    }

    .special-schedule {
        color: #F59E0B;
        font-size: 14px;
        margin-top: 4px;
        animation: fadeInUp 0.6s ease-out forwards;
    }

    /* Section headers with gradual reveal */
    h1, h2, h3 {
        opacity: 0;
        animation: fadeInUp 0.6s ease-out forwards;
        animation-delay: 0.2s;
    }
</style>
""", unsafe_allow_html=True)

def create_missing_records_section(stats):
    """Creates a section for displaying missing records"""
    missing_records = [
        ('Sin Registro de Entrada', stats['missing_entry_days'], "Total días sin marcar"),
        ('Sin Registro de Salida', stats['missing_exit_days'], "Total días sin marcar"),
        ('Sin Registro de Almuerzo', stats['missing_lunch_days'], "Total días sin marcar")
    ]

    st.markdown("""
        <div class="stat-group">
            <h3>📋 Registros Faltantes</h3>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px;">
    """, unsafe_allow_html=True)

    for label, value, subtitle in missing_records:
        status = get_status(value)
        st.markdown(f"""
            <div class="stat-card">
                <div class="metric-label">{label}</div>
                <div class="metric-value {status} typewriter">{value}</div>
                <div class="metric-label">{subtitle}</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("</div></div>", unsafe_allow_html=True)

def create_employee_dashboard(processor, employee_name):
    """Create a detailed dashboard for a single employee"""
    stats = processor.get_employee_stats(employee_name)

    # Header with employee info
    st.markdown(f"""
        <div class="info-group">
            <h2>{stats['name']}</h2>
            <div class="department-label">Departamento: {stats['department']}</div>
            {f'<div class="special-schedule">Horario Especial</div>' if stats.get('special_schedule', False) else ''}
        </div>
    """, unsafe_allow_html=True)

    # Hours Summary - Full Width
    hours_ratio = (stats['actual_hours'] / stats['required_hours'] * 100) if stats['required_hours'] > 0 else 0
    hours_status = 'success' if hours_ratio >= 95 else 'warning' if hours_ratio >= 85 else 'danger'

    st.markdown(f"""
        <div class="info-group">
            <h3>📊 Resumen de Horas</h3>
            <div style="text-align: center; padding: 20px;">
                <div class="metric-label">Horas Trabajadas</div>
                <div class="metric-value {hours_status} typewriter" style="font-size: 32px;">
                    {stats['actual_hours']:.1f}/{stats['required_hours']:.1f}
                </div>
                <div class="metric-label">({hours_ratio:.1f}%)</div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    # Regular Attendance Metrics
    st.markdown("""
        <div class="stat-group">
            <h3>📈 Métricas de Asistencia Regular</h3>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px;">
    """, unsafe_allow_html=True)

    regular_metrics = [
        ('Inasistencias', stats['absences'], "Total días"),
        ('Días con Llegada Tarde', stats['late_days'], f"{stats['late_minutes']:.0f} minutos en total"),
        ('Días con Exceso en Almuerzo', stats['lunch_overtime_days'], f"{stats['total_lunch_minutes']:.0f} minutos en total")
    ]

    for label, value, subtitle in regular_metrics:
        status = get_status(value)
        st.markdown(f"""
            <div class="stat-card">
                <div class="metric-label">{label}</div>
                <div class="metric-value {status}">{value}</div>
                <div class="metric-label">{subtitle}</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("</div></div>", unsafe_allow_html=True)

    # Missing Records Section
    create_missing_records_section(stats)

    # Metrics Requiring Authorization
    st.markdown("""
        <div class="stat-group">
            <h3>🔒 Situaciones que Requieren Autorización</h3>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px;">
    """, unsafe_allow_html=True)

    auth_metrics = [
        ('Retiros Anticipados', stats['early_departures'], f"{stats['early_minutes']:.0f} minutos en total"),
        ('Ingresos con Retraso', stats['late_days'], f"{stats['late_minutes']:.0f} minutos en total")
    ]

    for label, value, subtitle in auth_metrics:
        status = 'success' if value == 0 else 'warning' if value <= 2 else 'danger'
        st.markdown(f"""
            <div class="stat-card auth-required">
                <div class="metric-label">{label}</div>
                <div class="metric-value {status}">{value}</div>
                <div class="metric-label">{subtitle}</div>
                <div class="metric-label warning">Requiere Autorización</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("</div></div>", unsafe_allow_html=True)

def get_status(value, warning_threshold=3, danger_threshold=5):
    return 'success' if value == 0 else 'warning' if value <= warning_threshold else 'danger'

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