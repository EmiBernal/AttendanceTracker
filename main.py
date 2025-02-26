import streamlit as st
from utils.excel_processor import ExcelProcessor

# Page configuration
st.set_page_config(
    page_title="Visualizador de Asistencia",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Updated CSS for layout with premium animations and transitions
st.markdown("""
<style>
    /* Base transitions for all components */
    .stApp {
        transition: all 0.3s ease-in-out;
    }

    /* Premium info card animations */
    @keyframes slideInFade {
        from { 
            opacity: 0;
            transform: translateY(30px);
        }
        to { 
            opacity: 1;
            transform: translateY(0);
        }
    }

    @keyframes gradientShift {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }

    /* Info cards with premium feel */
    .info-group {
        background: linear-gradient(135deg, rgba(33, 150, 243, 0.05) 0%, rgba(33, 150, 243, 0.1) 100%);
        background-size: 200% 200%;
        border-radius: 16px;
        padding: 24px;
        margin: 20px 0;
        transition: all 0.5s cubic-bezier(0.4, 0, 0.2, 1);
        position: relative;
        overflow: hidden;
        border: 1px solid rgba(33, 150, 243, 0.1);
        animation: slideInFade 0.8s ease-out, gradientShift 8s ease-in-out infinite;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
    }

    .info-group:hover {
        transform: translateY(-4px) scale(1.01);
        box-shadow: 0 8px 30px rgba(33, 150, 243, 0.15);
        background-position: right center;
    }

    .info-group::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: linear-gradient(45deg, transparent, rgba(255,255,255,0.2), transparent);
        transform: translateX(-100%);
        transition: transform 0.8s cubic-bezier(0.4, 0, 0.2, 1);
    }

    .info-group:hover::before {
        transform: translateX(100%);
    }

    .info-card {
        background: rgba(255, 255, 255, 0.8);
        backdrop-filter: blur(10px);
        border-radius: 12px;
        padding: 20px;
        margin: 10px;
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        border: 1px solid rgba(233, 236, 239, 0.3);
        position: relative;
        overflow: hidden;
    }

    .info-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(33, 150, 243, 0.1);
        border-color: rgba(33, 150, 243, 0.3);
    }

    /* Regular stat cards (existing style) */
    .stat-group {
        background-color: rgba(33, 150, 243, 0.1);
        border-radius: 12px;
        padding: 20px;
        margin: 15px 0;
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        position: relative;
        overflow: hidden;
    }

    .stat-card {
        background-color: rgba(255, 255, 255, 0.05);
        border-radius: 10px;
        padding: 16px;
        margin: 8px;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        border: 1px solid rgba(233, 236, 239, 0.2);
        position: relative;
        overflow: hidden;
    }

    /* Fade in animation */
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(20px); }
        to { opacity: 1; transform: translateY(0); }
    }

    /* Add animation to main components */
    .main, .stat-group, .stat-card {
        animation: fadeIn 0.5s ease-in-out;
    }

    /* Enhanced hover effects with smooth transitions */
    .stat-card::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: linear-gradient(45deg, transparent, rgba(255,255,255,0.1), transparent);
        transform: translateX(-100%);
        transition: transform 0.6s;
    }

    .stat-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 15px rgba(0,0,0,0.1);
        border-color: #2196F3;
    }

    .stat-card:hover::before {
        transform: translateX(100%);
    }

    /* Metric value animations */
    .metric-value {
        font-size: 24px;
        font-weight: bold;
        margin: 8px 0;
        transition: all 0.3s ease;
    }

    .stat-card:hover .metric-value {
        transform: scale(1.02);
        text-shadow: 0 1px 2px rgba(0,0,0,0.05);
    }

    /* Metric label animations */
    .metric-label {
        font-size: 14px;
        color: #6C757D;
        transition: color 0.3s ease;
    }

    .stat-card:hover .metric-label {
        color: #2196F3;
    }

    /* Status colors with transitions */
    .warning { color: #FFC107; }
    .danger { color: #DC3545; }
    .success { color: #28A745; }

    /* Authorization indicator */
    .auth-required {
        border-left: 4px solid #FFC107;
        transition: all 0.3s ease;
    }

    .auth-required:hover {
        border-left-width: 8px;
        background-color: rgba(255, 193, 7, 0.05);
    }

    /* Pulse animation for warning indicators */
    @keyframes subtlePulse {
        0% { transform: scale(1); }
        50% { transform: scale(1.01); }
        100% { transform: scale(1); }
    }

    .stat-card:hover .warning,
    .stat-card:hover .danger {
        animation: subtlePulse 1.5s infinite;
    }

    /* Sidebar transitions */
    .css-1d391kg {  /* Streamlit's sidebar class */
        transition: all 0.3s ease-in-out;
    }

    /* File uploader animations */
    .stFileUploader {
        transition: all 0.3s ease-in-out;
    }

    .stFileUploader:hover {
        transform: translateY(-2px);
    }

    /* Button animations */
    .stButton>button {
        transition: all 0.3s ease-in-out;
    }

    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }

    /* Selectbox animations */
    .stSelectbox {
        transition: all 0.3s ease-in-out;
    }

    .stSelectbox:hover {
        transform: translateY(-2px);
    }
</style>
""", unsafe_allow_html=True)


def create_missing_records_section(stats):
    """Crea una sección expandible para mostrar los días sin registros"""
    with st.expander("📋 Registros Faltantes", expanded=False):
        st.markdown("""
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px;">
        """, unsafe_allow_html=True)

        missing_records = [
            ('Sin Registro de Entrada', stats['missing_entry_days'], "Total días sin marcar"),
            ('Sin Registro de Salida', stats['missing_exit_days'], "Total días sin marcar"),
            ('Sin Registro de Almuerzo', stats['missing_lunch_days'], "Total días sin marcar")
        ]

        for label, value, subtitle in missing_records:
            status = 'success' if value == 0 else 'warning' if value <= 3 else 'danger'
            st.markdown(f"""
                <div class="stat-card">
                    <div class="metric-label">{label}</div>
                    <div class="metric-value {status}">{value}</div>
                    <div class="metric-label">{subtitle}</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("</div>", unsafe_allow_html=True)

def create_employee_dashboard(processor, employee_name):
    """Create a detailed dashboard for a single employee"""
    stats = processor.get_employee_stats(employee_name)

    # Header with employee info using new info-group class
    schedule_note = "📅 Horario Especial" if stats.get('special_schedule', False) else ""
    st.markdown(f"""
        <div class="info-group">
            <h2>{stats['name']} {schedule_note}</h2>
            <p style="color: #6C757D;">Departamento: {stats['department']}</p>
            {f'<p style="color: #F59E0B; font-size: 14px;">Empleado con horario especial</p>' if stats.get('special_schedule', False) else ''}
        </div>
    """, unsafe_allow_html=True)

    # Work Hours Overview
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
            <div class="info-group">
                <h3>📊 Resumen de Horas</h3>
                <div class="info-card">
        """, unsafe_allow_html=True)

        hours_ratio = (stats['actual_hours'] / stats['required_hours'] * 100) if stats['required_hours'] > 0 else 0
        hours_status = 'success' if hours_ratio >= 95 else 'warning' if hours_ratio >= 85 else 'danger'

        st.markdown(f"""
                    <div class="metric-label">Horas Trabajadas</div>
                    <div class="metric-value {hours_status}">
                        {stats['actual_hours']:.1f}/{stats['required_hours']:.1f}
                    </div>
                    <div class="metric-label">({hours_ratio:.1f}%)</div>
                </div>
            </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
            <div class="info-group">
                <h3>📋 Resumen de Asistencia</h3>
                <div class="info-card">
                    <div class="metric-label">Ausencias</div>
                    <div class="metric-value">{stats['absences']}</div>
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
        ('Días con Exceso en Almuerzo', stats['lunch_overtime_days'], f"{stats['total_lunch_minutes']:.0f} minutos en total"),
    ]

    for label, value, subtitle in regular_metrics:
        status = 'success' if value == 0 else 'warning' if value <= 3 else 'danger'
        st.markdown(f"""
            <div class="stat-card">
                <div class="metric-label">{label}</div>
                <div class="metric-value {status}">{value}</div>
                <div class="metric-label">{subtitle}</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("</div></div>", unsafe_allow_html=True)

    # Sección expandible para días sin registros
    create_missing_records_section(stats)

    # Metrics Requiring Authorization
    st.markdown("""
        <div class="stat-group">
            <h3>🔒 Situaciones que Requieren Autorización</h3>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px;">
    """, unsafe_allow_html=True)

    auth_metrics = [
        ('Retiros Anticipados', stats['early_departures'], f"{stats['early_minutes']:.0f} minutos en total"),
        ('Ingresos con Retraso', stats['late_days'], f"{stats['late_minutes']:.0f} minutos en total"),
        ('Retiros Durante Horario Laboral', stats.get('mid_day_departures', 0), "Total ocasiones")
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
            help="Sube el archivo Excel de asistencia con las hojas necesarias"
        )

    if uploaded_file:
        try:
            processor = ExcelProcessor(uploaded_file)
            attendance_summary = processor.process_attendance_summary()

            # Employee selector in sidebar
            with st.sidebar:
                st.subheader("👤 Selección de Empleado")
                selected_employee = st.selectbox(
                    "Selecciona un empleado para ver sus detalles",
                    attendance_summary['employee_name'].unique(),
                    format_func=lambda x: x  # Mostrar el nombre completo
                )

            # Create dashboard for selected employee
            create_employee_dashboard(processor, selected_employee)

        except Exception as e:
            st.error(f"Error procesando el archivo: {str(e)}")
            st.exception(e)

if __name__ == "__main__":
    main()