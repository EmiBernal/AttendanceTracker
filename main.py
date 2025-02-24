import streamlit as st
from utils.excel_processor import ExcelProcessor

# Page configuration
st.set_page_config(
    page_title="Visualizador de Asistencia",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for layout
st.markdown("""
<style>
    .stat-group {
        background-color: rgba(33, 150, 243, 0.1);
        border-radius: 10px;
        padding: 15px;
        margin: 10px 0;
    }
    .stat-card {
        background-color: #F8F9FA;
        border-radius: 8px;
        padding: 12px;
        margin: 5px;
        text-align: center;
        border: 1px solid #E9ECEF;
    }
    .metric-value {
        font-size: 24px;
        font-weight: bold;
        margin: 5px 0;
    }
    .metric-label {
        font-size: 14px;
        color: #6C757D;
    }
    .warning {
        color: #FFC107;
    }
    .danger {
        color: #DC3545;
    }
    .success {
        color: #28A745;
    }
    .auth-required {
        border-left: 4px solid #FFC107;
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

    # Header with employee info
    schedule_note = "📅 Horario Especial" if stats.get('special_schedule', False) else ""
    st.markdown(f"""
        <div class="stat-group">
            <h2>{stats['name']} {schedule_note}</h2>
            <p style="color: #6C757D;">Departamento: {stats['department']}</p>
            {f'<p style="color: #F59E0B; font-size: 14px;">Empleado con horario especial</p>' if stats.get('special_schedule', False) else ''}
        </div>
    """, unsafe_allow_html=True)

    # Work Hours Overview
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
            <div class="stat-group">
                <h3>📊 Resumen de Horas</h3>
        """, unsafe_allow_html=True)

        hours_ratio = (stats['actual_hours'] / stats['required_hours'] * 100) if stats['required_hours'] > 0 else 0
        hours_status = 'success' if hours_ratio >= 95 else 'warning' if hours_ratio >= 85 else 'danger'

        st.markdown(f"""
            <div class="stat-card">
                <div class="metric-label">Horas Trabajadas</div>
                <div class="metric-value {hours_status}">
                    {stats['actual_hours']:.1f}/{stats['required_hours']:.1f}
                </div>
                <div class="metric-label">({hours_ratio:.1f}%)</div>
            </div>
        """, unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with col2:
        st.markdown("""
            <div class="stat-group">
                <h3>📋 Resumen de Asistencia</h3>
        """, unsafe_allow_html=True)

        attendance_ratio = float(stats['attendance_ratio']) * 100
        attendance_status = 'success' if attendance_ratio >= 95 else 'warning' if attendance_ratio >= 85 else 'danger'

        st.markdown(f"""
            <div class="stat-card">
                <div class="metric-label">Tasa de Asistencia</div>
                <div class="metric-value {attendance_status}">{attendance_ratio:.1f}%</div>
            </div>
        """, unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

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