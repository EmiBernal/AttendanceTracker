import streamlit as st
from utils.excel_processor import ExcelProcessor

# Page configuration
st.set_page_config(
    page_title="Visualizador de Asistencia",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Updated CSS with animations and slide transitions
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

    .metric-value {
        font-size: 32px;
        font-weight: bold;
        margin: 8px 0;
        transition: transform 0.3s ease,
                    text-shadow 0.3s ease;
    }

    .metric-label {
        font-size: 14px;
        color: #6C757D;
        transition: color 0.3s ease;
    }

    .warning { color: #FFC107; }
    .danger { color: #DC3545; }
    .success { color: #28A745; }

    h1, h2, h3 {
        text-align: center;
        animation: fadeInUp 0.6s ease-out forwards;
        color: #2C3E50;
    }
</style>
""", unsafe_allow_html=True)

def create_employee_dashboard(processor, employee_name):
    """Create a detailed dashboard for a single employee"""
    stats = processor.get_employee_stats(employee_name)

    # Header with employee info
    st.markdown(f"""
        <div class="info-group">
            <h2>{stats['name']}</h2>
            <div class="department-label">Departamento: {stats['department']}</div>
        </div>
    """, unsafe_allow_html=True)

    # Hours Summary Card - Simplified version
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
        status = get_status(value)
        st.markdown(f"""
            <div class="stat-card">
                <div class="metric-label">{label}</div>
                <div class="metric-value {status}">{value}</div>
                <div class="metric-label">{subtitle}</div>
                <div class="metric-label warning">Requiere Autorización</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("</div></div>", unsafe_allow_html=True)


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
                <div class="metric-value {status}">{value}</div>
                <div class="metric-label">{subtitle}</div>
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