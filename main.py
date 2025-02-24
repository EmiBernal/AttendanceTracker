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
    .dashboard-grid {
        display: grid;
        grid-template-columns: repeat(5, 1fr);
        grid-template-rows: repeat(5, 1fr);
        gap: 16px;
        padding: 20px;
        height: calc(100vh - 80px);
    }
    .stat-group {
        background-color: #F8F9FA;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .stat-card {
        background-color: white;
        border-radius: 10px;
        padding: 16px;
        height: 100%;
        transition: transform 0.2s, box-shadow 0.2s;
        border: 1px solid #E9ECEF;
    }
    .stat-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    .metric-value {
        font-size: 28px;
        font-weight: 600;
        margin: 8px 0;
        line-height: 1.2;
    }
    .metric-label {
        font-size: 14px;
        color: #6C757D;
        margin-bottom: 4px;
    }
    .metric-subtitle {
        font-size: 12px;
        color: #ADB5BD;
        margin-top: 4px;
    }
    .warning { color: #F59E0B; }
    .danger { color: #EF4444; }
    .success { color: #10B981; }
    .auth-required {
        border-left: 3px solid #F59E0B;
    }
    .section-title {
        font-size: 20px;
        font-weight: 600;
        color: #1F2937;
        margin-bottom: 16px;
    }

    /* Grid Areas */
    .overview-section { grid-row: span 5 / span 5; }
    .hours-section { grid-column: span 2 / span 2; grid-row: span 2 / span 2; }
    .late-days { grid-column-start: 4; }
    .early-leaves { grid-column-start: 4; grid-row-start: 2; }
    .absences { grid-column-start: 5; grid-row-start: 2; }
    .lunch-overtime { grid-column-start: 5; grid-row-start: 1; }
    .attendance-metrics { 
        grid-column: span 2 / span 2;
        grid-row: span 3 / span 3;
        grid-column-start: 2;
        grid-row-start: 3;
    }
    .auth-metrics {
        grid-column: span 2 / span 2;
        grid-row: span 2 / span 2;
        grid-column-start: 4;
        grid-row-start: 4;
    }
    .missing-records {
        grid-column: span 2 / span 2;
        grid-column-start: 4;
        grid-row-start: 3;
    }
</style>
""", unsafe_allow_html=True)

def create_employee_dashboard(processor, employee_name):
    """Create a detailed dashboard for a single employee"""
    stats = processor.get_employee_stats(employee_name)

    st.markdown('<div class="dashboard-grid">', unsafe_allow_html=True)

    # Overview Section
    st.markdown(f"""
        <div class="stat-group overview-section">
            <h1 style="font-size: 32px; font-weight: 700; color: #1F2937; margin: 0;">
                {stats['name']}
            </h1>
            <p style="font-size: 16px; color: #6C757D; margin-top: 8px;">
                {stats['department'].title()}
            </p>
        </div>
    """, unsafe_allow_html=True)

    # Hours Section
    hours_ratio = (stats['actual_hours'] / stats['required_hours'] * 100) if stats['required_hours'] > 0 else 0
    hours_status = 'success' if hours_ratio >= 95 else 'warning' if hours_ratio >= 85 else 'danger'

    st.markdown(f"""
        <div class="stat-group hours-section">
            <div class="section-title">📊 Horas Trabajadas</div>
            <div class="stat-card">
                <div class="metric-value {hours_status}">
                    {stats['actual_hours']:.1f}/{stats['required_hours']:.1f}
                </div>
                <div class="metric-subtitle">({hours_ratio:.1f}% completado)</div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    # Individual Metrics
    metrics = {
        'late-days': ('Llegadas Tarde', stats['late_days'], f"{stats['late_minutes']:.0f} min"),
        'early-leaves': ('Salidas Temprano', stats['early_departures'], f"{stats['early_minutes']:.0f} min"),
        'absences': ('Inasistencias', stats['absences'], "días"),
        'lunch-overtime': ('Exceso Almuerzo', stats['lunch_overtime_days'], f"{stats['total_lunch_minutes']:.0f} min")
    }

    for key, (label, value, subtitle) in metrics.items():
        status = 'success' if value == 0 else 'warning' if value <= 3 else 'danger'
        st.markdown(f"""
            <div class="stat-group {key}">
                <div class="stat-card">
                    <div class="metric-label">{label}</div>
                    <div class="metric-value {status}">{value}</div>
                    <div class="metric-subtitle">{subtitle}</div>
                </div>
            </div>
        """, unsafe_allow_html=True)

    # Missing Records Section
    st.markdown(f"""
        <div class="stat-group missing-records">
            <div class="section-title">📋 Registros Faltantes</div>
            <div class="stat-card">
                <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px;">
                    <div>
                        <div class="metric-label">Entrada</div>
                        <div class="metric-value {get_status(stats['missing_entry_days'])}">
                            {stats['missing_entry_days']}
                        </div>
                    </div>
                    <div>
                        <div class="metric-label">Salida</div>
                        <div class="metric-value {get_status(stats['missing_exit_days'])}">
                            {stats['missing_exit_days']}
                        </div>
                    </div>
                    <div>
                        <div class="metric-label">Almuerzo</div>
                        <div class="metric-value {get_status(stats['missing_lunch_days'])}">
                            {stats['missing_lunch_days']}
                        </div>
                    </div>
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

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
                    format_func=lambda x: x
                )

            # Create dashboard for selected employee
            create_employee_dashboard(processor, selected_employee)

        except Exception as e:
            st.error(f"Error procesando el archivo: {str(e)}")
            st.exception(e)

if __name__ == "__main__":
    main()