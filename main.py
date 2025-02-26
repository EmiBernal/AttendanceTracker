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

    /* Container for slide transitions */
    .view-container {
        position: relative;
        width: 100%;
        overflow: hidden;
    }

    .main-view, .detail-view {
        width: 100%;
        transition: transform 0.5s cubic-bezier(0.4, 0, 0.2, 1);
    }

    .main-view.slide-out {
        transform: translateX(-100%);
    }

    .detail-view {
        position: absolute;
        top: 0;
        left: 100%;
        transform: translateX(0);
        background: linear-gradient(135deg, rgba(33, 150, 243, 0.02) 0%, rgba(33, 150, 243, 0.05) 100%);
        padding: 20px;
        border-radius: 16px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
    }

    .detail-view.slide-in {
        transform: translateX(0);
    }

    /* Premium card styling */
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
        cursor: pointer;
    }

    .info-group:hover {
        transform: translateY(-4px) scale(1.01);
        box-shadow: 0 8px 30px rgba(33, 150, 243, 0.15);
        background-position: right center;
    }

    /* Back button styling */
    .back-button {
        display: inline-flex;
        align-items: center;
        padding: 8px 16px;
        margin-bottom: 20px;
        border: none;
        border-radius: 8px;
        background: rgba(33, 150, 243, 0.1);
        color: #2196F3;
        font-family: 'SF Pro Display', system-ui;
        font-weight: 500;
        cursor: pointer;
        transition: all 0.3s ease;
    }

    .back-button:hover {
        background: rgba(33, 150, 243, 0.2);
        transform: translateX(-4px);
    }

    /* Detail view table */
    .detail-table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 20px;
        font-family: 'SF Pro Display', system-ui;
    }

    .detail-table th,
    .detail-table td {
        padding: 12px;
        text-align: left;
        border-bottom: 1px solid rgba(33, 150, 243, 0.1);
    }

    .detail-table th {
        background: rgba(33, 150, 243, 0.05);
        font-weight: 500;
        color: #2C3E50;
    }

    .detail-table tr:hover {
        background: rgba(33, 150, 243, 0.05);
    }

    /* Rest of your existing styles remain unchanged */
    .stat-card {
        background-color: rgba(255, 255, 255, 0.05);
        border-radius: 10px;
        padding: 20px;
        margin: 12px;
        transition: transform 0.4s cubic-bezier(0.4, 0, 0.2, 1),
                    box-shadow 0.4s cubic-bezier(0.4, 0, 0.2, 1),
                    border-color 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        border: 1px solid rgba(33, 150, 243, 0.1);
        animation: fadeInUp 0.6s ease-out forwards;
        text-align: center;
        position: relative;
        overflow: hidden;
        backdrop-filter: blur(10px);
    }

    .stat-card:hover {
        transform: translateY(-2px) scale(1.02);
        box-shadow: 0 8px 24px rgba(33, 150, 243, 0.15);
        border-color: rgba(33, 150, 243, 0.3);
    }

    .warning { 
        color: #FFC107;
        transition: color 0.3s ease;
        text-shadow: 0 0 10px rgba(255, 193, 7, 0.2);
    }
    .danger { 
        color: #DC3545;
        transition: color 0.3s ease;
        text-shadow: 0 0 10px rgba(220, 53, 69, 0.2);
    }
    .success { 
        color: #28A745;
        transition: color 0.3s ease;
        text-shadow: 0 0 10px rgba(40, 167, 69, 0.2);
    }

    .metric-value {
        font-size: 24px;
        font-weight: bold;
        margin: 8px 0;
        transition: transform 0.3s ease,
                    text-shadow 0.3s ease;
    }

    .stat-card:hover .metric-value {
        transform: scale(1.05);
        text-shadow: 0 0 15px rgba(33, 150, 243, 0.3);
    }

    .metric-label {
        font-size: 14px;
        color: #6C757D;
        transition: color 0.3s ease;
    }

    .stat-card:hover .metric-label {
        color: #2196F3;
    }

    h1, h2, h3 {
        text-align: center;
        animation: fadeInUp 0.6s ease-out forwards;
        animation-delay: 0.2s;
        color: #2C3E50;
        text-shadow: 0 0 20px rgba(33, 150, 243, 0.1);
        transition: text-shadow 0.3s ease;
    }

    h1:hover, h2:hover, h3:hover {
        text-shadow: 0 0 30px rgba(33, 150, 243, 0.2);
    }

    .department-label {
        color: #6C757D;
        font-size: 14px;
        margin-bottom: 8px;
        transition: color 0.3s ease;
    }

    .special-schedule {
        color: #F59E0B;
        font-size: 14px;
        margin-top: 4px;
        animation: fadeInUp 0.6s ease-out forwards;
        text-shadow: 0 0 10px rgba(245, 158, 11, 0.2);
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

<script>
function showDetailView(viewId) {
    const mainView = document.querySelector('.main-view');
    const detailView = document.querySelector('.detail-view-' + viewId);

    mainView.classList.add('slide-out');
    detailView.classList.add('slide-in');
}

function hideDetailView(viewId) {
    const mainView = document.querySelector('.main-view');
    const detailView = document.querySelector('.detail-view-' + viewId);

    mainView.classList.remove('slide-out');
    detailView.classList.remove('slide-in');
}
</script>
""", unsafe_allow_html=True)

def create_hours_detail_view(stats, daily_data):
    """Creates a detailed view for hours worked"""
    st.markdown(f"""
        <div class="detail-view detail-view-hours">
            <button class="back-button" onclick="hideDetailView('hours')">
                ← Volver
            </button>
            <h2>Detalle de Horas Trabajadas</h2>
            <div class="info-group">
                <div class="metric-label">Total de Horas</div>
                <div class="metric-value">{stats['actual_hours']:.1f}/{stats['required_hours']:.1f}</div>
            </div>
            <table class="detail-table">
                <thead>
                    <tr>
                        <th>Fecha</th>
                        <th>Entrada</th>
                        <th>Salida</th>
                        <th>Horas</th>
                        <th>Estado</th>
                    </tr>
                </thead>
                <tbody>
    """, unsafe_allow_html=True)

    for day in daily_data:
        hours_ratio = day['hours'] / stats['required_hours'] * 100 if stats['required_hours'] >0 else 0
        status = 'success' if hours_ratio >= 95 else 'warning' if hours_ratio >= 85 else 'danger'
        st.markdown(f"""
            <tr>
                <td>{day['date']}</td>
                <td>{day['entry']}</td>
                <td>{day['exit']}</td>
                <td>{day['hours']:.1f}</td>
                <td class="{status}">{status}</td>
            </tr>
        """, unsafe_allow_html=True)

    st.markdown("""
                </tbody>
            </table>
        </div>
    """, unsafe_allow_html=True)

def create_employee_dashboard(processor, employee_name):
    """Create a detailed dashboard for a single employee"""
    stats = processor.get_employee_stats(employee_name)
    daily_data = processor.get_employee_daily_data(employee_name)

    # Container for all views
    st.markdown("""
        <div class="view-container">
            <div class="main-view">
    """, unsafe_allow_html=True)

    # Header with employee info
    st.markdown(f"""
        <div class="info-group">
            <h2>{stats['name']}</h2>
            <div class="department-label">Departamento: {stats['department']}</div>
            {f'<div class="special-schedule">Horario Especial</div>' if stats.get('special_schedule', False) else ''}
        </div>
    """, unsafe_allow_html=True)

    # Hours Summary Card - Clickable
    hours_ratio = (stats['actual_hours'] / stats['required_hours'] * 100) if stats['required_hours'] > 0 else 0
    hours_status = 'success' if hours_ratio >= 95 else 'warning' if hours_ratio >= 85 else 'danger'

    st.markdown(f"""
        <div class="info-group" onclick="showDetailView('hours')">
            <h3>📊 Resumen de Horas</h3>
            <div class="metric-label">Horas Trabajadas</div>
            <div class="metric-value {hours_status}" style="font-size: 32px;">
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

    st.markdown("</div>", unsafe_allow_html=True)  # Close main-view

    # Create detail view for hours
    create_hours_detail_view(stats, daily_data)

    st.markdown("</div>", unsafe_allow_html=True)  # Close view-container


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