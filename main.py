import streamlit as st
from utils.excel_processor import ExcelProcessor

# Page configuration
st.set_page_config(
    page_title="Visualizador de Asistencia",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Updated CSS with hover menu and animations
st.markdown("""
<style>
    /* Base transitions and animations */
    .stApp {
        transition: all 0.3s ease-in-out;
    }

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

    /* Premium info cards with advanced animations */
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

    /* Regular stat cards with hover effects */
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

    /* Status colors with transitions */
    .warning { color: #FFC107; transition: all 0.3s ease; }
    .danger { color: #DC3545; transition: all 0.3s ease; }
    .success { color: #28A745; transition: all 0.3s ease; }

    /* Metrics styling with animations */
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

    .metric-label {
        font-size: 14px;
        color: #6C757D;
        transition: color 0.3s ease;
    }

    .stat-card:hover .metric-label {
        color: #2196F3;
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

    /* Header text styling */
    h1, h2, h3 {
        color: #2C3E50;
        margin-bottom: 16px;
    }

    /* Department label */
    .department-label {
        color: #6C757D;
        font-size: 14px;
        margin-bottom: 8px;
    }

    /* Special schedule indicator */
    .special-schedule {
        color: #F59E0B;
        font-size: 14px;
        margin-top: 4px;
    }

    /* Interactive element animations */
    .stFileUploader, .stButton>button, .stSelectbox {
        transition: all 0.3s ease-in-out;
    }

    .stFileUploader:hover, .stButton>button:hover, .stSelectbox:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }

    /* Main content fade in */
    .main {
        animation: fadeIn 0.5s ease-in-out;
    }

    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(20px); }
        to { opacity: 1; transform: translateY(0); }
    }

    /* New hover menu styles */
    .hover-menu {
        display: none;
        position: absolute;
        left: 105%;
        top: 0;
        min-width: 300px;
        background: rgba(255, 255, 255, 0.95);
        backdrop-filter: blur(10px);
        border-radius: 12px;
        padding: 16px;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
        z-index: 1000;
        animation: slideIn 0.3s ease-out;
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

    .stat-card:hover .hover-menu {
        display: block;
    }

    /* Typewriter effect */
    @keyframes typewriter {
        from {
            width: 0;
        }
        to {
            width: 100%;
        }
    }

    .typewriter {
        overflow: hidden;
        white-space: nowrap;
        animation: typewriter 1s steps(40, end);
    }

    /* Pin button */
    .pin-button {
        position: absolute;
        top: 8px;
        right: 8px;
        background: transparent;
        border: none;
        color: #6C757D;
        cursor: pointer;
        transition: all 0.3s ease;
    }

    .pin-button:hover {
        color: #2196F3;
        transform: scale(1.1);
    }

    /* Daily details table */
    .daily-details {
        width: 100%;
        border-collapse: collapse;
        margin-top: 12px;
        font-size: 0.9em;
    }

    .daily-details th,
    .daily-details td {
        padding: 8px;
        text-align: left;
        border-bottom: 1px solid rgba(108, 117, 125, 0.2);
    }

    .daily-details th {
        font-weight: 600;
        color: #2C3E50;
    }

    .daily-details tr:hover {
        background: rgba(33, 150, 243, 0.05);
    }

    /* Enhanced info group */
    .info-group {
        position: relative;
    }

    /* Modern typography */
    .stat-value {
        font-family: 'SF Pro Display', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Open Sans', 'Helvetica Neue', sans-serif;
        letter-spacing: -0.5px;
    }
</style>

<script>
function togglePin(menuId) {
    const menu = document.getElementById(menuId);
    menu.classList.toggle('pinned');
    if (menu.classList.contains('pinned')) {
        menu.style.display = 'block';
    } else {
        menu.style.display = 'none';
    }
}
</script>
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
    daily_data = processor.get_employee_daily_data(employee_name) #added line to get daily data

    # Header with employee info
    schedule_note = "📅 Horario Especial" if stats.get('special_schedule', False) else ""
    st.markdown(f"""
        <div class="info-group">
            <h2>{stats['name']} {schedule_note}</h2>
            <div class="department-label">Departamento: {stats['department']}</div>
            {f'<div class="special-schedule">Empleado con horario especial</div>' if stats.get('special_schedule', False) else ''}
        </div>
    """, unsafe_allow_html=True)

    # Hours Summary - Full Width
    create_hours_summary(stats, daily_data) #Calling new function


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

def create_hours_summary(stats, daily_data):
    """Creates an enhanced hours summary card with hover details"""
    hours_ratio = (stats['actual_hours'] / stats['required_hours'] * 100) if stats['required_hours'] > 0 else 0
    hours_status = 'success' if hours_ratio >= 95 else 'warning' if hours_ratio >= 85 else 'danger'

    st.markdown(f"""
        <div class="info-group">
            <h3>📊 Resumen de Horas</h3>
            <div class="stat-card" style="position: relative;">
                <div class="metric-label">Horas Trabajadas</div>
                <div class="metric-value {hours_status} typewriter" style="font-size: 32px;">
                    {stats['actual_hours']:.1f}/{stats['required_hours']:.1f}
                </div>
                <div class="metric-label">({hours_ratio:.1f}%)</div>

                <div class="hover-menu" id="hoursMenu">
                    <button class="pin-button" onclick="togglePin('hoursMenu')">
                        📌
                    </button>
                    <h4>Detalles Diarios</h4>
                    <table class="daily-details">
                        <thead>
                            <tr>
                                <th>Fecha</th>
                                <th>Horas</th>
                                <th>Estado</th>
                            </tr>
                        </thead>
                        <tbody>
    """, unsafe_allow_html=True)

    for day_data in daily_data:
        date = day_data['date']
        hours = day_data['hours']
        status = 'success' if hours >= stats['required_hours'] else 'warning' if hours >= 0.8 * stats['required_hours'] else 'danger'
        st.markdown(f"""
                            <tr>
                                <td>{date}</td>
                                <td>{hours:.1f}</td>
                                <td class="{status}">{status}</td>
                            </tr>
        """, unsafe_allow_html=True)

    st.markdown("""
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)


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