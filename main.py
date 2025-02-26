import streamlit as st
from utils.excel_processor import ExcelProcessor

# Page configuration
st.set_page_config(
    page_title="Visualizador de Asistencia",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Updated CSS with sliding transitions and animations
st.markdown("""
<style>
    /* Base transitions */
    .stApp {
        transition: all 0.3s ease-in-out;
    }

    /* Container for all views */
    .view-container {
        position: relative;
        width: 100%;
        overflow: hidden;
    }

    /* Main view and detail view styling */
    .main-view, .detail-view {
        width: 100%;
        transition: transform 0.5s cubic-bezier(0.4, 0, 0.2, 1);
    }

    /* Slide animations */
    .slide-out-left {
        transform: translateX(-100%);
    }

    .slide-in-right {
        transform: translateX(0);
    }

    .slide-out-right {
        transform: translateX(100%);
    }

    .slide-in-left {
        transform: translateX(0);
    }

    /* Card styling */
    .stat-card {
        background-color: rgba(255, 255, 255, 0.05);
        border-radius: 10px;
        padding: 20px;
        margin: 12px;
        cursor: pointer;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        border: 1px solid rgba(233, 236, 239, 0.2);
    }

    .stat-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 24px rgba(0,0,0,0.12);
    }

    /* Typewriter effect */
    @keyframes typewriter {
        from { width: 0; }
        to { width: 100%; }
    }

    .typewriter {
        overflow: hidden;
        white-space: nowrap;
        animation: typewriter 1s steps(40, end);
    }

    /* Detail view table */
    .detail-table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 20px;
        font-family: 'SF Pro Display', system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }

    .detail-table th,
    .detail-table td {
        padding: 12px;
        text-align: left;
        border-bottom: 1px solid rgba(233, 236, 239, 0.2);
    }

    .detail-table th {
        background-color: rgba(33, 150, 243, 0.1);
        font-weight: 500;
    }

    .detail-table tr:hover {
        background-color: rgba(33, 150, 243, 0.05);
    }

    /* Back button */
    .back-button {
        display: inline-flex;
        align-items: center;
        padding: 8px 16px;
        border: none;
        border-radius: 8px;
        background: rgba(33, 150, 243, 0.1);
        color: #2196F3;
        font-family: 'SF Pro Display', system-ui;
        font-weight: 500;
        cursor: pointer;
        transition: all 0.3s ease;
        margin-bottom: 20px;
    }

    .back-button:hover {
        background: rgba(33, 150, 243, 0.2);
        transform: translateX(-4px);
    }

    /* Status colors */
    .warning { color: #FFC107; }
    .danger { color: #DC3545; }
    .success { color: #28A745; }
</style>
<script>
function showDetailView(cardId) {
    const mainView = document.querySelector('.main-view');
    const detailView = document.querySelector('.detail-view-' + cardId);

    mainView.classList.add('slide-out-left');
    detailView.classList.remove('slide-out-right');
    detailView.classList.add('slide-in-right');
}

function hideDetailView(cardId) {
    const mainView = document.querySelector('.main-view');
    const detailView = document.querySelector('.detail-view-' + cardId);

    mainView.classList.remove('slide-out-left');
    detailView.classList.remove('slide-in-right');
    detailView.classList.add('slide-out-right');
}
</script>
""", unsafe_allow_html=True)

def create_hours_detail_view(stats, daily_data):
    """Creates a detailed view for hours worked"""
    st.markdown("""
        <div class="detail-view detail-view-hours slide-out-right">
            <button class="back-button" onclick="hideDetailView('hours')">
                ← Volver
            </button>
            <h2>Detalle de Horas Trabajadas</h2>
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
        status = 'success' if day['hours'] >= stats['required_hours']/20 else 'warning' if day['hours'] >= stats['required_hours']/25 else 'danger'
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

    # Main view container
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

    # Hours Summary Card
    hours_ratio = (stats['actual_hours'] / stats['required_hours'] * 100) if stats['required_hours'] > 0 else 0
    hours_status = 'success' if hours_ratio >= 95 else 'warning' if hours_ratio >= 85 else 'danger'

    st.markdown(f"""
        <div class="stat-card" onclick="showDetailView('hours')">
            <div class="metric-label">Horas Trabajadas</div>
            <div class="metric-value {hours_status} typewriter">
                {stats['actual_hours']:.1f}/{stats['required_hours']:.1f}
            </div>
            <div class="metric-label">({hours_ratio:.1f}%)</div>
        </div>
    """, unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)  # Close main-view

    # Create detail view for hours
    create_hours_detail_view(stats, daily_data)

    st.markdown("</div>", unsafe_allow_html=True)  # Close view-container

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