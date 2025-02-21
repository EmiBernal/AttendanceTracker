import streamlit as st
import pandas as pd
from utils.excel_processor import ExcelProcessor
from utils.visualizations import Visualizer

# Page configuration
st.set_page_config(
    page_title="Attendance Visualizer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for theme toggle and layout
st.markdown("""
<style>
    .theme-toggle {
        position: fixed;
        top: 0.5rem;
        right: 1rem;
        z-index: 1000;
        background: transparent;
        border: none;
        cursor: pointer;
        padding: 0.5rem;
        font-size: 1.5rem;
    }
    .stat-group {
        background-color: rgba(33, 150, 243, 0.1);
        border-radius: 10px;
        padding: 15px;
        margin: 10px 0;
    }
    .stat-card {
        background-color: #2D2D2D;
        border-radius: 8px;
        padding: 12px;
        margin: 5px;
        text-align: center;
    }
    .metric-value {
        font-size: 24px;
        font-weight: bold;
        margin: 5px 0;
    }
    .metric-label {
        font-size: 14px;
        color: #888;
    }
    .warning {
        color: #FFC107;
    }
    .danger {
        color: #F44336;
    }
    .success {
        color: #4CAF50;
    }
</style>
""", unsafe_allow_html=True)

# Theme toggle in header
theme = st.session_state.get('theme', 'Light')
col_spacer, col_toggle = st.columns([6, 1])
with col_toggle:
    if st.button('🌞' if theme == 'Dark' else '🌙', key='theme_toggle'):
        theme = 'Light' if theme == 'Dark' else 'Dark'
        st.session_state.theme = theme

def safe_float_convert(value, default=0):
    """Convierte de manera segura a float, manejando NaN y valores no válidos."""
    if pd.isna(value) or value is None:
        return default
    try:
        return round(float(value), 1)
    except (ValueError, TypeError):
        return default

def create_employee_dashboard(processor, employee_name):
    """Create a detailed dashboard for a single employee"""
    stats = processor.get_employee_stats(employee_name)

    # Header with employee info
    st.markdown(f"""
        <div class="stat-group">
            <h2>{stats['name']}</h2>
            <p style="color: #888;">Department: {stats['department']}</p>
        </div>
    """, unsafe_allow_html=True)

    # Work Hours Overview
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
            <div class="stat-group">
                <h3>📊 Work Hours Overview</h3>
        """, unsafe_allow_html=True)

        hours_ratio = (stats['actual_hours'] / stats['required_hours'] * 100) if stats['required_hours'] > 0 else 0
        hours_status = 'success' if hours_ratio >= 95 else 'warning' if hours_ratio >= 85 else 'danger'

        st.markdown(f"""
            <div class="stat-card">
                <div class="metric-label">Hours Worked</div>
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
                <h3>📋 Attendance Summary</h3>
        """, unsafe_allow_html=True)

        attendance_ratio = float(stats['attendance_ratio']) * 100
        attendance_status = 'success' if attendance_ratio >= 95 else 'warning' if attendance_ratio >= 85 else 'danger'

        st.markdown(f"""
            <div class="stat-card">
                <div class="metric-label">Attendance Rate</div>
                <div class="metric-value {attendance_status}">{attendance_ratio:.1f}%</div>
            </div>
        """, unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    # Detailed Statistics
    st.markdown("""
        <div class="stat-group">
            <h3>📈 Detailed Statistics</h3>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px;">
    """, unsafe_allow_html=True)

    # Function to determine status color
    def get_status_color(value, threshold_warning=1, threshold_danger=3):
        return 'success' if value == 0 else 'warning' if value <= threshold_warning else 'danger'

    # List of metrics to display
    metrics = [
        ('Late Days', stats['late_days'], f"{stats['late_minutes']:.0f} minutes total"),
        ('Early Departures', stats['early_departures'], f"{stats['early_minutes']:.0f} minutes total"),
        ('Extended Lunch Breaks', stats['extended_lunch_days'], "Over 20 minutes"),
        ('Missing Records', stats['missing_record_days'], "Incomplete entries"),
        ('Absences', stats['absences'], "Total days")
    ]

    for label, value, subtitle in metrics:
        status = get_status_color(value)
        st.markdown(f"""
            <div class="stat-card">
                <div class="metric-label">{label}</div>
                <div class="metric-value {status}">{value}</div>
                <div class="metric-label">{subtitle}</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("</div></div>", unsafe_allow_html=True)

def main():
    st.title("📊 Interactive Attendance Visualizer")

    # File uploader in sidebar
    with st.sidebar:
        st.subheader("📂 Data Source")
        uploaded_file = st.file_uploader(
            "Upload Excel File",
            type=['xlsx', 'xls'],
            help="Upload the attendance Excel file with required sheets"
        )

    if uploaded_file:
        try:
            processor = ExcelProcessor(uploaded_file)
            attendance_summary = processor.process_attendance_summary()

            # Employee selector in sidebar
            with st.sidebar:
                st.subheader("👤 Employee Selection")
                selected_employee = st.selectbox(
                    "Select an employee to view their attendance details",
                    attendance_summary['employee_name'].unique(),
                    format_func=lambda x: x  # Display full name
                )

            # Create dashboard for selected employee
            create_employee_dashboard(processor, selected_employee)

        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            st.exception(e)

if __name__ == "__main__":
    main()