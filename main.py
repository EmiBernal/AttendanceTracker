import streamlit as st
import pandas as pd
from utils.excel_processor import ExcelProcessor
from utils.visualizations import Visualizer
import plotly.graph_objects as go

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
    .dark-theme {
        background-color: #1E1E1E;
        color: #FFFFFF;
    }
    .light-theme {
        background-color: #FFFFFF;
        color: #262730;
    }
    .employee-stats {
        background-color: rgba(33, 150, 243, 0.1);
        border-radius: 10px;
        padding: 20px;
        margin: 10px 0;
    }
    .stat-card {
        background-color: #2D2D2D;
        border-radius: 8px;
        padding: 15px;
        margin: 5px;
        text-align: center;
    }
    .stSelectbox {
        width: 100%;
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

# Apply theme
if theme == "Dark":
    st.markdown("""
    <style>
        .stApp {
            background-color: #1E1E1E;
            color: #FFFFFF;
        }
        .css-1d391kg {
            background-color: #2D2D2D;
        }
        .stMarkdown {
            color: #FFFFFF;
        }
    </style>
    """, unsafe_allow_html=True)

# Load custom CSS
with open('assets/style.css') as f:
    st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

def safe_int_convert(value, default=0):
    """Safely convert value to integer, handling NaN and None"""
    if pd.isna(value) or value is None:
        return default
    try:
        return int(float(value))
    except (ValueError, TypeError):
        return default

def create_employee_dashboard(employee_data, records_list, visualizer):
    """Create a detailed dashboard for a single employee"""
    col1, col2 = st.columns([1, 3])

    with col1:
        # Profile section
        st.markdown(f"""
            <div class="employee-stats">
                <h2>{employee_data['Employee_Name']}</h2>
                <p style="color: #888;">Department: {employee_data['Department']}</p>
            </div>
        """, unsafe_allow_html=True)

        # Key metrics - only show if data exists
        metrics_html = ""

        # Hours
        if not pd.isna(employee_data.get('Actual_Hours')) and not pd.isna(employee_data.get('Required_Hours')):
            metrics_html += f"""
            <div class="stat-card">
                <h4>Hours</h4>
                <h2>{safe_int_convert(employee_data.get('Actual_Hours'))}/{safe_int_convert(employee_data.get('Required_Hours'))}</h2>
            </div>
            """

        # Late Minutes
        if not pd.isna(employee_data.get('Late_Minutes')):
            metrics_html += f"""
            <div class="stat-card">
                <h4>Late Minutes</h4>
                <h2>{safe_int_convert(employee_data.get('Late_Minutes'))}</h2>
            </div>
            """

        # Early Departures
        if not pd.isna(employee_data.get('Early_Departure_Minutes')):
            metrics_html += f"""
            <div class="stat-card">
                <h4>Early Departures</h4>
                <h2>{safe_int_convert(employee_data.get('Early_Departure_Minutes'))}</h2>
            </div>
            """

        if metrics_html:
            st.markdown(metrics_html, unsafe_allow_html=True)
        else:
            st.warning("No attendance data available for this employee")

    with col2:
        # Daily progress
        st.subheader("Daily Progress")
        employee_records = records_list[records_list['Employee_Name'] == employee_data['Employee_Name']]
        daily_chart = visualizer.create_employee_timeline(employee_records)
        st.plotly_chart(daily_chart, use_container_width=True)

        # Monthly statistics
        col_stats1, col_stats2 = st.columns(2)
        with col_stats1:
            st.subheader("Hours Distribution")
            hours_chart = visualizer.create_hours_distribution(employee_data)
            st.plotly_chart(hours_chart, use_container_width=True)

        with col_stats2:
            st.subheader("Attendance Analysis")
            attendance_chart = visualizer.create_attendance_stats(pd.Series(employee_data))
            st.plotly_chart(attendance_chart, use_container_width=True)

def main():
    st.title("📊 Interactive Attendance Visualizer")
    st.markdown("""
    <div class='custom-header'>
        Transform your Excel data into interactive visualizations
    </div>
    """, unsafe_allow_html=True)

    # File uploader in sidebar
    with st.sidebar:
        st.subheader("📂 Data Source")
        uploaded_file = st.file_uploader(
            "Upload Excel File",
            type=['xlsx', 'xls'],
            help="Upload the attendance Excel file with required sheets: Summary, Shifts, Logs, and Exceptional"
        )

    if uploaded_file:
        try:
            # Initialize processor and visualizer
            processor = ExcelProcessor(uploaded_file)
            visualizer = Visualizer()

            # Process data
            attendance_summary = processor.process_attendance_summary()
            shift_table = processor.process_shift_table()
            records_list = processor.process_records_list()
            individual_reports = processor.process_individual_reports()

            # Employee selector in sidebar
            with st.sidebar:
                st.subheader("👤 Employee Selection")
                selected_employee = st.selectbox(
                    "Select an employee to view their attendance details",
                    attendance_summary['Employee_Name'].unique(),
                    format_func=lambda x: x  # Display full name
                )

            # Create dashboard for selected employee
            employee_data = attendance_summary[
                attendance_summary['Employee_Name'] == selected_employee
            ].iloc[0].to_dict()  # Convert to dictionary to avoid index issues

            create_employee_dashboard(employee_data, records_list, visualizer)

            # Department overview tab
            with st.expander("📊 Department Overview"):
                col1, col2 = st.columns(2)
                with col1:
                    st.subheader("Department Attendance")
                    dept_chart = visualizer.create_department_chart(attendance_summary)
                    st.plotly_chart(dept_chart, use_container_width=True)

                with col2:
                    st.subheader("Department Network")
                    network = visualizer.create_department_network(attendance_summary)
                    st.plotly_chart(network, use_container_width=True)

        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            st.exception(e)  # This will show the full error trace in development

if __name__ == "__main__":
    main()