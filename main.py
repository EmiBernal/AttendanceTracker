import streamlit as st
import pandas as pd
from utils.excel_processor import ExcelProcessor
from utils.visualizations import Visualizer
import plotly.graph_objects as go

# Page configuration with dark theme
st.set_page_config(
    page_title="Attendance Visualizer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Theme toggle in sidebar
theme = st.sidebar.selectbox(
    "Choose Theme",
    ["Light", "Dark"],
    key="theme_selector"
)

# Apply theme
if theme == "Dark":
    dark_theme = """
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
    </style>
    """
    st.markdown(dark_theme, unsafe_allow_html=True)

# Load custom CSS
with open('assets/style.css') as f:
    st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

def create_employee_dashboard(employee_data, visualizer):
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

        # Key metrics
        metrics_html = f"""
        <div class="stat-card">
            <h4>Hours</h4>
            <h2>{int(employee_data['Actual_Hours'])}/{int(employee_data['Required_Hours'])}</h2>
        </div>
        <div class="stat-card">
            <h4>Late Minutes</h4>
            <h2>{int(employee_data['Late_Minutes'])}</h2>
        </div>
        <div class="stat-card">
            <h4>Early Departures</h4>
            <h2>{int(employee_data['Early_Departure_Minutes'])}</h2>
        </div>
        """
        st.markdown(metrics_html, unsafe_allow_html=True)

    with col2:
        # Daily progress
        st.subheader("Daily Progress")
        daily_chart = visualizer.create_employee_timeline(
            pd.DataFrame([employee_data])
        )
        st.plotly_chart(daily_chart, use_container_width=True)

        # Monthly statistics
        col_stats1, col_stats2 = st.columns(2)
        with col_stats1:
            st.subheader("Hours Distribution")
            hours_chart = visualizer.create_hours_distribution(employee_data)
            st.plotly_chart(hours_chart, use_container_width=True)

        with col_stats2:
            st.subheader("Attendance Analysis")
            attendance_chart = visualizer.create_attendance_stats(
                pd.DataFrame([employee_data])
            )
            st.plotly_chart(attendance_chart, use_container_width=True)

def main():
    st.title("📊 Interactive Attendance Visualizer")
    st.markdown("""
    <div class='custom-header'>
        Transform your Excel data into interactive visualizations
    </div>
    """, unsafe_allow_html=True)

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
            selected_employee = st.sidebar.selectbox(
                "Select Employee",
                attendance_summary['Employee_Name'].unique()
            )

            # Create dashboard for selected employee
            employee_data = attendance_summary[
                attendance_summary['Employee_Name'] == selected_employee
            ].iloc[0]
            create_employee_dashboard(employee_data, visualizer)

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