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
    </style>
    """
    st.markdown(dark_theme, unsafe_allow_html=True)

# Load custom CSS
with open('assets/style.css') as f:
    st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

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

            # Dashboard tabs
            tab1, tab2, tab3, tab4 = st.tabs([
                "📊 Department Overview",
                "⏰ Time Records",
                "👥 Employee Analysis",
                "📋 Individual Reports"
            ])

            with tab1:
                col1, col2 = st.columns(2)
                with col1:
                    st.subheader("Department Attendance")
                    dept_chart = visualizer.create_department_chart(attendance_summary)
                    st.plotly_chart(dept_chart, use_container_width=True)

                with col2:
                    st.subheader("Attendance Statistics")
                    stats_chart = visualizer.create_attendance_stats(attendance_summary)
                    st.plotly_chart(stats_chart, use_container_width=True)

            with tab2:
                st.subheader("Daily Time Records")
                timeline = visualizer.create_employee_timeline(records_list)
                st.plotly_chart(timeline, use_container_width=True)

            with tab3:
                st.subheader("Department Network")
                network = visualizer.create_department_network(attendance_summary)
                st.plotly_chart(network, use_container_width=True)

            with tab4:
                st.subheader("Employee Cards")
                employee_cols = st.columns(4)
                for idx, employee in enumerate(individual_reports.head(4).iterrows()):
                    with employee_cols[idx]:
                        visualizer.create_employee_card(employee[1])

        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            st.exception(e)  # This will show the full error trace in development

if __name__ == "__main__":
    main()