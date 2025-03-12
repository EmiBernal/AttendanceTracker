from datetime import datetime, timedelta
import streamlit as st
from streamlit_custom_notification_box import custom_notification_box

def check_employee_milestones(employee_stats):
    """
    Check if an employee has reached any milestones worth celebrating
    """
    milestones = []
    
    # Perfect Attendance Milestone
    if not employee_stats['absences'] and not employee_stats['late_days']:
        milestones.append({
            'title': '¡Asistencia Perfecta! 🌟',
            'message': f"¡Felicitaciones {employee_stats['name']}! Has mantenido una asistencia perfecta este mes.",
            'type': 'success'
        })
    
    # Overtime Achievement
    if employee_stats.get('overtime_minutes', 0) > 0:
        milestones.append({
            'title': 'Compromiso Destacado ⭐',
            'message': f"{employee_stats['name']} ha contribuido con horas extras este mes.",
            'type': 'info'
        })
    
    # PPP Hours Achievement
    if 'weekly_hours' in employee_stats:
        total_hours = sum(employee_stats['weekly_hours'].values())
        if total_hours >= 80:  # Meta mensual para PPP
            milestones.append({
                'title': '¡Meta Alcanzada! 🎯',
                'message': f"{employee_stats['name']} ha completado sus horas objetivo del mes.",
                'type': 'success'
            })
    
    return milestones

def show_celebration_popup(milestone):
    """
    Display an animated celebration popup for a milestone
    """
    styles = {
        'success': {
            'icon': '✅',
            'style': {
                'background-color': '#D4EDDA',
                'border-left-color': '#3C763D',
                'border-left-width': '10px',
                'border-left-style': 'solid',
                'color': '#3C763D'
            }
        },
        'info': {
            'icon': 'ℹ️',
            'style': {
                'background-color': '#CCE5FF',
                'border-left-color': '#004085',
                'border-left-width': '10px',
                'border-left-style': 'solid',
                'color': '#004085'
            }
        }
    }
    
    style = styles.get(milestone['type'], styles['info'])
    
    custom_notification_box(
        title=milestone['title'],
        message=milestone['message'],
        icon=style['icon'],
        type=milestone['type'],
        **style['style']
    )

def process_celebrations(employee_stats):
    """
    Process and display celebrations for an employee's achievements
    """
    milestones = check_employee_milestones(employee_stats)
    
    if milestones:
        st.markdown("""
        <style>
        @keyframes slideIn {
            from {
                transform: translateY(-100%);
                opacity: 0;
            }
            to {
                transform: translateY(0);
                opacity: 1;
            }
        }
        .stNotification {
            animation: slideIn 0.5s ease-out;
        }
        </style>
        """, unsafe_allow_html=True)
        
        for milestone in milestones:
            show_celebration_popup(milestone)
