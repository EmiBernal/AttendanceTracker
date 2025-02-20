import plotly.graph_objects as go
import plotly.express as px
import networkx as nx
import streamlit as st
import pandas as pd
import numpy as np

class Visualizer:
    def create_department_chart(self, df):
        dept_stats = df.groupby('Department').agg({
            'Required_Hours': 'sum',
            'Actual_Hours': 'sum'
        }).reset_index()

        fig = go.Figure()
        fig.add_trace(go.Bar(
            name='Required Hours',
            x=dept_stats['Department'],
            y=dept_stats['Required_Hours'],
            marker_color='#2196F3'
        ))
        fig.add_trace(go.Bar(
            name='Actual Hours',
            x=dept_stats['Department'],
            y=dept_stats['Actual_Hours'],
            marker_color='#4CAF50'
        ))

        fig.update_layout(
            barmode='group',
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            margin=dict(t=20, l=20, r=20, b=20),
            font=dict(color='#FFFFFF')
        )
        return fig

    def create_employee_timeline(self, df):
        fig = go.Figure()

        colors = px.colors.qualitative.Set3

        for idx, row in df.iterrows():
            times = [
                row['Initial_Entry'],
                row['Midday_Exit'],
                row['Midday_Entry'],
                row['Final_Exit']
            ]
            times = [t for t in times if pd.notna(t)]

            if len(times) > 1:
                fig.add_trace(go.Scatter(
                    x=times,
                    y=[row['Employee_Name']] * len(times),
                    mode='lines+markers',
                    name=row['Employee_Name'],
                    line=dict(color=colors[idx % len(colors)], width=2),
                    marker=dict(size=8, symbol='circle')
                ))

        fig.update_layout(
            showlegend=False,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            margin=dict(t=20, l=20, r=20, b=20),
            yaxis_title='Time Records',
            xaxis_title='Time',
            font=dict(color='#FFFFFF')
        )
        return fig

    def create_hours_distribution(self, employee_data):
        labels = ['Required', 'Actual', 'Overtime']
        values = [
            float(employee_data['Required_Hours']),
            float(employee_data['Actual_Hours']),
            float(employee_data['Normal_Overtime']) + float(employee_data['Special_Overtime'])
        ]

        colors = ['#2196F3', '#4CAF50', '#FFC107']

        fig = go.Figure(data=[go.Pie(
            labels=labels,
            values=values,
            hole=.3,
            marker_colors=colors
        )])

        fig.update_layout(
            showlegend=True,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            margin=dict(t=20, l=20, r=20, b=20),
            font=dict(color='#FFFFFF')
        )
        return fig

    def create_attendance_stats(self, df):
        stats = {
            'On Time': int(len(df[df['Late_Minutes'] == 0])),
            'Late': int(len(df[df['Late_Minutes'] > 0])),
            'Early Departure': int(len(df[df['Early_Departure_Minutes'] > 0]))
        }

        colors = ['#4CAF50', '#FFC107', '#F44336']

        fig = go.Figure(data=[go.Pie(
            labels=list(stats.keys()),
            values=list(stats.values()),
            hole=.3,
            marker_colors=colors
        )])

        fig.update_layout(
            showlegend=True,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            margin=dict(t=20, l=20, r=20, b=20),
            font=dict(color='#FFFFFF')
        )
        return fig

    def create_department_network(self, df):
        G = nx.Graph()
        departments = df['Department'].unique()

        # Create nodes
        for dept in departments:
            G.add_node(dept)

        # Create edges based on shared employees
        for i in range(len(departments)):
            for j in range(i+1, len(departments)):
                shared = len(df[df['Department'].isin([departments[i], departments[j]])])
                if shared > 0:
                    G.add_edge(departments[i], departments[j], weight=shared)

        pos = nx.spring_layout(G)

        edge_trace = go.Scatter(
            x=[], y=[],
            line=dict(width=0.5, color='#888'),
            hoverinfo='none',
            mode='lines')

        node_trace = go.Scatter(
            x=[], y=[],
            mode='markers+text',
            hoverinfo='text',
            marker=dict(
                size=20,
                color='#2196F3',
                line_width=2))

        for edge in G.edges():
            x0, y0 = pos[edge[0]]
            x1, y1 = pos[edge[1]]
            edge_trace['x'] += (x0, x1, None)
            edge_trace['y'] += (y0, y1, None)

        node_trace['x'] = [pos[node][0] for node in G.nodes()]
        node_trace['y'] = [pos[node][1] for node in G.nodes()]
        node_trace['text'] = list(G.nodes())

        fig = go.Figure(data=[edge_trace, node_trace],
                     layout=go.Layout(
                        showlegend=False,
                        hovermode='closest',
                        margin=dict(t=20, l=20, r=20, b=20),
                        plot_bgcolor='rgba(0,0,0,0)',
                        paper_bgcolor='rgba(0,0,0,0)',
                        font=dict(color='#FFFFFF')
                    ))
        return fig

    def create_employee_card(self, employee_data):
        card_style = """
        background-color: rgba(33, 150, 243, 0.1);
        border: 1px solid #2196F3;
        border-radius: 10px;
        padding: 15px;
        margin: 10px 0;
        """

        st.markdown(
            f"""
            <div style="{card_style}">
                <h3 style="color: #2196F3; margin: 0;">{employee_data['Employee_Name']}</h3>
                <p style="color: #888;">Department: {employee_data['Department']}</p>
                <div style="display: flex; justify-content: space-between; margin-top: 10px;">
                    <div>
                        <span style="color: #888; font-size: 12px;">Hours</span><br>
                        <span style="color: #FFFFFF;">{int(employee_data['Actual_Hours'])}/{int(employee_data['Required_Hours'])}</span>
                    </div>
                    <div>
                        <span style="color: #888; font-size: 12px;">Late</span><br>
                        <span style="color: #FFFFFF;">{int(employee_data['Late_Minutes'])} min</span>
                    </div>
                    <div>
                        <span style="color: #888; font-size: 12px;">Early Leave</span><br>
                        <span style="color: #FFFFFF;">{int(employee_data['Early_Leave_Minutes'])} min</span>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )