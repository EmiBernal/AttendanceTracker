import plotly.graph_objects as go
import plotly.express as px
import networkx as nx
import streamlit as st
import pandas as pd

class Visualizer:
    def create_department_chart(self, df):
        dept_stats = df.groupby('Summary_Department').agg({
            'Required_Hours': 'sum',
            'Actual_Hours': 'sum'
        }).reset_index()

        fig = go.Figure()
        fig.add_trace(go.Bar(
            name='Required Hours',
            x=dept_stats['Summary_Department'],
            y=dept_stats['Required_Hours'],
            marker_color='#2196F3'
        ))
        fig.add_trace(go.Bar(
            name='Actual Hours',
            x=dept_stats['Summary_Department'],
            y=dept_stats['Actual_Hours'],
            marker_color='#4CAF50'
        ))

        fig.update_layout(
            barmode='group',
            plot_bgcolor='white',
            margin=dict(t=20, l=20, r=20, b=20)
        )
        return fig

    def create_employee_timeline(self, df):
        fig = go.Figure()

        for idx, row in df.iterrows():
            fig.add_trace(go.Scatter(
                x=[row['Entry_Time'], row['Exit_Time']],
                y=[idx, idx],
                mode='lines+markers',
                name=row['Name'],
                line=dict(color='#2196F3', width=2),
                marker=dict(size=8, symbol='circle')
            ))

        fig.update_layout(
            showlegend=False,
            plot_bgcolor='white',
            margin=dict(t=20, l=20, r=20, b=20),
            yaxis_title='Employee',
            xaxis_title='Time'
        )
        return fig

    def create_attendance_stats(self, df):
        stats = {
            'On Time': len(df[df['Late_Minutes'] == 0]),
            'Late': len(df[df['Late_Minutes'] > 0]),
            'Early Departure': len(df[df['Early_Minutes'] > 0])
        }

        fig = go.Figure(data=[go.Pie(
            labels=list(stats.keys()),
            values=list(stats.values()),
            hole=.3,
            marker_colors=['#4CAF50', '#FFC107', '#F44336']
        )])

        fig.update_layout(
            showlegend=True,
            margin=dict(t=20, l=20, r=20, b=20)
        )
        return fig

    def create_department_network(self, df):
        G = nx.Graph()
        departments = df['Summary_Department'].unique()

        # Create nodes
        for dept in departments:
            G.add_node(dept)

        # Create edges based on shared employees
        for i in range(len(departments)):
            for j in range(i+1, len(departments)):
                shared = len(df[df['Summary_Department'].isin([departments[i], departments[j]])])
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
                        plot_bgcolor='white'
                    ))
        return fig

    def create_employee_card(self, employee_data):
        st.markdown(
            f"""
            <div class="employee-card">
                <h3>{employee_data['Name']}</h3>
                <p>Department: {employee_data['Report_Department']}</p>
                <div class="stats">
                    <div class="stat">
                        <span class="label">Hours</span>
                        <span class="value">{employee_data['Actual_Hours']}/{employee_data['Required_Hours']}</span>
                    </div>
                    <div class="stat">
                        <span class="label">Late</span>
                        <span class="value">{employee_data['Late_Minutes']} min</span>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )