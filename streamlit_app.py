import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px

# Load the employee KPI sheets
employee_sheets = [
    'Role_vs_Reality_Analysis',
    'Hidden_Capacity_Burnout_Risk',
    'Work_Models_Effectiveness',
    'Digital_Collaboration_Overload',
    'Digital_Wellbeing_Index',
    'Data_Driven_Skill_Gap_Analysis',
    'High_Value_Work_Ratio',
    'Future_Skill_Readiness_Index',
    'Shadow_IT_Risk_Score'
]

st.set_page_config(
    page_title="Employee KPIs Multi-Chart Dashboard",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("Employee KPIs Dashboard — Multi-Chart View")

uploaded_file = st.sidebar.file_uploader("Upload Updated_18_KPI_Dashboard.xlsx", type="xlsx")
data_file = uploaded_file if uploaded_file else "Updated_18_KPI_Dashboard.xlsx"

selected_kpi = st.sidebar.multiselect("Select KPIs to Display", employee_sheets, default=employee_sheets[:3])

# Load data for selected KPIs
dfs = {}
for kpi in selected_kpi:
    try:
        df = pd.read_excel(data_file, sheet_name=kpi)
        dfs[kpi] = df
    except Exception as e:
        st.sidebar.error(f"Error loading sheet {kpi}: {e}")

# Employee selector considering selected KPIs
employee_sets = []
for kpi, df in dfs.items():
    if "Employee_ID" in df.columns:
        employee_sets.append(set(df["Employee_ID"].unique()))
employee_list = sorted(set.union(*employee_sets)) if employee_sets else []

selected_employees = st.sidebar.multiselect(
    "Select Employees",
    employee_list,
    default=employee_list[:5] if len(employee_list) > 5 else employee_list,
)

st.markdown("### KPI Grid\nSelect a KPI and Employee to get interactive mini-charts with trends.")

def plot_sparkline(df, time_col, metric_col, employee):
    emp_df = df[df["Employee_ID"] == employee]
    if emp_df.empty:
        return None
    emp_df = emp_df.sort_values(by=time_col)
    fig = px.line(emp_df, x=time_col, y=metric_col)
    fig.update_layout(
        height=150,
        margin=dict(l=20, r=20, t=20, b=20),
        xaxis=dict(showgrid=False, showticklabels=False),
        yaxis=dict(showgrid=False, showticklabels=False),
        showlegend=False
    )
    return fig

# For KPI list, pick a metric column for sparkline (numeric excluding Employee_ID/time)
def get_metric_cols(kpi_name):
    df = dfs[kpi_name]
    time_col = next((col for col in df.columns if "report" in col.lower() or "week" in col.lower() or "quarter" in col.lower()), None)
    exclude_cols = {"Employee_ID", time_col}
    return [c for c in df.select_dtypes(include="number").columns if c not in exclude_cols]

if not selected_kpi or not selected_employees:
    st.info("Please select at least one KPI and one employee on the sidebar to see visualizations.")
    st.stop()

for kpi in selected_kpi:
    st.subheader(f"{kpi.replace('_', ' ')}")
    metric_cols = get_metric_cols(kpi)
    if not metric_cols:
        st.write("No numeric metrics found in this KPI sheet.")
        continue

    for emp in selected_employees:
        cols = st.columns(len(metric_cols))
        for col, metric in zip(cols, metric_cols):
            fig = plot_sparkline(dfs[kpi], "Reporting_Period" if "Reporting_Period" in dfs[kpi].columns else next((c for c in dfs[kpi].columns if 'report' in c.lower()), None), metric, emp)
            if fig:
                avg_val = dfs[kpi][(dfs[kpi]['Employee_ID'] == emp)][metric].mean()
                trend_val = dfs[kpi][(dfs[kpi]['Employee_ID'] == emp)].sort_values(by="Reporting_Period" if "Reporting_Period" in dfs[kpi].columns else dfs[kpi].columns[0])[metric]
                trend_direction = "↑" if trend_val.iloc[-1] >= trend_val.iloc[0] else "↓"
                col.markdown(f"**{metric.replace('_', ' ')}**  \nEmployee: {emp}  \nAvg: {avg_val:.2f} {trend_direction}")
                col.plotly_chart(fig, use_container_width=True)
            else:
                col.write(f'No data for {emp} on {metric}')

st.markdown("---")
st.markdown("_Use sidebar filters to add/remove KPIs and employees. Hover on mini-charts for detailed view._")
