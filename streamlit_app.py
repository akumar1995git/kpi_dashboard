import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# Page setup with theme friendly styling
st.set_page_config(page_title="Employee KPI Dashboard - Modern View", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .stApp, .main, .block-container {background-color: #f5f7fa; color: #0E1117;}
    .stMetricLabel, .stMetricValue {color: #0E1117 !important;}
    .stButton>button {background-color: #2A5599 !important; color: white !important;}
    .stSelectbox>div>div>div>select, .stSelectbox>div>div>div>option {background-color:#d0e0fa; color:#0E1117;}
    .css-1r6slb0 {background-color:#dde9fa !important;}
    .stTabs [data-baseweb="tab-list"] button {background-color:#2A5599 !important; color:#fff; font-weight:bold;}
    .plotly-graph-div {background-color: white !important;}
    </style>
""", unsafe_allow_html=True)

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

uploaded_file = st.sidebar.file_uploader("Upload Updated_18_KPI_Dashboard.xlsx", type='xlsx')
df_path = uploaded_file if uploaded_file else "Updated_18_KPI_Dashboard.xlsx"

selected_kpis = st.sidebar.multiselect("Select KPIs to Display as Charts", employee_sheets, default=employee_sheets[:3])

dfs = {}
for kpi in selected_kpis:
    try:
        dfs[kpi] = pd.read_excel(df_path, sheet_name=kpi)
    except Exception as e:
        st.sidebar.error(f"Error loading {kpi}: {str(e)}")

# Get combined employee list
all_emps = sorted(set(emp for df in dfs.values() if 'Employee_ID' in df.columns for emp in df['Employee_ID']))

selected_employees = st.sidebar.multiselect("Select Employees to Analyze", all_emps, default=all_emps[:5])

def plot_sparkline(df, time_col, metric_col, employee):
    emp_df = df[df["Employee_ID"] == employee]
    if emp_df.empty:
        return None
    emp_df = emp_df.sort_values(time_col)
    fig = px.line(emp_df, x=time_col, y=metric_col, markers=True)
    fig.update_layout(
        height=160, margin=dict(l=10, r=10, t=20, b=30),
        xaxis=dict(showgrid=False, showticklabels=False),
        yaxis=dict(showgrid=False, showticklabels=False),
        template='plotly_white',
        font=dict(size=11)
    )
    return fig

st.title("Employee KPI Dashboard — Modern Multi-Chart Grid")

if not dfs:
    st.warning("Upload file or pick KPIs from sidebar to begin.")
    st.stop()

for kpi, df in dfs.items():
    st.subheader(kpi.replace("_", " "))
    time_candidate = next((c for c in df.columns if 'report' in c.lower() or 'week' in c.lower() or 'quarter' in c.lower()), None)
    if time_candidate is None or 'Employee_ID' not in df.columns:
        st.write("Unable to plot this KPI due to missing time or employee column.")
        continue

    metric_options = [c for c in df.columns if c not in ['Employee_ID', time_candidate] and pd.api.types.is_numeric_dtype(df[c])]
    if not metric_options:
        st.write("No numerical metric columns found for this KPI.")
        continue

    cols = st.columns(len(metric_options))
    for idx, metric in enumerate(metric_options):
        with cols[idx]:
            st.markdown(f"**{metric.replace('_', ' ')}**")

            for emp in selected_employees:
                fig = plot_sparkline(df, time_candidate, metric, emp)
                if fig:
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.write(f"No data for {emp}")

st.markdown("---")
with st.expander("Raw data view"):
    for kpi, df in dfs.items():
        st.subheader(kpi.replace("_", " "))
        st.dataframe(df[df['Employee_ID'].isin(selected_employees)] if 'Employee_ID' in df.columns else df, height=200)
