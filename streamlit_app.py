import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Employee-centric KPI sheets
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

st.set_page_config(page_title="Employee KPI Dashboard", layout="wide", initial_sidebar_state="expanded")
st.markdown(
    """
    <style>
    .main {
        background-color: #EEF6FA;
    }
    .stApp {
        background-color: #ECF4F9;
    }
    .css-1d391kg { background: #357ABD !important; }
    </style>
    """, unsafe_allow_html=True
)

st.title("🔎 Employee KPI Dashboard")
st.markdown("##### View and analyze detailed, time-trended employee metrics with rich visualizations.")

excel_file = "Updated_18_KPI_Dashboard.xlsx"

selected_kpi = st.sidebar.selectbox("📊 Select Employee KPI", employee_sheets)
df = pd.read_excel(excel_file, sheet_name=selected_kpi)

# Sidebar filter: show employees
emp_col = "Employee_ID" if "Employee_ID" in df.columns else None
if emp_col:
    emp_list = df[emp_col].unique()
    selected_emp = st.sidebar.selectbox("👤 Select Employee", emp_list)

    emp_df = df[df[emp_col] == selected_emp]
    st.subheader(f"Details for Employee: `{selected_emp}`")
else:
    emp_df = df.copy()

# Time columns auto-detection
time_col = next((col for col in df.columns if 'report' in col.lower() or 'week' in col.lower() or 'quarter' in col.lower()), None)
numeric_cols = emp_df.select_dtypes(include='number').columns.tolist()
metric_cols = [col for col in numeric_cols if col != time_col]

# Layout - 3 columns for top metrics
col1, col2, col3 = st.columns([1, 1, 1])
for i, metric in enumerate(metric_cols[:3]):
    val = emp_df[metric].mean() if not emp_df.empty else None
    with [col1, col2, col3][i]:
        st.metric(label=f"{metric.replace('_',' ').title()}", value=f"{val:,.2f}" if val is not None else '')

st.markdown("---")

# Main Visualization - Multi-graph display
st.subheader("Employee KPI Trends & Distributions")

tab1, tab2, tab3 = st.tabs(["Trend Analysis", "Distribution", "Employee Comparison"])

with tab1:
    if time_col and metric_cols:
        metric_to_plot = st.selectbox("Select Metric to Plot (Trend)", metric_cols, key="trend_metric")
        fig = px.line(emp_df, x=time_col, y=metric_to_plot,
                      markers=True,
                      title=f"{metric_to_plot.replace('_',' ')} over Time",
                      color_discrete_sequence=px.colors.qualitative.Plotly)
        fig.update_layout(template='plotly_white', title_x=0.5)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Trend chart not available for this metric.")

with tab2:
    dist_metric = st.selectbox("Select Metric for Distribution", metric_cols, key="dist_metric")
    fig = px.histogram(emp_df, x=dist_metric,
                       nbins=20,
                       title=f"{dist_metric.replace('_',' ')} Distribution",
                       color_discrete_sequence=['#357ABD'])
    fig.update_layout(template='plotly_white', title_x=0.5)
    st.plotly_chart(fig, use_container_width=True)

with tab3:
    if emp_col and metric_cols:
        comp_metric = st.selectbox("Compare Employees by Metric", metric_cols, key="comp_metric")
        comp_df = df[[emp_col, time_col, comp_metric]].dropna()
        fig = px.box(comp_df, x=emp_col, y=comp_metric,
                     color=emp_col,
                     title=f"{comp_metric.replace('_',' ')} Across Employees",
                     color_discrete_sequence=px.colors.sequential.Blues)
        fig.update_layout(template='plotly_white', title_x=0.5)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No employee comparison available.")

st.markdown("---")
st.subheader("Raw KPI Data Preview")
st.dataframe(emp_df, use_container_width=True, height=250)

st.success("Tip: Use the sidebar to choose KPIs or filter employees. Click tabs above for more views.\n")

st.markdown(
    "<style>div.stTabs>div>button[data-baseweb='tab']{background:#D2E2F4;color:#05386B;}</style>",
    unsafe_allow_html=True
)

st.markdown("Data source: [Updated_18_KPI_Dashboard.xlsx](link to your repo)")
