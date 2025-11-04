import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# -----------------------------
# Setup Page Layout & Styling
# -----------------------------
st.set_page_config(
    page_title="Professional Employee KPI Dashboard",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for dual dark/light mode readability
st.markdown(
    """
    <style>
    /* Background & text colors for whole app */
    .stApp, .main, .block-container {
        background-color: #f0f2f6 !important;
        color: #0e1117 !important;
    }

    /* Metric widget text colors */
    .stMetricLabel, .stMetricValue {
        color: #0e1117 !important;
    }

    /* Tabs styling with accessible blue & white text */
    .stTabs [data-baseweb="tab-list"] button {
        background-color: #1f77b4 !important;
        color: white !important;
        font-weight: 600;
    }
    .stTabs [data-baseweb="tab-list"] button:focus, 
    .stTabs [data-baseweb="tab-list"] button:hover {
        background-color: #155d8b !important;
        color: white !important;
    }

    /* Plotly chart background for contrast */
    .plotly-graph-div {
        background-color: white !important;
    }

    /* Sidebar background for uniformity */
    .css-1d391kg {
        background-color: #e3e8f0 !important;
        color: #0e1117 !important;
    }
    
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------
# Define Employee-centric KPIs
# -----------------------------
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

# -----------------------------
# Application Title & Text
# -----------------------------
st.title("🔎 Employee KPI Dashboard")
st.markdown("Explore detailed trends and actionable insights on employee performance and well-being.")

# -----------------------------
# Sidebar Filters
# -----------------------------
uploaded_file = st.sidebar.file_uploader("Upload Updated_18_KPI_Dashboard.xlsx", type=['xlsx'])
if uploaded_file:
    excel_path = uploaded_file
else:
    excel_path = "Updated_18_KPI_Dashboard.xlsx"  # Change to your filepath or repo URL if needed

selected_kpi = st.sidebar.selectbox("Select Employee KPI", employee_sheets)

try:
    df = pd.read_excel(excel_path, sheet_name=selected_kpi)
except Exception as e:
    st.error(f"Error loading KPI data: {e}")
    st.stop()

# Filter by Employee if available
emp_col = "Employee_ID"
if emp_col in df.columns:
    employees = sorted(df[emp_col].unique())
    selected_emp = st.sidebar.selectbox("Select Employee", ["All"] + employees)
    if selected_emp != "All":
        df = df[df[emp_col] == selected_emp]

# -----------------------------
# Detect columns for display
# -----------------------------
time_col = next((c for c in df.columns if 'report' in c.lower() or 'week' in c.lower() or 'quarter' in c.lower()), None)
numeric_cols = df.select_dtypes(include=np.number).columns.to_list()
metric_cols = [c for c in numeric_cols if c != time_col]

# -----------------------------
# Display key averages as metrics
# -----------------------------
st.subheader("Key Metrics Summary")
cols = st.columns(min(4, len(metric_cols)))

def detect_metric_color(metric_name: str, avg_val: float):
    metric_name = metric_name.lower()
    # Metrics with high value = good
    if 'risk' in metric_name or 'overload' in metric_name or 'burnout' in metric_name:
        # Lower is better
        if avg_val <= 30:
            return "🟢"
        elif avg_val <=50:
            return "🟡"
        else:
            return "🔴"
    else:
        # Higher is better
        if avg_val >= 70:
            return "🟢"
        elif avg_val >= 50:
            return "🟡"
        else:
            return "🔴"

for i, metric in enumerate(metric_cols[:4]):
    avg_val = df[metric].mean()
    emoji = detect_metric_color(metric, avg_val)
    with cols[i]:
        st.metric(label=metric.replace('_',' ').title(),
                  value=f"{avg_val:.2f}",
                  delta_color="normal")
        st.markdown(f"{emoji} <sub>Average over period</sub>", unsafe_allow_html=True)

# -----------------------------
# Visualization Tabs
# -----------------------------
st.markdown("---")
st.subheader("Visualizations")

tab_trend, tab_dist, tab_comp, tab_insights = st.tabs(["Trend", "Distribution", "Employee Comparison", "Insights & Suggestions"])

with tab_trend:
    if time_col and metric_cols:
        chosen_metric = st.selectbox("Select Metric to Plot Over Time", metric_cols, key="trend_metric")
        fig = px.line(df, x=time_col, y=chosen_metric, title=f"{chosen_metric.replace('_',' ').title()} Over Time",
                      markers=True, color_discrete_sequence=px.colors.qualitative.Safe)
        fig.update_layout(template='plotly_white', title_x=0.5)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No time-series data available to plot.")

with tab_dist:
    if metric_cols:
        chosen_metric = st.selectbox("Select Metric for Distribution", metric_cols, key="dist_metric")
        fig = px.histogram(df, x=chosen_metric, nbins=20, title=f"Distribution of {chosen_metric.replace('_',' ').title()}",
                           color_discrete_sequence=['#377eb8'])
        fig.update_layout(template='simple_white', title_x=0.5)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No numeric fields available for distribution.")

with tab_comp:
    if emp_col in df.columns and metric_cols:
        chosen_metric = st.selectbox("Compare Employees by Metric", metric_cols, key="comp_metric")
        fig = px.box(df, x=emp_col, y=chosen_metric, color=emp_col, title=f"{chosen_metric.replace('_',' ').title()} Distribution Across Employees",
                     color_discrete_sequence=px.colors.qualitative.Dark24)
        fig.update_layout(template='plotly_white', showlegend=False, title_x=0.5)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Employee ID or metric data not available for comparison.")

def generate_insights(metric):
    vals = df[metric].dropna()
    if len(vals) < 3:
        return "Insufficient data for insights."
    trend = np.polyfit(range(len(vals)), vals, 1)[0]
    improvement_direction = "improved" if trend > 0 else "declined"
    avg_val = np.mean(vals)
    msg = f"The average {metric.replace('_',' ')} is {avg_val:.2f} and it has {improvement_direction} over time.\n"
    if trend > 0:
        msg += "Keep up the good work and monitor key drivers for continuous improvement."
    else:
        msg += "Consider interventions to address declining performance."
    return msg

with tab_insights:
    st.subheader("Automated Insights & Suggestions")
    for metric in metric_cols[:6]:
        st.markdown(f"### {metric.replace('_',' ').title()}")
        st.write(generate_insights(metric))
        st.markdown("---")

# -----------------------------
# Raw Data Preview
# -----------------------------
st.markdown("---")
with st.expander("View Complete KPI Raw Data", expanded=False):
    st.dataframe(df, use_container_width=True, height=300)

st.markdown("Completed by Streamlit | Data Source: Updated_18_KPI_Dashboard.xlsx")
