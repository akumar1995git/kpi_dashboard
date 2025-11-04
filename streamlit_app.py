import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

# Setup and Styling (same as before)
st.set_page_config(page_title="Refined Employee KPI Dashboard", layout="wide", initial_sidebar_state="expanded")

st.markdown(
    """
    <style>
    .stApp, .main, .block-container {background-color: #f7f9fc; color:#101820;}
    .stMetricLabel, .stMetricValue {color: #101820 !important;}
    .stButton>button {background-color: #005b96 !important; color: white !important;}
    .stSelectbox>div>div>div>select, .stSelectbox>div>div>div>option {background-color:#c4d1e8; color:#101820;}
    .css-1r6slb0 {background-color: #e6efff !important;}
    .stTabs [data-baseweb="tab-list"] button {background:#005b96 !important;color:#fff;font-weight:bold;}
    .plotly-graph-div {background-color:white !important;}
    </style>
    """,
    unsafe_allow_html=True,
)

# Employee KPI sheets list
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

# Sidebar inputs and excel loading
uploaded_file = st.sidebar.file_uploader("Upload Updated_18_KPI_Dashboard.xlsx", type=['xlsx'])
excel_path = uploaded_file if uploaded_file else "Updated_18_KPI_Dashboard.xlsx"

selected_kpi = st.sidebar.selectbox("Select Employee KPI", employee_sheets)
df = pd.read_excel(excel_path, sheet_name=selected_kpi)

emp_col = "Employee_ID" if "Employee_ID" in df.columns else None
if emp_col:
    emp_list = sorted(df[emp_col].unique())
    selected_emp = st.sidebar.selectbox("Select Employee", ['All'] + emp_list)
    if selected_emp != 'All':
        df = df[df[emp_col] == selected_emp]

time_col_candidates = [col for col in df.columns if any(x in col.lower() for x in ["report", "week", "date", "quarter"])]
time_col = time_col_candidates[0] if time_col_candidates else None

numeric_cols = df.select_dtypes(include=np.number).columns.tolist()
metric_cols = [col for col in numeric_cols if col != time_col]

# Title & metrics summary
st.title("Refined Employee KPI Dashboard")

st.subheader("Summary Metrics")
for metric in metric_cols[:4]:
    val = df[metric].mean()
    st.metric(label=metric.replace('_',' ').title(), value=f"{val:.2f}")

st.markdown("---")

# Visualization Tabs - pie chart removed
tab1, tab2, tab3, tab4 = st.tabs(["Line Chart", "Bar Chart", "Heatmap", "Box Plot"])

with tab1:
    st.subheader("Metric Trends Over Time")
    if time_col:
        metric_line = st.selectbox("Select Metric", metric_cols, key="line_metric")
        fig = px.line(df, x=time_col, y=metric_line, title=f"{metric_line.replace('_',' ')} Over Time",
                      markers=True, color_discrete_sequence=px.colors.sequential.Plasma)
        fig.update_layout(template="plotly_white", title_x=0.5)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Time column not identified for trend chart")

with tab2:
    st.subheader("Sum/Average Per Employee Bar Chart")
    metric_bar = st.selectbox("Select Numeric Metric", metric_cols, key="bar_metric")
    if emp_col:
        agg_df = df.groupby(emp_col)[metric_bar].mean().reset_index().sort_values(by=metric_bar, ascending=False)
        fig = px.bar(agg_df.head(10), x=emp_col, y=metric_bar, title=f"Top 10 Employees by {metric_bar.replace('_',' ')}",
                     color_continuous_scale="Viridis")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Employee column missing for this chart.")

with tab3:
    st.subheader("Correlation Heatmap")
    if len(metric_cols) > 1:
        corr = df[metric_cols].corr()
        fig = px.imshow(corr, text_auto=".2f", aspect="auto", title="Correlation Matrix of Metrics")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Not enough numeric fields for heatmap.")

with tab4:
    st.subheader("Box Plot by Employee")
    metric_box = st.selectbox("Select Metric", metric_cols, key="box_metric")
    if emp_col:
        fig = px.box(df, x=emp_col, y=metric_box, color=emp_col, points="all",
                     title=f"{metric_box.replace('_',' ')} Distribution by Employee", color_discrete_sequence=px.colors.qualitative.Bold)
        fig.update_layout(template="plotly_white", title_x=0.5, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Employee column missing for box plot.")

st.markdown("---")
with st.expander("View full raw KPI data"):
    st.dataframe(df, use_container_width=True, height=300)
