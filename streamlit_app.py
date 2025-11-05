import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

st.set_page_config(page_title="Employee KPI Analytics Suite", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
[data-testid="stSidebar"] > div { background: #204051; color: #fff }
.stApp, .main, .block-container { background: #f7f9fc; }
.big-metric {font-size: 2em;font-weight:700; padding:18px 0 6px 0;}
.stMetricLabel, .stMetricValue {color: #101820 !important;}
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
source_file = uploaded_file if uploaded_file else "Updated_18_KPI_Dashboard.xlsx"

time_col_candidates = ['Reporting_Period', 'Week_Ending_Date', 'Date', 'Quarter']

all_dfs = []
for sheet in employee_sheets:
    df = pd.read_excel(source_file, sheet_name=sheet)
    valid_time_cols = [col for col in time_col_candidates if col in df.columns]
    if valid_time_cols:
        df.rename(columns={valid_time_cols[0]: 'Reporting_Period'}, inplace=True)
    else:
        df['Reporting_Period'] = pd.NA
    df['KPI_Sheet'] = sheet
    all_dfs.append(df)

full_df = pd.concat(all_dfs, ignore_index=True)

time_col = 'Reporting_Period'
emp_col = 'Employee_ID'
metric_cols = [col for col in full_df.select_dtypes(include=np.number).columns if col not in [emp_col, time_col]]

periods = sorted(full_df[time_col].dropna().unique())
employees = sorted(full_df[emp_col].dropna().unique())

with st.container():
    st.title("Employee KPI Analytics Dashboard")
    filters = st.columns([2, 2])
    with filters[0]:
        selected_periods = st.multiselect('Choose Period', periods, default=periods)
    with filters[1]:
        selected_emps = st.multiselect('Choose Employees', employees, default=employees)

filtered_df = full_df[(full_df[time_col].isin(selected_periods)) & (full_df[emp_col].isin(selected_emps))]

st.subheader("Key Metrics Overview")
metric_columns = metric_cols[:5] if len(metric_cols) >= 5 else metric_cols
cols = st.columns(len(metric_columns))
for i, metric in enumerate(metric_columns):
    metric_mean = filtered_df[metric].mean() if not filtered_df.empty else 0
    metric_min = filtered_df[metric].min() if not filtered_df.empty else 0
    metric_max = filtered_df[metric].max() if not filtered_df.empty else 0
    with cols[i]:
        st.markdown(f"<div class='big-metric'>{metric_mean:.2f}</div>", unsafe_allow_html=True)
        st.markdown(f"**{metric.replace('_', ' ').title()}**")
        st.caption(f"Min: {metric_min:.1f} | Max: {metric_max:.1f}")

st.markdown("---")

st.subheader("Time-based Metric Trends")

for metric in metric_columns:
    st.markdown(f"### {metric.replace('_', ' ').title()}")
    fig = px.line(filtered_df, x=time_col, y=metric, color=emp_col, markers=True, title=f"{metric} over Time")
    fig.update_layout(template='plotly_white')
    st.plotly_chart(fig, use_container_width=True)

st.markdown("---")

st.subheader("Employee-wise Metric Comparison")

for metric in metric_columns:
    st.markdown(f"### {metric.replace('_', ' ').title()}")
    fig = px.box(filtered_df, x=emp_col, y=metric, color=emp_col, points='all', title=f"Distribution of {metric} by Employee")
    fig.update_layout(template='plotly_white', showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

st.markdown("---")
st.subheader("Raw KPI Data")
st.dataframe(filtered_df, use_container_width=True)
st.download_button("Download Filtered Data as CSV", filtered_df.to_csv(index=False), "filtered_employee_kpis.csv")
