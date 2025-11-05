import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

st.set_page_config(page_title="Employee KPI Dashboard", layout="wide", initial_sidebar_state="expanded")

employee_kpi_sheets = [
    'Role_vs_Reality_Analysis',
    'Hidden_Capacity_Burnout_Risk',
    'Work_Models_Effectiveness',
    'Digital_Collaboration_Overload',
    'Digital_Wellbeing_Index',
    'High_Value_Work_Ratio',
    'Future_Skill_Readiness_Index'
]

TIME_COLS = ['Reporting_Period','Week_Ending_Date','Quarter','Date']

uploaded_file = st.sidebar.file_uploader("Upload Updated_18_KPI_Dashboard.xlsx", type='xlsx')
source_file = uploaded_file if uploaded_file else "Updated_18_KPI_Dashboard.xlsx"

all_dfs = []
for sheet in employee_kpi_sheets:
    df = pd.read_excel(source_file, sheet_name=sheet)
    # Standardize time (period) to 'Month'
    found_time = next((c for c in TIME_COLS if c in df.columns), None)
    if found_time:
        # Use only first 7 chars for YYYY-MM or if date, use .dt.to_period('M')
        if np.issubdtype(df[found_time].dtype, np.datetime64):
            df['Month'] = pd.to_datetime(df[found_time]).dt.to_period('M').astype(str)
        else:
            # For possible non-date/str cols
            df['Month'] = df[found_time].astype(str).str[:7]
    else:
        df['Month'] = pd.NA
    df['KPI_Sheet'] = sheet
    if 'Employee_ID' not in df:
        df['Employee_ID'] = pd.NA
    all_dfs.append(df)

full_df = pd.concat(all_dfs, ignore_index=True)
periods = sorted([p for p in full_df['Month'].dropna().unique() if p and p!='NaT'])
employees = sorted([e for e in full_df['Employee_ID'].dropna().unique() if e and e!='nan'])
metric_cols = [col for col in full_df.select_dtypes(include=np.number).columns if col not in ['Employee_ID']]

with st.container():
    st.title("Employee KPI Analytics Dashboard")
    filters = st.columns([2, 2])
    with filters[0]:
        selected_periods = st.multiselect('Choose Month', periods, default=periods)
    with filters[1]:
        selected_emps = st.multiselect('Choose Employees', employees, default=employees)

filtered_df = full_df[
    (full_df['Month'].isin(selected_periods)) & 
    (full_df['Employee_ID'].isin(selected_emps))
]

# --------------------------------------------
# Key metric summary & insight section
# --------------------------------------------
st.subheader("Key Metrics")
metric_columns = metric_cols[:5]
cols = st.columns(len(metric_columns))
for i, metric in enumerate(metric_columns):
    mean = filtered_df[metric].mean()
    prev = filtered_df.groupby('Month')[metric].mean().iloc[-2] if len(selected_periods) > 1 else None
    delta = mean - prev if prev is not None else 0
    trend_icon = "↑" if delta >=0 else "↓"
    with cols[i]:
        st.markdown(f"<div class='big-metric'>{mean:.2f}</div>", unsafe_allow_html=True)
        st.markdown(f"**{metric.replace('_',' ').title()}**")
        if prev is not None:
            st.caption(f"{trend_icon} {abs(delta):.2f} vs previous month")
    # Insight: simple trend
    if prev is not None:
        if delta > 0:
            insight = "This metric is increasing; monitor for opportunities or risk."
        elif delta < 0:
            insight = "This metric is decreasing; check recent changes or causes."
        else:
            insight = "No significant change from last month."
        st.caption(insight)

st.markdown("---")
# Trends
st.subheader("Monthly Trends")
for metric in metric_columns:
    st.markdown(f"### {metric.replace('_',' ').title()}")
    chart_df = filtered_df.groupby(['Month','Employee_ID'])[metric].mean().reset_index()
    fig = px.line(chart_df, x='Month', y=metric, color='Employee_ID', markers=True)
    st.plotly_chart(fig, use_container_width=True)

st.markdown("---")
st.subheader("Employee Comparison")
for metric in metric_columns:
    fig = px.box(filtered_df, x='Employee_ID', y=metric, color='Employee_ID', points='outliers')
    st.plotly_chart(fig, use_container_width=True)

st.markdown("---")
st.subheader("Raw KPI Data")
st.dataframe(filtered_df, use_container_width=True)
st.download_button("Download Filtered Data as CSV", filtered_df.to_csv(index=False), "filtered_employee_kpis.csv")
