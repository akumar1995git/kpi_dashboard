import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

st.set_page_config(page_title="Employee KPI Analytics", layout="wide", initial_sidebar_state="expanded")

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
    try:
        df = pd.read_excel(source_file, sheet_name=sheet)
        found_time = next((c for c in TIME_COLS if c in df.columns), None)
        found_emp = next((c for c in ['Employee_ID','Emp_ID','ID'] if c in df.columns), None)
        if not found_time:
            df['Reporting_Period'] = pd.NA
        else:
            df = df.rename(columns={found_time:'Reporting_Period'})
        if not found_emp:
            df['Employee_ID'] = pd.NA
        else:
            df = df.rename(columns={found_emp:'Employee_ID'})
        df['KPI_Sheet'] = sheet
        all_dfs.append(df)
    except Exception as e:
        st.warning(f"Failed to load sheet {sheet}: {e}")

if not all_dfs:
    st.error("No employee KPI data loaded. Check file upload or file contents.")
    st.stop()

full_df = pd.concat(all_dfs, ignore_index=True)
full_df['Reporting_Period'] = full_df['Reporting_Period'].astype(str)
full_df['Employee_ID'] = full_df['Employee_ID'].astype(str)
emp_col='Employee_ID'
time_col='Reporting_Period'
metric_cols = [col for col in full_df.select_dtypes(include=np.number).columns if col not in [emp_col]]

periods = sorted([p for p in full_df[time_col].dropna().unique() if str(p).lower()!='nan'])
employees = sorted([e for e in full_df[emp_col].dropna().unique() if str(e).lower()!='nan'])

with st.container():
    st.title("Employee KPI Analytics Dashboard")
    filters = st.columns([2,2])
    with filters[0]:
        selected_periods = st.multiselect('Choose Period', periods, default=periods if periods else [])
    with filters[1]:
        selected_emps = st.multiselect('Choose Employees', employees, default=employees if employees else [])

if not selected_periods or not selected_emps:
    st.warning("No periods or employees selected (or detected).")
    st.stop()

filtered_df = full_df[
    (full_df[time_col].isin(selected_periods)) &
    (full_df[emp_col].isin(selected_emps))
]

if filtered_df.empty:
    st.warning("No matching data after filtering. Please review your Excel source and selection filters.")
    st.stop()

# Summary Cards
metric_columns = metric_cols[:5]
cols = st.columns(len(metric_columns))
for i, metric in enumerate(metric_columns):
    mean = filtered_df[metric].mean()
    minv = filtered_df[metric].min()
    maxv = filtered_df[metric].max()
    with cols[i]:
        st.markdown(f"<div class='big-metric'>{mean:.2f}</div>", unsafe_allow_html=True)
        st.markdown(f"**{metric.replace('_',' ').title()}**")
        st.caption(f"Min: {minv:.1f} | Max: {maxv:.1f}")

st.markdown("---")

# Time trend chart per metric
for metric in metric_columns:
    st.markdown(f"### {metric.replace('_',' ').title()}")
    fig = px.line(filtered_df, x=time_col, y=metric, color=emp_col, markers=True, title=f"{metric} over Time")
    fig.update_layout(template='plotly_white')
    st.plotly_chart(fig, use_container_width=True)

st.markdown("---")

# Employee boxplot
for metric in metric_columns:
    st.markdown(f"### {metric.replace('_',' ').title()} by Employee")
    fig = px.box(filtered_df, x=emp_col, y=metric, color=emp_col, points='outliers', title=f"{metric} by Employee")
    fig.update_layout(template='plotly_white', showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

st.markdown("---")
st.subheader("Raw KPI Data")
st.dataframe(filtered_df, use_container_width=True)
st.download_button("Download Filtered Data as CSV", filtered_df.to_csv(index=False), "filtered_employee_kpis.csv")
